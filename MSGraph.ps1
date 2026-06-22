param(
  [int]$LookbackDays = 30,
  [int]$ExpiryWarnDays = 45,
  [string]$OutputCsv = ".\AppRegistration_SecurityFindings.csv"
)

$ErrorActionPreference = "Stop"


# --------------------------
# Risky permission allowlists
# --------------------------
# Application permissions that are commonly over-privileged
$riskyAppPerms = @(
  "Directory.AccessAsUser.All", "Directory.ReadWrite.All", "Directory.Read.All", "Directory.ReadWrite.All",
  "User.ReadWrite.All", "User.Read.All",
  "Group.ReadWrite.All", "Group.Read.All",
  "AuditLog.Read.All",
  "Policy.ReadWrite.AuthenticationMethod", "Policy.ReadWrite.TrustFramework",
  "Reports.Read.All",
  "Mail.ReadWrite", "Mail.Read", "MailboxSettings.ReadWrite",
  "Files.ReadWrite.All", "Files.Read.All",
  "Sites.FullControl.All", "Sites.ReadWrite.All",
  "Device.ReadWrite.All",
  "PrivilegedAccess.ReadWrite.AzureAD", "PrivilegedAccess.ReadWrite.AzureResources",
  "RoleManagement.ReadWrite.Directory"
) | Select-Object -Unique

# Delegated permissions worth flagging if broadly consented
$riskyDelegatedPerms = @(
  "Directory.ReadWrite.All", "User.ReadWrite.All", "Group.ReadWrite.All",
  "Mail.ReadWrite", "Files.ReadWrite", "Files.ReadWrite.All",
  "Sites.FullControl.All", "Sites.ReadWrite.All"
) | Select-Object -Unique

# --------------------------
# Connect (if not already)
# --------------------------
#if (-not (Get-MgContext)) {
#  Connect-MgGraph -Scopes "Application.Read.All",,"AuditLog.Read.All","Directory.Read.All","Policy.Read.All","RoleManagement.Read.Directory"
#}
#Select-MgProfile -Name "beta" | Out-Null  # SignIn logs for SPs more reliable in beta

# --------------------------
# Tenant-level consent posture
# --------------------------
#$authzPolicy   = Try-Graph { Get-MgPolicyAuthorizationPolicy -AuthorizationPolicyId "authorizationPolicy" } "AuthorizationPolicy"
#$grantPolicies = Try-Graph { Get-MgPolicyPermissionGrantPolicy -All } "PermissionGrantPolicies"

#$tenantConsentSummary = if ($authzPolicy -and $grantPolicies) {
#  # Crude read: if default policy is very restrictive, user consent is likely limited.
#  [PSCustomObject]@{
#    AuthorizationPolicyId = $authzPolicy.Id
#    DefaultUserRole_AllowedToGrantAppPerms = $authzPolicy.DefaultUserRolePermissions.AllowedToCreateApps -ne $true -and $false # legacy; informational only
#    GrantPolicies = ($grantPolicies.DisplayName -join "; ")
#  }
#} else {
#  [PSCustomObject]@{
#    AuthorizationPolicyId = "NotAvailable"
#    DefaultUserRole_AllowedToGrantAppPerms = "NotAvailable"
#    GrantPolicies = "NotAvailable"
#  }
#}

# --------------------------
# Pull Applications & SPs
# --------------------------
$applications = Try-Graph { Get-MgApplication -All -Property "id,appId,displayName,signInAudience,requiredResourceAccess,owners,keyCredentials,passwordCredentials" } "Applications"
$servicePrincipals = Try-Graph { Get-MgServicePrincipal -All -Property "id,appId,displayName,signInAudience,owners,appRoleAssignments,AppRoleAssignmentRequired" } "ServicePrincipals"

if (-not $applications) { throw "Could not retrieve applications." }
if (-not $servicePrincipals) { throw "Could not retrieve service principals." }

# Index SP by appId
$spByAppId = @{}
foreach ($sp in $servicePrincipals) {
  if ($sp.AppId) { $spByAppId[$sp.AppId] = $sp }
}

# --------------------------
# Sign-in logs for SPs
# --------------------------
$since = (Get-Date).AddDays(-$LookbackDays)
$spSignIns = Try-Graph {
  Get-MgAuditLogSignIn -All -Filter "createdDateTime ge $($since.ToString("o")) and servicePrincipalId ne null" `
    -Property "appDisplayName,servicePrincipalId,createdDateTime,ipAddress,status"
} "ServicePrincipalSignInLogs"

# Group by SP Id to detect activity
$spLastSeen = @{}
if ($spSignIns) {
  $spSignIns | Sort-Object -Property createdDateTime -Descending | ForEach-Object {
    $id = $_.ServicePrincipalId
    if ($id -and -not $spLastSeen.ContainsKey($id)) { $spLastSeen[$id] = $_.createdDateTime }
  }
}

# --------------------------
# Directory role assignments to SPs (elevated)
# --------------------------
$roleAssignments = Try-Graph { Get-MgRoleManagementDirectoryRoleAssignment -All -Property "id,principalId,roleDefinitionId,principal,roleDefinition,resourceScope" } "DirectoryRoleAssignments"
# Build a quick map: principalId => roles
$rolesBySpId = @{}
if ($roleAssignments) {
  foreach ($ra in $roleAssignments) {
    $p = $ra.PrincipalId
    if ($p) {
      if (-not $rolesBySpId.ContainsKey($p)) { $rolesBySpId[$p] = @() }
      $rolesBySpId[$p] += $ra.RoleDefinition.DisplayName
    }
  }
}

# --------------------------
# Helper: Resolve permissions from requiredResourceAccess
# --------------------------
function Get-AppRequiredPermissions {
  param($app)
  $appPerms = @()
  $dlgPerms = @()

  if ($app.requiredResourceAccess) {
    foreach ($r in $app.requiredResourceAccess) {
      # r.resourceAppId (e.g., Graph appId)
      $resourceAppId = $r.resourceAppId
      foreach ($ra in $r.resourceAccess) {
        # ra.type: Scope (delegated) or Role (application)
        $entry = [PSCustomObject]@{
          ResourceAppId = $resourceAppId
          AccessType    = $ra.type
          Id            = $ra.id
        }
        if ($ra.type -eq "Role") { $appPerms += $entry } else { $dlgPerms += $entry }
      }
    }
  }

  return [PSCustomObject]@{
    ApplicationPerms = $appPerms
    DelegatedPerms   = $dlgPerms
  }
}

# --------------------------
# Helper: Severity scoring
# --------------------------
#function Rate-Severity {
#  param(
#    [int]$critical = 0,
#   [int]$high = 0,
#    [int]$medium = 0,
#    [int]$low = 0
#  )
#  if ($critical -gt 0) { return "Critical" }
#  elseif ($high -gt 0) { return "High" }
#  elseif ($medium -gt 0) { return "Medium" }
#  elseif ($low -gt 0) { return "Low" }
#  else { return "Info" }
#}

# --------------------------
# Iterate apps and evaluate
# --------------------------
$now = Get-Date
$warnDate = $now.AddDays($ExpiryWarnDays)

$results = New-Object System.Collections.Generic.List[object]

foreach ($app in $applications) {
  $sp = $null
  if ($app.AppId -and $spByAppId.ContainsKey($app.AppId)) { $sp = $spByAppId[$app.AppId] }

  # Ownership
  $owners = Try-Graph { Get-MgApplicationOwner -ApplicationId $app.Id -All } "AppOwners"
  $ownerCount = if ($owners) { ($owners | Measure-Object).Count } else { 0 }
  $ownerFindings = @()
  if ($owners -eq $null) { $ownerFindings += "Owners: NotAvailable" }
  elseif ($ownerCount -eq 0) { $ownerFindings += "No owners" }
  elseif ($ownerCount -gt 5) { $ownerFindings += "Too many owners ($ownerCount)" }

  # Secrets/Certs
  $secretFindings = @()
  foreach ($pw in ($app.passwordCredentials | ForEach-Object { $_ })) {
    $end = [datetime]$pw.endDateTime
    if ($end -lt $now) { $secretFindings += "Expired secret ($($end.ToString('u')))" }
    elseif ($end -lt $warnDate) { $secretFindings += "Secret expiring soon ($($end.ToString('u')))" }
    if (($end - $now).TotalDays -gt 365) { $secretFindings += "Long-lived secret (>365d)" }
  }
  foreach ($kc in ($app.keyCredentials | ForEach-Object { $_ })) {
    $end = [datetime]$kc.endDateTime
    if ($end -lt $now) { $secretFindings += "Expired certificate ($($end.ToString('u')))" }
    elseif ($end -lt $warnDate) { $secretFindings += "Certificate expiring soon ($($end.ToString('u')))" }
    if (($end - $now).TotalDays -gt 730) { $secretFindings += "Long-lived certificate (>730d)" }
  }

  # Multi-tenant
  $aud = $app.signInAudience
  $isMultiTenant = $aud -in @("AzureADMultipleOrgs","AzureADandPersonalMicrosoftAccount")

  # Permissions requested (requiredResourceAccess)
  $perm = Get-AppRequiredPermissions -app $app
  $appPermIds = $perm.ApplicationPerms
  $dlgPermIds = $perm.DelegatedPerms

  # Best-effort permission names: Graph name resolution would require pulling the service principal for the resource and mapping IDs.
  # For now, we will flag by AccessType and count, and mark "risky" if resource is Graph and later resolve names if available.
  # Attempt name resolution for Microsoft Graph only:
  $graphSp = $servicePrincipals | Where-Object { $_.AppId -eq "00000003-0000-0000-c000-000000000000" } # Graph
  $graphAppRoles = @{}
  $graphScopes   = @{}
  if ($graphSp) {
    foreach ($ar in $graphSp.AppRoles) { $graphAppRoles[$ar.Id] = $ar.Value }
    foreach ($sc in $graphSp.Oauth2PermissionScopes) { $graphScopes[$sc.Id] = $sc.Value }
  }
  $resolvedAppPerms = @()
  foreach ($p in $appPermIds) {
    if ($p.ResourceAppId -eq "00000003-0000-0000-c000-000000000000" -and $graphAppRoles.ContainsKey($p.Id)) {
      $resolvedAppPerms += $graphAppRoles[$p.Id]
    }
  }
  $resolvedDlgPerms = @()
  foreach ($p in $dlgPermIds) {
    if ($p.ResourceAppId -eq "00000003-0000-0000-c000-000000000000" -and $graphScopes.ContainsKey($p.Id)) {
      $resolvedDlgPerms += $graphScopes[$p.Id]
    }
  }

  $riskyAppHits = @()
  foreach ($n in $resolvedAppPerms) {
    if ($riskyAppPerms -contains $n) { $riskyAppHits += $n }
  }
  $riskyDlgHits = @()
  foreach ($n in $resolvedDlgPerms) {
    if ($riskyDelegatedPerms -contains $n) { $riskyDlgHits += $n }
  }

  # Sign-in last seen
  $spLast = $null
  if ($sp -and $spLastSeen.ContainsKey($sp.Id)) { $spLast = $spLastSeen[$sp.Id] }
  $staleApp = $true
  if ($spLast) { $staleApp = ([datetime]$spLast) -lt $since } else { $staleApp = $true }

  # Directory roles assigned to SP (elevated)
  $spElevatedRoles = @()
  if ($sp -and $rolesBySpId.ContainsKey($sp.Id)) { $spElevatedRoles = ($rolesBySpId[$sp.Id] | Select-Object -Unique) }

  # App role assignments breadth
  $appRoleAssignments = Try-Graph { Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All } "SPAppRoleAssignments"
  $assignmentCount = if ($appRoleAssignments) { $appRoleAssignments.Count } else { 0 }

  # Build findings + severity
  $critical = 0; $high = 0; $medium = 0; $low = 0
  $findings = @()

  # Secrets/Certs
  if ($secretFindings -match "Expired") { $critical++ ; $findings += ($secretFindings | Where-Object {$_ -match "Expired"}) }
  if ($secretFindings -match "expiring soon") { $high++ ; $findings += ($secretFindings | Where-Object {$_ -match "expiring soon"}) }
  if ($secretFindings -match "Long-lived secret") { $medium++ ; $findings += ($secretFindings | Where-Object {$_ -match "Long-lived secret"}) }
  if ($secretFindings -match "Long-lived certificate") { $medium++ ; $findings += ($secretFindings | Where-Object {$_ -match "Long-lived certificate"}) }

  # Owners
  if ($ownerFindings -contains "No owners") { $high++; $findings += "No owners" }
  elseif ($ownerFindings -match "Too many owners") { $medium++; $findings += $ownerFindings }

  if ($owners -eq $null) { $low++; $findings += "Owners: NotAvailable" }

  # Multi-tenant
  if ($isMultiTenant) { $high++; $findings += "Multi-tenant app (signInAudience=$aud)" }

  # Permissions
  if ($resolvedAppPerms.Count -gt 0 -and $riskyAppHits.Count -gt 0) { $critical++; $findings += ("Risky Application perms: " + ($riskyAppHits -join ", ")) }
  if ($resolvedDlgPerms.Count -gt 0 -and $riskyDlgHits.Count -gt 0) { $high++; $findings += ("Risky Delegated perms: " + ($riskyDlgHits -join ", ")) }

  # Stale
  if ($staleApp) { $medium++; $findings += "No service principal sign-ins in last $LookbackDays days" }

  # Elevated directory roles
  if ($spElevatedRoles.Count -gt 0) { $critical++; $findings += ("Service principal has elevated directory role(s): " + ($spElevatedRoles -join ", ")) }

  # Broad assignments
  if ($assignmentCount -gt 50) { $medium++; $findings += "Large number of app role assignments (>50) — review exposure" }

  $severity = Rate-Severity -critical $critical -high $high -medium $medium -low $low

  $results.Add([PSCustomObject]@{
    AppDisplayName       = $app.DisplayName
    AppObjectId          = $app.Id
    AppId                = $app.AppId
    SPDisplayName        = $sp.DisplayName
    SPObjectId           = $sp.Id
    SignInAudience       = $aud
    OwnersCount          = $ownerCount
    RiskyAppPerms        = if ($riskyAppHits) { $riskyAppHits -join "; " } else { "" }
    RiskyDelegatedPerms  = if ($riskyDlgHits) { $riskyDlgHits -join "; " } else { "" }
    LastSpSignIn         = if ($spLast) { [datetime]$spLast } else { $null }
    ElevatedDirRoles     = if ($spElevatedRoles) { $spElevatedRoles -join "; " } else { "" }
    AppRoleAssignments   = $assignmentCount
    Findings             = if ($findings) { $findings -join " | " } else { "None" }
    Severity             = $severity
  })
}

# --------------------------
# Output
# --------------------------
$results | Sort-Object { @{"Critical"=1;"High"=2;"Medium"=3;"Low"=4;"Info"=5}[$_.Severity] }, AppDisplayName | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutputCsv
Write-Host "Exported findings to: $OutputCsv"

# Summary table
$results | Group-Object Severity | Select-Object Name,Count | Sort-Object Name | Format-Table -AutoSize

# Top risky examples
Write-Host "`nTop apps with Critical findings:" -ForegroundColor Yellow
$results | Where-Object Severity -eq "Critical" | Select-Object AppDisplayName, RiskyAppPerms, ElevatedDirRoles, LastSpSignIn, Findings | Format-Table -AutoSize

