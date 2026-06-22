

<# 
Prereqs:
  Install-Module Microsoft.Graph -Scope CurrentUser
  Install-Module Az.Accounts, Az.ResourceGraph, Az.OperationalInsights -Scope CurrentUser
  Connect-MgGraph -Scopes "Application.Read.All","Directory.Read.All","Policy.Read.All","AuditLog.Read.All"
  Connect-AzAccount
Edit the PARAMS block and run.
#>

#=============== PARAMS ===============#
$TenantId         = (Get-MgContext).TenantId
$SubscriptionId   = "<SUBSCRIPTION_ID>"          # Needed for ARG / Log Analytics
$ResourceGroup    = "<LOGANALYTICS_RG>"          # Only if using LA for inactivity
$WorkspaceName    = "<LOGANALYTICS_WS_NAME>"     # Only if using LA for inactivity
$InactiveDays     = 90                           # Abandoned threshold
$HighPrivScopes   = @(
  # Graph delegated / app permissions commonly abused or overly broad:
  "Mail.Read","Mail.ReadWrite","Mail.ReadBasic","Mail.ReadBasic.All",
  "Mail.ReadWrite","Mail.ReadWrite.All","MailboxSettings.Read","MailboxSettings.ReadWrite",
  "Files.Read","Files.Read.All","Files.ReadWrite","Files.ReadWrite.All",
  "Directory.ReadWrite.All","Directory.AccessAsUser.All",
  "RoleManagement.ReadWrite.Directory","AppRoleAssignment.ReadWrite.All",
  "Reports.Read.All","SecurityEvents.Read.All","AuditLog.Read.All"
)
$HighPrivAppRoles = @(
  # App roles (application permissions) by value match text (works across many APIs)
  "Mail.Read","Mail.ReadWrite","Files.Read.All","Files.ReadWrite.All",
  "Directory.ReadWrite.All","RoleManagement.ReadWrite.Directory","AppRoleAssignment.ReadWrite.All"
)
$OutputDir        = Join-Path $PWD "EntraAppAudit-$(Get-Date -Format yyyyMMdd-HHmm)"
New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

#=============== HELPERS ===============#
function Invoke-WithRetry {
  param([ScriptBlock]$Script, [int]$Retries=6)
  for($i=1;$i -le $Retries;$i++){
    try { return & $Script } catch {
      if($i -eq $Retries){ throw }
      #Start-Sleep -Seconds ([Math]::Min(2**$i, 30))
    }
  }
}

function Get-AllPages {
  param([ScriptBlock]$FirstCall, [ScriptBlock]$NextCall, [string]$ItemName="Value")
  $all = @()
  $res = Invoke-WithRetry $FirstCall
  if ($res.$ItemName) { $all += $res.$ItemName }
  $next = $res.'@odata.nextLink'
  while ($next) {
    $res = Invoke-WithRetry { & $NextCall.Invoke($next) }
    if ($res.$ItemName) { $all += $res.$ItemName }
    $next = $res.'@odata.nextLink'
  }
  return $all
}

#=============== 1) ENTERPRISE APPS (SERVICE PRINCIPALS) ===============#
Write-Host "Pulling all Service Principals…" -ForegroundColor Cyan
$spFields = "id,appId,displayName,accountEnabled,appOwnerOrganizationId,appRoles,servicePrincipalType,verificationStatus,tags,createdDateTime"
$spAll = Get-AllPages `
  -FirstCall { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$top=999&`$select=$spFields" } `
  -NextCall  { param($url) Invoke-MgGraphRequest -Method GET -Uri $url }

# Owners (batched)
Write-Host "Resolving owners (this can take a while for 20k)..." -ForegroundColor Cyan
$ownerMap = @{}
foreach ($chunk in $spAll.Id | ForEach-Object -Begin {$i=0} -Process {
  $_; $i++; if($i%200 -eq 0) {','}
} -split ',' | Where-Object {$_}) {
  foreach($spId in $chunk){
    $owners = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$spId/owners?`$select=id,displayName,userPrincipalName,mail" }
    $ownerMap[$spId] = @($owners.value | ForEach-Object {
      $_.displayName ?? $_.userPrincipalName ?? $_.mail ?? $_.id
    })
  }
}

# App Role Assignments (application permissions granted to client SPs)
Write-Host "Pulling app role assignments…" -ForegroundColor Cyan
# For each SP as a client: /servicePrincipals/{id}/appRoleAssignments gets assignments the SP (client) has to resource apps
$appRoleAssignMap = @{}
foreach ($sp in $spAll) {
  $assigns = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments?`$top=999" }
  $appRoleAssignMap[$sp.id] = @($assigns.value)
}

# Delegated OAuth Grants (user or admin consented delegated permissions)
Write-Host "Pulling delegated permission grants…" -ForegroundColor Cyan
$oauthGrantMap = @{}
foreach ($sp in $spAll) {
  # clientId equals the client service principal id
  $grants = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=clientId eq '$($sp.id)'" }
  $oauthGrantMap[$sp.id] = @($grants.value)
}

# Credentials hygiene (secrets & certs)
Write-Host "Pulling credentials for SPs…" -ForegroundColor Cyan
$credMap = @{}
foreach ($sp in $spAll) {
  $spCreds = Invoke-WithRetry { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)?`$select=id,displayName,passwordCredentials,keyCredentials" }
  $credMap[$sp.id] = @{
    passwordCredentials = @($spCreds.passwordCredentials)
    keyCredentials      = @($spCreds.keyCredentials)
  }
}

#=============== 2) APP REGISTRATIONS (APPLICATIONS) ===============#
Write-Host "Pulling App Registrations (applications) …" -ForegroundColor Cyan
$appFields = "id,appId,displayName,signInAudience,requiredResourceAccess,createdDateTime,owners"
$appAll = Get-AllPages `
  -FirstCall { Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications?`$top=999&`$select=$appFields" } `
  -NextCall  { param($url) Invoke-MgGraphRequest -Method GET -Uri $url }

#=============== 3) MANAGED IDENTITIES (ARG + Graph) ===============#
Write-Host "Querying Managed Identities via Azure Resource Graph…" -ForegroundColor Cyan
Select-AzSubscription -SubscriptionId $SubscriptionId | Out-Null

$queryUAMI = @"
resources
| where type =~ 'microsoft.managedidentity/userAssignedIdentities'
| project id, name, location, tenantId = tostring(properties.tenantId), principalId = tostring(properties.principalId), clientId = tostring(properties.clientId), resourceGroup, subscriptionId
"@

$uami = Search-AzGraph -Query $queryUAMI -First 5000

# What resources are using those UAMIs?
$queryUsers = @"
resources
| where isnotempty(identity.userAssignedIdentities)
| mv-expand identity_user = bag_keys(identity.userAssignedIdentities)
| project consumerId = id, consumerType = type, uamiId = tostring(identity_user), subscriptionId
"@
$uamiConsumers = Search-AzGraph -Query $queryUsers -First 50000

# System-assigned identities coverage (just counts and which resources have them)
$querySAMI = @"
resources
| where isnotempty(identity.type) and identity.type =~ 'SystemAssigned'
| project id, type, name, identityPrincipalId = tostring(identity.principalId), identityTenantId = tostring(identity.tenantId)
"@
$sami = Search-AzGraph -Query $querySAMI -First 50000

#=============== 4) LAST SIGN-IN (OPTIONAL, from Log Analytics) ===============#
$lastSignInMap = @{}
if ($WorkspaceName -and $ResourceGroup) {
  Write-Host "Pulling last Service Principal sign-ins from Log Analytics…" -ForegroundColor Cyan
  $ws = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroup -Name $WorkspaceName
  $kql = @"
ServicePrincipalSignInLogs
| summarize LastSignIn=max(TimeGenerated) by ServicePrincipalId
"@
  $res = Invoke-AzOperationalInsightsQuery -WorkspaceId $ws.CustomerId -Query $kql -Timespan (New-TimeSpan -Days ([Math]::Max($InactiveDays,120)))
  if ($res.Tables.Count -gt 0) {
    foreach ($row in $res.Tables[0].Rows) {
      $spId = $row[0]; $last = [datetime]$row[1]
      $lastSignInMap[$spId] = $last
    }
  }
}

#=============== 5) CLASSIFY / FLAG RISKS ===============#
Write-Host "Classifying risks…" -ForegroundColor Cyan
$now = Get-Date
$inactiveCutoff = $now.AddDays(-$InactiveDays)

$spReport = foreach ($sp in $spAll) {
  $owners = $ownerMap[$sp.id] | Where-Object { $_ } 
  $ownerless = @($owners).Count -eq 0

  $appRoleAssigns = $appRoleAssignMap[$sp.id]
  $oauthGrants    = $oauthGrantMap[$sp.id]
  $creds          = $credMap[$sp.id]

  # high-priv via application permissions (app role assignments)
  $appTooPermissive = $false
  foreach ($a in $appRoleAssigns) {
    $roleName = $a.appRoleId   # this is GUID; we also have resourceDisplayName and sometimes app role display names in expanded calls
    $resourceName = $a.resourceDisplayName
    $displayHint = "$($a.PrincipalDisplayName) -> $resourceName"
    # best-effort: also check assignment 'appRoleId' name requires expansion; fallback to resourceDisplayName heuristic
    if ($resourceName -match "Graph|SharePoint|Exchange|Office|Microsoft") { 
      # Mark permissive by resource and we’ll enrich separately if needed
      $appTooPermissive = $true
    }
  }

  # high-priv via delegated grants (scope text available)
  $delegatedTooPermissive = $false
  foreach ($g in $oauthGrants) {
    $scopeText = ($g.scope ?? "").Split(" ")
    if (@($scopeText | Where-Object { $_ -in $HighPrivScopes }).Count -gt 0) { $delegatedTooPermissive = $true }
  }

  $lastSignIn = $null
  if ($lastSignInMap.ContainsKey($sp.id)) { $lastSignIn = $lastSignInMap[$sp.id] }

  # Abandoned criteria (tune as needed):
  # - No owners
  # - AND no sign-in in N days
  # - AND no app role assignments and no delegated grants
  $hasAssignmentsOrGrants = (@($appRoleAssigns).Count + @($oauthGrants).Count) -gt 0
  $inactive = $lastSignIn -lt $inactiveCutoff -or $null -eq $lastSignIn
  $abandoned = $ownerless -and $inactive -and -not $hasAssignmentsOrGrants

  # Credential hygiene
  $secretCount = @($creds.passwordCredentials).Count
  $certCount   = @($creds.keyCredentials).Count
  $hasSecrets  = $secretCount -gt 0
  $expiredSecrets = @($creds.passwordCredentials | Where-Object { $_.endDateTime -lt $now }).Count
  $longLivedSecrets = @($creds.passwordCredentials | Where-Object { $_.endDateTime -gt $now.AddYears(1) }).Count

  [pscustomobject]@{
    SP_ObjectId          = $sp.id
    AppId                = $sp.appId
    DisplayName          = $sp.displayName
    Type                 = $sp.servicePrincipalType
    VerifiedPublisher    = $sp.verificationStatus
    Owners               = ($owners -join "; ")
    Ownerless            = $ownerless
    LastSignInUtc        = $lastSignIn
    InactiveDays         = if ($lastSignIn) { [int]($now - $lastSignIn).TotalDays } else { $null }
    HasAssignmentsOrGrants = $hasAssignmentsOrGrants
    TooPermissive_App    = $appTooPermissive
    TooPermissive_Delegated = $delegatedTooPermissive
    HasSecrets           = $hasSecrets
    SecretCount          = $secretCount
    SecretExpiredCount   = $expiredSecrets
    SecretLongLivedCount = $longLivedSecrets
    CertCount            = $certCount
    Abandoned            = $abandoned
  }
}

#=============== 6) MANAGED IDENTITIES REPORTS ===============#
# Join UAMI with consumers
$uamiByConsumers = $uamiConsumers | Group-Object -Property uamiId | ForEach-Object {
  [pscustomobject]@{
    UamiId       = $_.Name
    ConsumerCount= ($_.Group | Select-Object -ExpandProperty consumerId -Unique | Measure-Object).Count
    Consumers    = ($_.Group | Select-Object -ExpandProperty consumerId -Unique) -join ";"
  }
}
$uamiReport = $uami | ForEach-Object {
  $rel = $uamiByConsumers | Where-Object { $_.UamiId -eq $_.id }
  [pscustomobject]@{
    UAMI_Id        = $_.id
    Name           = $_.name
    Location       = $_.location
    PrincipalId    = $_.principalId
    ClientId       = $_.clientId
    ResourceGroup  = $_.resourceGroup
    SubscriptionId = $_.subscriptionId
    ConsumerCount  = $rel.ConsumerCount
    Consumers      = $rel.Consumers
  }
}

# System-assigned identities list (cannot be "attached", they are intrinsic to resources)
$samiReport = $sami | ForEach-Object {
  [pscustomobject]@{
    ResourceId   = $_.id
    ResourceType = $_.type
    ResourceName = $_.name
    PrincipalId  = $_.identityPrincipalId
    TenantId     = $_.identityTenantId
  }
}

#=============== 7) EXPORTS & SUMMARIES ===============#
$spCsv       = Join-Path $OutputDir "ServicePrincipals_Audit.csv"
$uamiCsv     = Join-Path $OutputDir "ManagedIdentities_UserAssigned.csv"
$samiCsv     = Join-Path $OutputDir "ManagedIdentities_SystemAssigned.csv"
$summaryCsv  = Join-Path $OutputDir "Summary_Metrics.csv"

$spReport | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $spCsv
$uamiReport | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $uamiCsv
$samiReport | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $samiCsv

# Quick metrics
$metrics = [pscustomobject]@{
  TotalServicePrincipals     = ($spAll | Measure-Object).Count
  Ownerless                  = ($spReport | Where-Object Ownerless | Measure-Object).Count
  Abandoned                  = ($spReport | Where-Object Abandoned | Measure-Object).Count
  TooPermissive_Any          = ($spReport | Where-Object { $_.TooPermissive_App -or $_.TooPermissive_Delegated } | Measure-Object).Count
  WithSecrets                = ($spReport | Where-Object HasSecrets | Measure-Object).Count
  ExpiredSecrets             = ($spReport | Where-Object { $_.SecretExpiredCount -gt 0 } | Measure-Object).Count
  LongLivedSecretsGT1Y       = ($spReport | Where-Object { $_.SecretLongLivedCount -gt 0 } | Measure-Object).Count
  UAMI_Count                 = ($uami | Measure-Object).Count
  UAMI_WithConsumers         = ($uamiReport | Where-Object { $_.ConsumerCount -gt 0 } | Measure-Object).Count
  SAMI_Count                 = ($sami | Measure-Object).Count
}
$metrics | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $summaryCsv

Write-Host "`n=== OUTPUTS ===" -ForegroundColor Green
Write-Host "Service Principals Audit: $spCsv"
Write-Host "UAMI Report:              $uamiCsv"
Write-Host "SAMI Report:              $samiCsv"
Write-Host "Summary Metrics:          $summaryCsv"
