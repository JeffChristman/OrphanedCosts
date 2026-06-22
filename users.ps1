

<# 
.SYNOPSIS
  VA Entra ID Users & Groups Configuration Audit + AD Connect posture checks (config-level; no per-user deep dive).

.DESCRIPTION
  Focuses on tenant-wide *configuration and governance* for Users/Groups and Entra ID <-> AD Connect posture, not individual account triage.
  Produces a concise JSON + CSV summary of security gaps, red flags, and best-practice misalignments.

.REQUIREMENTS
  - PowerShell 7+ recommended
  - Microsoft Graph PowerShell SDK:
        Install-Module Microsoft.Graph -Scope CurrentUser
  - Permissions for Connect-MgGraph:
        Policy.Read.All, Directory.Read.All, Group.Read.All, AuditLog.Read.All, RoleManagement.Read.Directory
  - (Optional, on AAD Connect server) ADSync module for local connector checks

.USAGE
  Run with sufficient Graph permissions:
    pwsh -File .\Entra_UserGroup_Config_Audit.ps1

  Optional fast mode to avoid large enumerations:
    pwsh -File .\Entra_UserGroup_Config_Audit.ps1 -Lightweight

.OUTPUT
  - ./output/Entra_UserGroup_Config_Audit.json
  - ./output/Entra_UserGroup_Config_Audit.csv
  - Console summary table
#>

param(
  [switch]$Lightweight
)

# ---------- Helpers ----------
#function Ensure-Module {
#  param([string]$Name)
#  if (-not (Get-Module -ListAvailable -Name $Name)) {
#   Write-Host "Installing module: $Name ..." -ForegroundColor Yellow
#   Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
#  }
#  Import-Module $Name -ErrorAction Stop
#}

function Connect-GraphSafe {
  $scopes = @(
    'Directory.Read.All',
    'Group.Read.All',
    'Policy.Read.All',
    'AuditLog.Read.All',
    'RoleManagement.Read.Directory'
  )
  try {
    if (-not (Get-MgContext)) {
      Connect-MgGraph -Scopes $scopes -NoWelcome
    }
    Select-MgProfile -Name "beta"  # beta needed for some policy endpoints; safe fallback to v1 for others
  } catch {
    Write-Warning "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    throw
  }
}

function New-OutputFolder {
  $out = Join-Path -Path (Get-Location) -ChildPath "output"
  if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
  return $out
}

# ---------- Begin ----------
try {
  Ensure-Module -Name Microsoft.Graph
  Connect-GraphSafe
} catch {
  Write-Error "Cannot continue without Microsoft Graph connectivity. $_"
  exit 1
}

$results = [System.Collections.Generic.List[hashtable]]::new()

# ---------- Tenant & Policy Baselines ----------
Write-Host "`n[1/7] Reading tenant authorization & security defaults..." -ForegroundColor Cyan
$authz = $null; $secDefaults = $null; $consentPolicies = @()

try { $authz = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop } catch { Write-Warning "AuthZ policy read failed: $_" }
try { $secDefaults = Get-MgPolicyIdentitySecurityDefaultsEnforcementPolicy -ErrorAction Stop } catch { Write-Warning "Security defaults read failed: $_" }

# User consent / app consent posture (policy-based + legacy defaults)
try { $pgPolicies = Get-MgPolicyPermissionGrantPolicy -All -ErrorAction Stop } catch { $pgPolicies = @(); Write-Warning "Permission grant policies read failed: $_" }

# ---------- Conditional Access (MFA/Legacy Auth posture) ----------
Write-Host "[2/7] Reading Conditional Access policies..." -ForegroundColor Cyan
$caPolicies = @()
try { $caPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop } catch { Write-Warning "CA policy read failed: $_" }

# Heuristics: detect presence of MFA-for-all/Admins, and Legacy Auth block
$hasMfaAll = $false
$hasMfaAdmins = $false
$blocksLegacy = $false

foreach ($p in $caPolicies) {
  $grant = $p.grantControls
  $conditions = $p.conditions
  $apps = $conditions.applications
  $clientApps = $conditions.clientAppTypes

  $requiresMfa = ($grant) -and ($grant.builtInControls -contains "mfa")
  $targetsAllUsers = ($conditions.users.includeUsers -contains "All") -or ($conditions.users.includeRoles -contains "62e90394-69f5-4237-9190-012177145e10") # Global Admin role id

  if ($requiresMfa -and $targetsAllUsers) { $hasMfaAll = $true }
  if ($requiresMfa -and ($conditions.users.includeRoles -and $conditions.users.includeRoles.Count -gt 0)) { $hasMfaAdmins = $true }

  # Legacy auth block heuristic: policy includes client app "legacyAuthenticationClients" OR conditions include Exchange ActiveSync/other
  if ($clientApps -and ($clientApps -contains "legacyAuthenticationClients" -or $clientApps -contains "exchangeActiveSync")) {
    if ($p.state -eq "enabled" -or $p.state -eq "enabledForReportingButNotEnforced") {
      if ($grant -and $grant.builtInControls -contains "block") { $blocksLegacy = $true }
    }
  }
}

# ---------- User & Guest Policies ----------
Write-Host "[3/7] Reading user & guest configuration..." -ForegroundColor Cyan
$defaultUserPermissions = $null
try { $defaultUserPermissions = $authz.defaultUserRolePermissions } catch {}

$guestConfig = @{
  AllowInvitesFrom = $authz.AllowInvitesFrom
  GuestUserRoleId  = $authz.GuestUserRoleId
}
# Sanity: important booleans if present
$userFlags = [ordered]@{
  UsersCanRegisterApps            = $defaultUserPermissions.AllowedToCreateApps
  UsersCanCreateSecurityGroups    = $defaultUserPermissions.AllowedToCreateSecurityGroups
  UsersCanReadOtherUsers          = $defaultUserPermissions.AllowedToReadOtherUsers
  PermissionGrantPoliciesAssigned = ($defaultUserPermissions.PermissionGrantPoliciesAssigned -join ", ")
}

# ---------- Group Settings (Unified/M365 + Security) ----------
Write-Host "[4/7] Reading directory settings for Groups..." -ForegroundColor Cyan
$dirSettings = @()
try { $dirSettings = Get-MgDirectorySetting -All -ErrorAction Stop } catch { Write-Warning "Directory settings read failed: $_" }

$groupUnified = $dirSettings | Where-Object { $_.DisplayName -like "*Group.Unified*" }
$groupGuest   = $dirSettings | Where-Object { $_.DisplayName -like "*Group.Unified.Guest*" }

function Get-SettingValue {
  param($setting, $key)
  if (-not $setting) { return $null }
  $item = $setting.Values | Where-Object { $_.Name -eq $key }
  return $item.Value
}

$groupCfg = [ordered]@{
  EnableMIPLabels                    = (Get-SettingValue $groupUnified "EnableMIPLabels")
  ClassificationList                 = (Get-SettingValue $groupUnified "ClassificationList")
  AllowGuestsToAccessGroups          = (Get-SettingValue $groupUnified "AllowGuestsToAccessGroups")
  AllowToAddGuests                   = (Get-SettingValue $groupUnified "AllowToAddGuests")
  AllowToCreateUpdateDeleteConnectors= (Get-SettingValue $groupUnified "AllowToCreateUpdateDeleteConnectors")
  AllowToAddGuests_ByDefault         = (Get-SettingValue $groupUnified "AllowToAddGuests")
  UsageGuidelinesUrl                 = (Get-SettingValue $groupUnified "UsageGuidelinesUrl")
  GuestUsageGuidelinesUrl            = (Get-SettingValue $groupUnified "GuestUsageGuidelinesUrl")
}

# ---------- Role/Privileged Access Snapshot ----------
Write-Host "[5/7] Reading privileged role assignments..." -ForegroundColor Cyan
$activeRoles = @()
try { $activeRoles = Get-MgDirectoryRole -All -ErrorAction Stop } catch { Write-Warning "Role list failed: $_" }

$criticalRoles = @(
  'Company Administrator',      # Global Admin
  'Privileged Role Administrator',
  'Security Administrator',
  'Conditional Access Administrator',
  'User Administrator'
)
$roleSummary = @()

foreach ($r in $activeRoles) {
  $members = @()
  try { $members = Get-MgDirectoryRoleMember -DirectoryRoleId $r.Id -All -ErrorAction Stop } catch {}
  $count = ($members | Measure-Object).Count
  if ($criticalRoles -contains $r.DisplayName) {
    $roleSummary += [pscustomobject]@{ Role = $r.DisplayName; MemberCount = $count }
  }
}

# ---------- Lightweight Group Ownership Signal (optional) ----------
$groupsWithoutOwners = $null
if (-not $Lightweight) {
  Write-Host "[6/7] Sampling group ownership signals (non-exhaustive)..." -ForegroundColor Cyan
  try {
    # sample top 1000 to avoid massive tenants
    $groups = Get-MgGroup -All -Property Id,DisplayName,GroupTypes,SecurityEnabled,Visibility -ErrorAction Stop | Select-Object -First 1000
    $groupsWithoutOwners = @()
    foreach ($g in $groups) {
      $owners = Get-MgGroupOwner -GroupId $g.Id -All -ErrorAction SilentlyContinue
      if (-not $owners -or $owners.Count -eq 0) {
        $groupsWithoutOwners += [pscustomobject]@{
          GroupId      = $g.Id
          DisplayName  = $g.DisplayName
          Security     = $g.SecurityEnabled
          IsM365       = ($g.GroupTypes -contains "Unified")
          Visibility   = $g.Visibility
        }
      }
    }
  } catch { Write-Warning "Group sampling failed: $_" }
}

# ---------- AD Connect Posture (if local ADSync available) ----------
Write-Host "[7/7] Checking AD Connect (local, optional)..." -ForegroundColor Cyan
$adSyncInfo = $null
if (Get-Module -ListAvailable -Name ADSync) {
  try {
    Import-Module ADSync -ErrorAction Stop
    $sched = Get-ADSyncScheduler
    $global = Get-ADSyncGlobalSettings
    $connectors = Get-ADSyncConnector

    $adSyncInfo = [ordered]@{
      SyncCycleEnabled          = $sched.SyncCycleEnabled
      NextSyncCyclePolicyType   = $sched.NextSyncCyclePolicyType
      StagingMode               = $global.Parameters["StagingMode"]
      PasswordHashSync          = $global.Parameters["PasswordHashSync"]
      SeamlessSSO               = $global.Parameters["EnableSoftMatchOnUpn"] # heuristic; real SSO flag varies by config
      DeviceWriteback           = $global.Parameters["EnableDeviceWriteback"]
      GroupWriteback            = $global.Parameters["EnableGroupWriteback"]
      ConnectorCount            = ($connectors | Measure-Object).Count
    }
  } catch {
    Write-Warning "ADSync module present but query failed: $_"
  }
} else {
  $adSyncInfo = @{ Note = "ADSync module not detected on this host. Run on the Azure AD Connect server for full details." }
}

# ---------- Evaluation & Gaps ----------
function Add-Result {
  param($Category, $Setting, $Observed, $Risk, $Recommendation)
  $results.Add(@{
    Category       = $Category
    Setting        = $Setting
    ObservedValue  = $Observed
    Risk           = $Risk
    Recommendation = $Recommendation
  })
}

# Security defaults vs CA
if ($secDefaults -and $secDefaults.IsEnabled) {
  Add-Result "Baseline" "SecurityDefaults" "Enabled" "Info" "If Conditional Access is mature, consider disabling Security Defaults to avoid conflicts; otherwise keep enabled."
} else {
  Add-Result "Baseline" "SecurityDefaults" "Disabled/Unknown" "Medium" "Ensure Conditional Access comprehensively replaces Security Defaults (MFA for admins/all, legacy auth blocked)."
}

Add-Result "ConditionalAccess" "MFA for All Users" ($hasMfaAll)                  ($(if($hasMfaAll){'Low'}else{'High'})) "Enforce MFA for all users via CA with exceptions only by break-glass."
Add-Result "ConditionalAccess" "MFA for Admin Roles" ($hasMfaAdmins)            ($(if($hasMfaAdmins){'Low'}else{'High'})) "Require MFA for privileged roles at all times."
Add-Result "ConditionalAccess" "Block Legacy Auth"   ($blocksLegacy)            ($(if($blocksLegacy){'Low'}else{'High'})) "Create a CA policy to block legacy authentication clients globally."

# User defaults & consent
Add-Result "Users" "UsersCanRegisterApps" $($userFlags.UsersCanRegisterApps)     ($(if($userFlags.UsersCanRegisterApps){'Medium'}else{'Low'})) "Disable broad user ability to register apps; route through admin/DevOps processes."
Add-Result "Users" "UsersCanCreateSecurityGroups" $($userFlags.UsersCanCreateSecurityGroups) ($(if($userFlags.UsersCanCreateSecurityGroups){'Medium'}else{'Low'})) "Restrict security group creation to admins or approved owners to prevent sprawl."
Add-Result "Users" "UsersCanReadOtherUsers" $($userFlags.UsersCanReadOtherUsers) "Info" "If enabled, ensure privacy review; otherwise keep restricted."
Add-Result "Consent" "PermissionGrantPoliciesAssigned" $($userFlags.PermissionGrantPoliciesAssigned) "Medium" "Ensure user consent to apps is restricted to verified publishers and low-permission scopes."

# Group settings
Add-Result "Groups" "AllowGuestsToAccessGroups" $($groupCfg.AllowGuestsToAccessGroups) "Medium" "Confirm guest access policy aligns with VA baseline; restrict if unnecessary."
Add-Result "Groups" "AllowToAddGuests" $($groupCfg.AllowToAddGuests_ByDefault) "High" "Limit who can add guests; enforce B2B invite policy with review."
Add-Result "Groups" "EnableMIPLabels" $($groupCfg.EnableMIPLabels) "Medium" "Enable sensitivity labels for M365 Groups to enforce data governance."
Add-Result "Groups" "ClassificationListConfigured" $( [string]::IsNullOrEmpty($groupCfg.ClassificationList) -eq $false ) "Low" "Maintain consistent group classification taxonomy."
Add-Result "Groups" "UsageGuidelinesUrl" $($groupCfg.UsageGuidelinesUrl) "Info" "Publish usage guidelines for group owners."
Add-Result "Groups" "GuestUsageGuidelinesUrl" $($groupCfg.GuestUsageGuidelinesUrl) "Info" "Publish guidance for guest collaboration."

# Role snapshot
foreach ($rs in $roleSummary) {
  $risk = if ($rs.MemberCount -gt 5) { "High" } elseif ($rs.MemberCount -gt 0) { "Medium" } else { "Low" }
  Add-Result "PrivilegedRoles" $rs.Role $rs.MemberCount $risk "Reduce permanent members; use PIM for time-bound elevation."
}

# Group ownership sampling
if ($groupsWithoutOwners) {
  Add-Result "Groups" "GroupsWithoutOwners(Sample1000)" ($groupsWithoutOwners.Count) ($(if($groupsWithoutOwners.Count -gt 0){'Medium'}else{'Low'})) "Ensure every group has at least one owner; remediate ownerless groups."
}

# AD Connect
Add-Result "ADConnect" "Summary" ($( $adSyncInfo | ConvertTo-Json -Compress )) "Info" "Validate PHS/SSO/writeback settings and confirm staging/connector design per VA baseline."

# ---------- Output ----------
$outDir = New-OutputFolder
$jsonPath = Join-Path $outDir "Entra_UserGroup_Config_Audit.json"
$csvPath  = Join-Path $outDir "Entra_UserGroup_Config_Audit.csv"

$results | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonPath -Encoding UTF8
$results | ForEach-Object { [pscustomobject]$_ } | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "`nSummary:" -ForegroundColor Green
$results | ForEach-Object { [pscustomobject]$_ } | Sort-Object Category, Risk -Descending | Format-Table -AutoSize

Write-Host "`nWrote:`n  $jsonPath`n  $csvPath" -ForegroundColor Green
