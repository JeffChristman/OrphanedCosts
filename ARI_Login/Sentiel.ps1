<#
.SYNOPSIS
    Collects raw Microsoft Sentinel configuration from a single CENTRAL Azure
    Government workspace (default: vaecla-security-gov), and separately checks
    every accessible subscription to confirm it forwards logs into that
    workspace. Writes everything to one JSON file for analysis in Python.

.DESCRIPTION
    READ-ONLY data collector. No scoring, no pass/fail, no formatting.

    Topology assumed: ONE central Sentinel workspace; all subscriptions ship
    their logs to it. So:
      - Sentinel config (rules, connectors, UEBA, content, etc.) is collected
        ONCE from the central workspace.
      - Diagnostic-settings wiring is collected PER subscription, so the Python
        layer can verify each sub actually targets the central workspace
        (CIS Azure 5.1.x - Activity Log to Log Analytics).

    Output sections:
        metadata
        workspace            (all Sentinel config from the central workspace)
        subscriptionForwarding[]  (per-sub diagnostic settings + resource sample)
        scanErrors

    All CIS mapping / scoring / reporting is left to the Python layer.

.NOTES
    Requires: Az.Accounts, Az.SecurityInsights, Az.OperationalInsights,
              Az.Resources, Az.Monitor  (PowerShell 7+ recommended)
    Identity: read-only (Security Reader + Log Analytics Reader on the central
              workspace; Reader at subscription scope for diagnostic checks)
    Cloud:    AzureUSGovernment
#>

[CmdletBinding()]
param(
    [string]   $WorkspaceName = 'vaecla-security-gov',
    # Sub + RG that HOST the central workspace. If omitted, the script will
    # search accessible subscriptions to locate the workspace automatically.
    [string]   $WorkspaceSubscriptionId,
    [string]   $WorkspaceResourceGroup,
    # Subscriptions to check for log forwarding. Omit to check ALL accessible.
    [string[]] $SubscriptionIds,
    [string]   $TenantId,
    [string]   $OutputPath = ".\Sentinel-Raw-$($WorkspaceName)-$(Get-Date -Format 'yyyyMMdd-HHmmss').json",
    [int]      $IngestionLookbackDays = 3,
    [int]      $ResourceSampleLimit   = 200,
    [switch]   $SkipForwardingCheck    # collect only central workspace config
)

$ErrorActionPreference = 'Stop'
$collectorVersion = '4.0-central'
$scanErrors = [System.Collections.Generic.List[object]]::new()

# ----------------------------------------------------------------------------
# Modules + connect (Azure Government)
# ----------------------------------------------------------------------------
foreach ($m in 'Az.Accounts','Az.SecurityInsights','Az.OperationalInsights','Az.Resources','Az.Monitor') {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        throw "Required module '$m' is not installed. Run: Install-Module $m -Scope CurrentUser"
    }
}

Write-Host "Connecting to Azure Government..." -ForegroundColor Cyan
$connectParams = @{ Environment = 'AzureUSGovernment' }
if ($TenantId) { $connectParams.Tenant = $TenantId }
if (-not (Get-AzContext)) { Connect-AzAccount @connectParams | Out-Null }

$armRoot = (Get-AzContext).Environment.ResourceManagerUrl
function Get-ArmToken { (Get-AzAccessToken -ResourceUrl $armRoot).Token }
function Invoke-Arm {
    param([string]$RelativeUri)
    $uri = "$armRoot$($RelativeUri.TrimStart('/'))"
    Invoke-RestMethod -Uri $uri -Headers @{ Authorization = "Bearer $(Get-ArmToken)" } -Method GET
}

$allSubs = Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' }
if (-not $allSubs) { throw "No accessible/enabled subscriptions found." }

# ----------------------------------------------------------------------------
# Locate the central workspace
# ----------------------------------------------------------------------------
$ws = $null
if ($WorkspaceSubscriptionId -and $WorkspaceResourceGroup) {
    Set-AzContext -Subscription $WorkspaceSubscriptionId | Out-Null
    $ws = Get-AzOperationalInsightsWorkspace -ResourceGroupName $WorkspaceResourceGroup -Name $WorkspaceName
}
else {
    Write-Host "Searching subscriptions for workspace '$WorkspaceName'..." -ForegroundColor Cyan
    foreach ($sub in $allSubs) {
        try {
            Set-AzContext -Subscription $sub.Id | Out-Null
            $found = Get-AzOperationalInsightsWorkspace -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -eq $WorkspaceName } | Select-Object -First 1
            if ($found) {
                $ws = $found
                $WorkspaceSubscriptionId = $sub.Id
                $WorkspaceResourceGroup  = ($found.ResourceId -split '/resourceGroups/')[1].Split('/')[0]
                Write-Host ("  Found in sub {0} (RG {1})" -f $sub.Name, $WorkspaceResourceGroup) -ForegroundColor Green
                break
            }
        } catch { }
    }
}
if (-not $ws) { throw "Workspace '$WorkspaceName' not found in any accessible subscription. Pass -WorkspaceSubscriptionId/-WorkspaceResourceGroup explicitly." }

$wsId  = $ws.ResourceId
$wsCid = $ws.CustomerId
Set-AzContext -Subscription $WorkspaceSubscriptionId | Out-Null

# ----------------------------------------------------------------------------
# Collect Sentinel config ONCE from the central workspace
# ----------------------------------------------------------------------------
$wsErrors = [System.Collections.Generic.List[object]]::new()
function Invoke-Section {
    param([string]$Name, [scriptblock]$Block)
    try { & $Block }
    catch {
        $wsErrors.Add([pscustomobject]@{ section = $Name; error = $_.Exception.Message })
        Write-Warning "  [$Name] $($_.Exception.Message)"
        return $null
    }
}

Write-Host "Collecting Sentinel config from central workspace..." -ForegroundColor Cyan
$rg = $WorkspaceResourceGroup; $subId = $WorkspaceSubscriptionId; $wn = $WorkspaceName

$automationRules = Invoke-Section 'automationRules' {
    Get-AzSentinelAutomationRule -ResourceGroupName $rg -WorkspaceName $wn |
        Select-Object Name, DisplayName, Order, TriggeringLogic, Action
}
$summaryRules = Invoke-Section 'summaryRules' {
    (Invoke-Arm "subscriptions/$subId/resourceGroups/$rg/providers/Microsoft.OperationalInsights/workspaces/$wn/summaryLogs?api-version=2023-09-01").value
}
$dataConnectors = Invoke-Section 'dataConnectors' {
    Get-AzSentinelDataConnector -ResourceGroupName $rg -WorkspaceName $wn | Select-Object Name, Kind, Id
}
$ingestionFreshness = Invoke-Section 'ingestionFreshness' {
    $kql = "union withsource=_Tbl * | where TimeGenerated > ago($($IngestionLookbackDays)d) | summarize LastSeen=max(TimeGenerated), Records=count() by _Tbl"
    (Invoke-AzOperationalInsightsQuery -WorkspaceId $wsCid -Query $kql).Results
}
$workspaceManager = Invoke-Section 'workspaceManager' {
    Get-AzSentinelOnboardingState -ResourceGroupName $rg -WorkspaceName $wn | Select-Object Name, CustomerManagedKey
}
$analyticsRules = Invoke-Section 'analyticsRules' {
    Get-AzSentinelAlertRule -ResourceGroupName $rg -WorkspaceName $wn |
        Select-Object Name, DisplayName, Enabled, Kind, Severity, Tactics, Techniques, Query, QueryFrequency, QueryPeriod
}
$ueba = Invoke-Section 'ueba' {
    (Invoke-Arm "subscriptions/$subId/resourceGroups/$rg/providers/Microsoft.OperationalInsights/workspaces/$wn/providers/Microsoft.SecurityInsights/settings/Ueba?api-version=2023-11-01").properties
}
$contentPackages = Invoke-Section 'contentPackages' {
    (Invoke-Arm "subscriptions/$subId/resourceGroups/$rg/providers/Microsoft.OperationalInsights/workspaces/$wn/providers/Microsoft.SecurityInsights/contentPackages?api-version=2023-11-01").value |
        ForEach-Object {
            [pscustomobject]@{
                name = $_.properties.displayName; contentId = $_.properties.contentId
                installedVersion = $_.properties.installedVersion
                latestVersion = $_.properties.version; contentKind = $_.properties.contentKind
            }
        }
}

$workspace = [pscustomobject]@{
    subscriptionId      = $subId
    resourceGroup       = $rg
    workspaceName       = $wn
    workspaceId         = $wsId
    workspaceLocation   = $ws.Location
    workspaceCustomerId = $wsCid
    automationRules     = @($automationRules)
    summaryRules        = @($summaryRules)
    dataConnectors      = @($dataConnectors)
    ingestionFreshness  = @($ingestionFreshness)
    workspaceManager    = $workspaceManager
    analyticsRules      = @($analyticsRules)
    ueba                = $ueba
    contentPackages     = @($contentPackages)
    collectionErrors    = @($wsErrors)
}

# ----------------------------------------------------------------------------
# Per-subscription forwarding check (does each sub send logs to central WS?)
# ----------------------------------------------------------------------------
$forwarding = [System.Collections.Generic.List[object]]::new()
if (-not $SkipForwardingCheck) {
    $subsToCheck = if ($SubscriptionIds) { $allSubs | Where-Object { $_.Id -in $SubscriptionIds } } else { $allSubs }
    Write-Host ("`nChecking log forwarding across {0} subscription(s)..." -f $subsToCheck.Count) -ForegroundColor Cyan

    foreach ($sub in $subsToCheck) {
        Write-Host ("  {0} ({1})" -f $sub.Name, $sub.Id) -ForegroundColor DarkCyan
        $entry = [ordered]@{
            subscriptionId   = $sub.Id
            subscriptionName = $sub.Name
            activityLogToCentral = $null
            subscriptionDiagnostics = @()
            resourceSample   = @()
            errors           = @()
        }
        try {
            Set-AzContext -Subscription $sub.Id | Out-Null

            # Subscription Activity Log diagnostic settings (raw) + does any target central WS
            $subDiag = Get-AzDiagnosticSetting -ResourceId "/subscriptions/$($sub.Id)" -ErrorAction SilentlyContinue
            $entry.subscriptionDiagnostics = @($subDiag | Select-Object Name, WorkspaceId, @{n='Logs';e={ $_.Log | Select-Object Category, Enabled }})
            $entry.activityLogToCentral = [bool]($subDiag | Where-Object { $_.WorkspaceId -eq $wsId })

            # Sampled resource diagnostic settings (raw, capped)
            $targets = @()
            $targets += Get-AzResource -ResourceType 'Microsoft.KeyVault/vaults' -ErrorAction SilentlyContinue
            $targets += Get-AzResource -ResourceType 'Microsoft.Network/networkSecurityGroups' -ErrorAction SilentlyContinue
            $targets += Get-AzResource -ResourceType 'Microsoft.Storage/storageAccounts' -ErrorAction SilentlyContinue
            $targets = $targets | Select-Object -First $ResourceSampleLimit
            $entry.resourceSample = foreach ($r in $targets) {
                $d = Get-AzDiagnosticSetting -ResourceId $r.ResourceId -ErrorAction SilentlyContinue
                [pscustomobject]@{
                    resourceId   = $r.ResourceId
                    resourceType = $r.ResourceType
                    toCentral    = [bool]($d | Where-Object { $_.WorkspaceId -eq $wsId })
                    settings     = @($d | Select-Object Name, WorkspaceId)
                }
            }
        }
        catch {
            $entry.errors += $_.Exception.Message
            $scanErrors.Add([pscustomobject]@{ scope = $sub.Id; error = $_.Exception.Message })
        }
        $forwarding.Add([pscustomobject]$entry)
    }
}

# ----------------------------------------------------------------------------
# Assemble single JSON document
# ----------------------------------------------------------------------------
$document = [ordered]@{
    metadata = [ordered]@{
        collectorVersion       = $collectorVersion
        collectedUtc           = (Get-Date).ToUniversalTime().ToString('o')
        cloud                  = 'AzureUSGovernment'
        centralWorkspace       = $WorkspaceName
        centralWorkspaceId     = $wsId
        subscriptionsChecked   = @($forwarding | Select-Object subscriptionId, subscriptionName)
        ingestionLookbackDays  = $IngestionLookbackDays
    }
    workspace              = $workspace
    subscriptionForwarding = @($forwarding)
    scanErrors             = @($scanErrors)
}

$document | ConvertTo-Json -Depth 14 | Out-File $OutputPath -Encoding UTF8

Write-Host "`n=================================================" -ForegroundColor Green
Write-Host ("Central workspace     : {0}" -f $WorkspaceName)
Write-Host ("Subscriptions checked : {0}" -f $forwarding.Count)
$fwd = @($forwarding | Where-Object { $_.activityLogToCentral }).Count
Write-Host ("Subs w/ Activity Log -> central : {0} of {1}" -f $fwd, $forwarding.Count) -ForegroundColor $(if ($fwd -eq $forwarding.Count) {'Green'} else {'Yellow'})
Write-Host ("Output                : {0}" -f $OutputPath) -ForegroundColor Cyan
Write-Host "Hand this JSON to the Python analyzer for CIS delta + reporting." -ForegroundColor Cyan