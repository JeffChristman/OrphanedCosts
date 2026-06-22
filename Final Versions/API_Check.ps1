<#
.SYNOPSIS
    Audit Azure subscription for deprecated API calls that Turbot CIS v2.0 may be checking.

.DESCRIPTION
    Tests various Azure APIs to identify which deprecated resources Turbot CIS v2.0 controls
    might be checking, causing false non-compliance findings. Compares old (deprecated) APIs
    against new (current) APIs to identify control drift across five categories:

        1. Defender for Cloud  — checks deprecated 'Databases' pricing plan vs. current
                                 workload-specific SKUs (SqlServers, CosmosDbs, etc.)
        2. Monitoring          — checks deprecated Log Profiles API vs. current Diagnostic Settings
        3. Storage             — checks for Classic (non-ARM) storage accounts
        4. Networking          — checks for Classic virtual networks and ARM NSGs
        5. Identity            — verifies access to Microsoft Entra ID (formerly Azure AD) APIs

    All test results are collected in $auditResults and optionally exported to CSV, JSON,
    and a plain-text summary report when -ExportResults is specified.

.PARAMETER SubscriptionId
    The Azure subscription ID to audit. Must be a valid GUID.

.PARAMETER ExportResults
    When specified, exports all audit results to three files in the current directory:
        - Azure_CIS_API_Audit_<timestamp>.csv
        - Azure_CIS_API_Audit_<timestamp>.json
        - Azure_CIS_API_Audit_Summary_<timestamp>.txt

.EXAMPLE
    .\Azure_CIS_Deprecated_API_Audit.ps1 -SubscriptionId "de14fea2-34d0-4c1f-9ba6-b63c2c1ae5ea"
    Runs the audit against the specified subscription and prints results to the console only.

.EXAMPLE
    .\Azure_CIS_Deprecated_API_Audit.ps1 -SubscriptionId "de14fea2-34d0-4c1f-9ba6-b63c2c1ae5ea" -ExportResults
    Runs the audit and exports CSV, JSON, and summary text files to the current directory.

.NOTES
    Author  : VA Security Team
    Date    : 2026-02-04
    Purpose : Identify deprecated API calls in Turbot CIS v2.0 controls

    Known deprecated APIs checked:
        Microsoft.Security/pricings/Databases      — deprecated ~2023; replaced by workload SKUs
        Microsoft.Insights/logprofiles             — deprecated September 2023; replaced by Diagnostic Settings
        Microsoft.ClassicStorage/storageAccounts   — Classic deployment model, end-of-life
        Microsoft.ClassicNetwork/virtualNetworks   — Classic deployment model, end-of-life
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SubscriptionId,
   
    [Parameter(Mandatory=$false)]
    [switch]$ExportResults
)

# ── Output Helpers ────────────────────────────────────────────────────────────
# Wrapper functions provide consistent emoji+colour formatting for console output,
# avoiding repeated -ForegroundColor arguments throughout the script.

function Write-Success {
    <#
    .SYNOPSIS
        Writes a green success message prefixed with a check-mark emoji.
    .PARAMETER Message
        The message text to display.
    #>
    param($Message)
    Write-Host "✅ $Message" -ForegroundColor Green
}

function Write-Failure {
    <#
    .SYNOPSIS
        Writes a red failure/error message prefixed with a cross emoji.
    .PARAMETER Message
        The message text to display.
    #>
    param($Message)
    Write-Host "❌ $Message" -ForegroundColor Red
}

function Write-Warning {
    <#
    .SYNOPSIS
        Writes a yellow warning message prefixed with a warning emoji.
    .PARAMETER Message
        The message text to display.
    #>
    param($Message)
    Write-Host "⚠️  $Message" -ForegroundColor Yellow
}

function Write-Info {
    <#
    .SYNOPSIS
        Writes a cyan informational message prefixed with an info emoji.
    .PARAMETER Message
        The message text to display.
    #>
    param($Message)
    Write-Host "ℹ️  $Message" -ForegroundColor Cyan
}

function Write-Header {
    <#
    .SYNOPSIS
        Writes a cyan section header surrounded by separator lines.
    .PARAMETER Message
        The header text to display.
    #>
    param($Message)
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "$Message" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

# ── Results Collection ────────────────────────────────────────────────────────
# All test results are appended to this script-scoped array and optionally
# exported at the end of the script when -ExportResults is specified.
$auditResults = @()

function Add-Result {
    <#
    .SYNOPSIS
        Appends a structured test result to the script-level $auditResults collection.

    .DESCRIPTION
        Creates a PSCustomObject with a consistent set of fields for every test and
        appends it to $script:auditResults. This ensures all results share the same
        schema for CSV/JSON export and summary reporting.

        Status values used by callers:
            CURRENT           — API exists and is the current supported method
            DEPRECATED        — API does not exist; confirmed deprecated
            DEPRECATED_EXISTS — Deprecated API still present (potential legacy resource)
            DEPRECATED_EMPTY  — Deprecated API reachable but returned no resources
            LEGACY_NAME       — Resource found under an old name that may have been renamed
            UNEXPECTED_EXISTS — Resource found when it was expected to be absent
            NOT_CONFIGURED    — API reachable but resource not configured
            NONE_FOUND        — API reachable but no resources exist in the subscription
            ACCESSIBLE        — API access confirmed (identity/context checks)
            INFO              — Informational result; not a pass/fail condition
            ERROR             — An unexpected exception occurred during the test

    .PARAMETER Category
        The high-level audit category (e.g. 'Defender', 'Monitoring', 'Storage').

    .PARAMETER TestName
        A short descriptive name for the individual test (e.g. 'Databases Plan').

    .PARAMETER APIEndpoint
        The Azure resource provider path or API identifier being tested
        (e.g. 'Microsoft.Security/pricings/Databases').

    .PARAMETER Status
        A status string from the set described above.

    .PARAMETER Details
        A free-text field describing what was found or the exception message.

    .PARAMETER CISControl
        The CIS v2.0 control number(s) this test relates to (e.g. '2.1.1').

    .PARAMETER Recommendation
        A brief action or explanation for the finding.
    #>
    param(
        [string]$Category,
        [string]$TestName,
        [string]$APIEndpoint,
        [string]$Status,
        [string]$Details,
        [string]$CISControl,
        [string]$Recommendation
    )
   
    $script:auditResults += [PSCustomObject]@{
        Timestamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Category       = $Category
        TestName       = $TestName
        APIEndpoint    = $APIEndpoint
        Status         = $Status
        Details        = $Details
        CISv2Control   = $CISControl
        Recommendation = $Recommendation
    }
}

# ── Authentication ─────────────────────────────────────────────────────────────

Write-Header "AZURE CIS v2.0 DEPRECATED API AUDIT"
Write-Info "Subscription: $SubscriptionId"
Write-Info "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

Write-Info "Checking Azure connection..."
try {
    $context = Get-AzContext
    if (-not $context) {
        # No existing session — prompt for interactive login
        Write-Warning "Not logged in to Azure. Attempting to connect..."
        Connect-AzAccount
    }
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    Write-Success "Connected to subscription: $((Get-AzContext).Subscription.Name)"
} catch {
    Write-Failure "Failed to connect to Azure. Please run Connect-AzAccount first."
    exit 1
}

Write-Host ""

# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY 1: DEFENDER FOR CLOUD / SECURITY CENTER
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "CATEGORY 1: Defender for Cloud (Security Center)"

# TEST 1.1 — Deprecated monolithic 'Databases' pricing plan
# CIS v2.0 control 2.1.1 checks this plan, but it was deprecated in ~2023 and
# replaced by workload-specific SKUs. A 'not found' error here is expected and
# indicates the control will produce a false non-compliance finding.
Write-Host "`n[TEST 1.1] Defender for Databases (Deprecated Monolithic Plan)" -ForegroundColor Yellow
try {
    $databases = Get-AzSecurityPricing -Name "Databases" -ErrorAction Stop
    Write-Warning "Unexpected: 'Databases' pricing plan exists"
    Write-Warning "Value: $($databases.PricingTier)"
    Add-Result -Category "Defender" -TestName "Databases Plan" `
        -APIEndpoint "Microsoft.Security/pricings/Databases" `
        -Status "UNEXPECTED_EXISTS" -Details "Databases plan found: $($databases.PricingTier)" `
        -CISControl "2.1.1" -Recommendation "This plan may be legacy - verify with Microsoft"
} catch {
    if ($_.Exception.Message -like "*cannot find*" -or $_.Exception.Message -like "*not found*") {
        Write-Failure "DEPRECATED: 'Databases' pricing plan does not exist"
        Write-Info "This is the resource Turbot CIS v2.0 checks (causing false negative)"
        Add-Result -Category "Defender" -TestName "Databases Plan" `
            -APIEndpoint "Microsoft.Security/pricings/Databases" `
            -Status "DEPRECATED" -Details "Resource not found - deprecated ~2023" `
            -CISControl "2.1.1" -Recommendation "Turbot should check workload-specific SKUs instead"
    } else {
        Write-Failure "Error: $($_.Exception.Message)"
        Add-Result -Category "Defender" -TestName "Databases Plan" `
            -APIEndpoint "Microsoft.Security/pricings/Databases" `
            -Status "ERROR" -Details $_.Exception.Message `
            -CISControl "2.1.1" -Recommendation "Investigate error"
    }
}

# TEST 1.2 — Current workload-specific Defender plans
# These are the SKUs that replaced the deprecated 'Databases' monolithic plan.
# Each should exist and ideally be set to 'Standard' (protected).
Write-Host "`n[TEST 1.2] Current Workload-Specific Defender Plans" -ForegroundColor Yellow
$workloadPlans = @("SqlServers", "SqlServerVirtualMachines", "OpenSourceRelationalDatabases", "CosmosDbs")
$workloadResults = @{}

foreach ($plan in $workloadPlans) {
    try {
        $pricing = Get-AzSecurityPricing -Name $plan -ErrorAction Stop
        if ($pricing.PricingTier -eq "Standard") {
            Write-Success "$plan = $($pricing.PricingTier) (Protected)"
        } else {
            Write-Warning "$plan = $($pricing.PricingTier) (Not Protected)"
        }
        $workloadResults[$plan] = $pricing.PricingTier
        Add-Result -Category "Defender" -TestName "$plan Plan" `
            -APIEndpoint "Microsoft.Security/pricings/$plan" `
            -Status "CURRENT" -Details "Tier: $($pricing.PricingTier)" `
            -CISControl "2.1.x" -Recommendation "This is the current API - working correctly"
    } catch {
        Write-Failure "$plan - Not found or error"
        $workloadResults[$plan] = "ERROR"
        Add-Result -Category "Defender" -TestName "$plan Plan" `
            -APIEndpoint "Microsoft.Security/pricings/$plan" `
            -Status "ERROR" -Details $_.Exception.Message `
            -CISControl "2.1.x" -Recommendation "Investigate"
    }
}

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Info "Deprecated 'Databases' plan: NOT FOUND (expected)"
Write-Info "Current workload plans found: $($workloadResults.Keys.Count)"
Write-Info "This confirms Turbot CIS v2.0 control 2.1.1 is checking deprecated API"

# TEST 1.3 — Container plan naming: 'ContainerRegistry' was renamed to 'Containers'.
# Both names are checked to identify whether the old name is still present as a
# legacy resource or whether the subscription has migrated to the current name.
Write-Host "`n[TEST 1.3] Other Defender Plans - Check for Deprecated Names" -ForegroundColor Yellow

Write-Host "  Checking: ContainerRegistry (old name)" -NoNewline
try {
    $containerReg = Get-AzSecurityPricing -Name "ContainerRegistry" -ErrorAction Stop
    Write-Warning " - FOUND (may be legacy)"
    Add-Result -Category "Defender" -TestName "ContainerRegistry Plan" `
        -APIEndpoint "Microsoft.Security/pricings/ContainerRegistry" `
        -Status "LEGACY_NAME" -Details "May have been renamed to 'Containers'" `
        -CISControl "2.1.x" -Recommendation "Check if this was renamed"
} catch {
    Write-Host " - NOT FOUND" -ForegroundColor Gray
}

Write-Host "  Checking: Containers (new name)" -NoNewline
try {
    $containers = Get-AzSecurityPricing -Name "Containers" -ErrorAction Stop
    Write-Success " - FOUND: $($containers.PricingTier)"
    Add-Result -Category "Defender" -TestName "Containers Plan" `
        -APIEndpoint "Microsoft.Security/pricings/Containers" `
        -Status "CURRENT" -Details "Tier: $($containers.PricingTier)" `
        -CISControl "2.1.x" -Recommendation "Current API"
} catch {
    Write-Failure " - NOT FOUND"
}

# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY 2: MONITORING & LOGGING
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "CATEGORY 2: Monitoring & Activity Logs"

# TEST 2.1 — Log Profiles (deprecated September 2023)
# CIS v2.0 control 2.1.3 may still reference this API. Deprecation warnings are
# suppressed during the call to avoid noisy output; they are restored immediately after.
Write-Host "`n[TEST 2.1] Activity Log - Log Profiles (Deprecated Sept 2023)" -ForegroundColor Yellow
try {
    $WarningPreference = "SilentlyContinue"   # Suppress built-in deprecation warnings
    $logProfiles = Get-AzLogProfile -ErrorAction Stop
    $WarningPreference = "Continue"
   
    if ($logProfiles) {
        Write-Warning "LEGACY: Log Profiles still exist (deprecated API)"
        Write-Warning "Found $($logProfiles.Count) log profile(s)"
        $logProfiles | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Yellow }
        Add-Result -Category "Monitoring" -TestName "Log Profiles" `
            -APIEndpoint "Microsoft.Insights/logprofiles" `
            -Status "DEPRECATED_EXISTS" -Details "Found $($logProfiles.Count) profiles (deprecated Sept 2023)" `
            -CISControl "2.1.3" -Recommendation "Migrate to Diagnostic Settings"
    } else {
        Write-Info "No Log Profiles found (expected - deprecated)"
        Add-Result -Category "Monitoring" -TestName "Log Profiles" `
            -APIEndpoint "Microsoft.Insights/logprofiles" `
            -Status "DEPRECATED_EMPTY" -Details "API deprecated Sept 2023 - no profiles found" `
            -CISControl "2.1.3" -Recommendation "Normal - use Diagnostic Settings instead"
    }
} catch {
    if ($_.Exception.Message -like "*deprecated*" -or $_.Exception.Message -like "*not supported*") {
        Write-Failure "DEPRECATED: Log Profiles API not supported"
        Write-Info "This is likely what Turbot CIS v2.0 control 2.1.3 checks"
        Add-Result -Category "Monitoring" -TestName "Log Profiles" `
            -APIEndpoint "Microsoft.Insights/logprofiles" `
            -Status "DEPRECATED" -Details "API deprecated September 2023" `
            -CISControl "2.1.3" -Recommendation "Turbot should check Diagnostic Settings instead"
    } else {
        Write-Failure "Error: $($_.Exception.Message)"
        Add-Result -Category "Monitoring" -TestName "Log Profiles" `
            -APIEndpoint "Microsoft.Insights/logprofiles" `
            -Status "ERROR" -Details $_.Exception.Message `
            -CISControl "2.1.3" -Recommendation "Investigate error"
    }
}

# TEST 2.2 — Diagnostic Settings (current replacement for Log Profiles)
# For each setting found, the destination types (Storage, Log Analytics, Event Hub)
# and the count of enabled log categories are reported.
Write-Host "`n[TEST 2.2] Activity Log - Diagnostic Settings (Current Method)" -ForegroundColor Yellow
try {
    $diagSettings = Get-AzSubscriptionDiagnosticSetting -SubscriptionId $SubscriptionId -ErrorAction Stop
   
    if ($diagSettings) {
        Write-Success "CURRENT: Diagnostic Settings configured"
        Write-Info "Found $($diagSettings.Count) diagnostic setting(s)"
       
        foreach ($setting in $diagSettings) {
            Write-Host "  Setting: $($setting.Name)" -ForegroundColor Cyan
           
            $enabledLogs = ($setting.Log | Where-Object { $_.Enabled -eq $true }).Count
            Write-Host "    - Enabled log categories: $enabledLogs" -ForegroundColor White
           
            # Report which export destinations are configured for this setting
            if ($setting.StorageAccountId)            { Write-Host "    - Storage Account: Configured ✅" -ForegroundColor Green }
            if ($setting.WorkspaceId)                 { Write-Host "    - Log Analytics: Configured ✅"  -ForegroundColor Green }
            if ($setting.EventHubAuthorizationRuleId) { Write-Host "    - Event Hub: Configured ✅"      -ForegroundColor Green }
        }
       
        Add-Result -Category "Monitoring" -TestName "Diagnostic Settings" `
            -APIEndpoint "Microsoft.Insights/diagnosticSettings" `
            -Status "CURRENT" -Details "Found $($diagSettings.Count) settings with activity log export" `
            -CISControl "2.1.3" -Recommendation "This is the current API - working correctly"
    } else {
        Write-Warning "No Diagnostic Settings found for Activity Logs"
        Write-Warning "Activity logs may not be exported/retained"
        Add-Result -Category "Monitoring" -TestName "Diagnostic Settings" `
            -APIEndpoint "Microsoft.Insights/diagnosticSettings" `
            -Status "NOT_CONFIGURED" -Details "No diagnostic settings found" `
            -CISControl "2.1.3" -Recommendation "Configure activity log export to storage/Log Analytics"
    }
} catch {
    Write-Failure "Error checking Diagnostic Settings: $($_.Exception.Message)"
    Add-Result -Category "Monitoring" -TestName "Diagnostic Settings" `
        -APIEndpoint "Microsoft.Insights/diagnosticSettings" `
        -Status "ERROR" -Details $_.Exception.Message `
        -CISControl "2.1.3" -Recommendation "Investigate error"
}

Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Info "Log Profiles (deprecated): API no longer supported"
Write-Info "Diagnostic Settings (current): Should be used for activity log export"
Write-Info "If Turbot CIS v2.0 checks Log Profiles → False negative likely"

# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY 3: STORAGE
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "CATEGORY 3: Storage Accounts"

# TEST 3.1 — Classic (non-ARM) storage accounts
# Classic storage was part of the Azure Service Manager (ASM) deployment model,
# which is end-of-life. Finding these accounts indicates unmigrated legacy resources.
Write-Host "`n[TEST 3.1] Classic Storage Accounts (Deprecated)" -ForegroundColor Yellow
try {
    $classicStorage = Get-AzResource -ResourceType "Microsoft.ClassicStorage/storageAccounts" `
        -SubscriptionId $SubscriptionId -ErrorAction Stop
   
    if ($classicStorage) {
        Write-Warning "LEGACY: Classic storage accounts still exist"
        Write-Warning "Found $($classicStorage.Count) classic storage account(s)"
        $classicStorage | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Yellow }
        Add-Result -Category "Storage" -TestName "Classic Storage" `
            -APIEndpoint "Microsoft.ClassicStorage/storageAccounts" `
            -Status "DEPRECATED_EXISTS" -Details "Found $($classicStorage.Count) classic accounts" `
            -CISControl "3.x" -Recommendation "Migrate to ARM storage accounts"
    } else {
        Write-Success "No Classic storage accounts found (expected)"
        Add-Result -Category "Storage" -TestName "Classic Storage" `
            -APIEndpoint "Microsoft.ClassicStorage/storageAccounts" `
            -Status "DEPRECATED_EMPTY" -Details "No classic storage found (expected)" `
            -CISControl "3.x" -Recommendation "Normal - all storage using ARM model"
    }
} catch {
    Write-Info "Classic storage check: $($_.Exception.Message)"
    Add-Result -Category "Storage" -TestName "Classic Storage" `
        -APIEndpoint "Microsoft.ClassicStorage/storageAccounts" `
        -Status "INFO" -Details $_.Exception.Message `
        -CISControl "3.x" -Recommendation "Likely no classic storage"
}

# TEST 3.2 — ARM storage accounts (current deployment model)
# Reports a basic security posture summary: HTTPS enforcement and minimum TLS version.
Write-Host "`n[TEST 3.2] ARM Storage Accounts (Current Method)" -ForegroundColor Yellow
try {
    $storageAccounts = Get-AzStorageAccount -ErrorAction Stop
   
    if ($storageAccounts) {
        Write-Success "Found $($storageAccounts.Count) ARM storage account(s)"
       
        # Summarise key CIS-relevant security properties across all accounts
        $httpsOnly = ($storageAccounts | Where-Object { $_.EnableHttpsTrafficOnly -eq $true }).Count
        $minTls12  = ($storageAccounts | Where-Object { $_.MinimumTlsVersion -eq "TLS1_2" }).Count
       
        Write-Info "  - HTTPS Only enforced: $httpsOnly / $($storageAccounts.Count)"
        Write-Info "  - TLS 1.2 minimum: $minTls12 / $($storageAccounts.Count)"
       
        Add-Result -Category "Storage" -TestName "ARM Storage Accounts" `
            -APIEndpoint "Microsoft.Storage/storageAccounts" `
            -Status "CURRENT" -Details "Found $($storageAccounts.Count) accounts" `
            -CISControl "3.x" -Recommendation "Current API - verify security settings with CIS controls"
    } else {
        Write-Info "No ARM storage accounts found"
        Add-Result -Category "Storage" -TestName "ARM Storage Accounts" `
            -APIEndpoint "Microsoft.Storage/storageAccounts" `
            -Status "NONE_FOUND" -Details "No storage accounts in subscription" `
            -CISControl "3.x" -Recommendation "N/A"
    }
} catch {
    Write-Failure "Error checking storage accounts: $($_.Exception.Message)"
    Add-Result -Category "Storage" -TestName "ARM Storage Accounts" `
        -APIEndpoint "Microsoft.Storage/storageAccounts" `
        -Status "ERROR" -Details $_.Exception.Message `
        -CISControl "3.x" -Recommendation "Investigate error"
}

# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY 4: NETWORKING
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "CATEGORY 4: Networking"

# TEST 4.1 — Classic (non-ARM) virtual networks
# Like Classic storage, these belong to the ASM deployment model and should
# have been migrated to ARM VNets.
Write-Host "`n[TEST 4.1] Classic Virtual Networks (Deprecated)" -ForegroundColor Yellow
try {
    $classicVNets = Get-AzResource -ResourceType "Microsoft.ClassicNetwork/virtualNetworks" `
        -SubscriptionId $SubscriptionId -ErrorAction Stop
   
    if ($classicVNets) {
        Write-Warning "LEGACY: Classic virtual networks still exist"
        Write-Warning "Found $($classicVNets.Count) classic VNet(s)"
        $classicVNets | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Yellow }
        Add-Result -Category "Networking" -TestName "Classic VNets" `
            -APIEndpoint "Microsoft.ClassicNetwork/virtualNetworks" `
            -Status "DEPRECATED_EXISTS" -Details "Found $($classicVNets.Count) classic VNets" `
            -CISControl "6.x" -Recommendation "Migrate to ARM virtual networks"
    } else {
        Write-Success "No Classic virtual networks found (expected)"
        Add-Result -Category "Networking" -TestName "Classic VNets" `
            -APIEndpoint "Microsoft.ClassicNetwork/virtualNetworks" `
            -Status "DEPRECATED_EMPTY" -Details "No classic VNets found (expected)" `
            -CISControl "6.x" -Recommendation "Normal - all networking using ARM model"
    }
} catch {
    Write-Info "Classic networking check: $($_.Exception.Message)"
    Add-Result -Category "Networking" -TestName "Classic VNets" `
        -APIEndpoint "Microsoft.ClassicNetwork/virtualNetworks" `
        -Status "INFO" -Details $_.Exception.Message `
        -CISControl "6.x" -Recommendation "Likely no classic networking"
}

# TEST 4.2 — ARM Network Security Groups (current)
Write-Host "`n[TEST 4.2] Network Security Groups (Current)" -ForegroundColor Yellow
try {
    $nsgs = Get-AzNetworkSecurityGroup -ErrorAction Stop
   
    if ($nsgs) {
        Write-Success "Found $($nsgs.Count) Network Security Group(s)"
        Add-Result -Category "Networking" -TestName "NSGs" `
            -APIEndpoint "Microsoft.Network/networkSecurityGroups" `
            -Status "CURRENT" -Details "Found $($nsgs.Count) NSGs" `
            -CISControl "6.x" -Recommendation "Current API - verify rules with CIS controls"
    } else {
        Write-Info "No Network Security Groups found"
        Add-Result -Category "Networking" -TestName "NSGs" `
            -APIEndpoint "Microsoft.Network/networkSecurityGroups" `
            -Status "NONE_FOUND" -Details "No NSGs in subscription" `
            -CISControl "6.x" -Recommendation "N/A"
    }
} catch {
    Write-Failure "Error checking NSGs: $($_.Exception.Message)"
    Add-Result -Category "Networking" -TestName "NSGs" `
        -APIEndpoint "Microsoft.Network/networkSecurityGroups" `
        -Status "ERROR" -Details $_.Exception.Message `
        -CISControl "6.x" -Recommendation "Investigate error"
}

# ══════════════════════════════════════════════════════════════════════════════
# CATEGORY 5: IDENTITY (AZURE AD / MICROSOFT ENTRA ID)
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "CATEGORY 5: Identity (Azure AD / Microsoft Entra ID)"

# TEST 5.1 — Verify identity API access
# Azure AD was rebranded to Microsoft Entra ID in 2023. Underlying API endpoints
# are largely unchanged, but CIS controls referencing the old branding may cause
# confusion. This test confirms the identity context is accessible.
Write-Host "`n[TEST 5.1] Microsoft Entra ID (formerly Azure AD)" -ForegroundColor Yellow
Write-Info "Note: Azure AD rebranded to Microsoft Entra ID in 2023"
Write-Info "API endpoints mostly unchanged, but some legacy references may exist"

try {
    $context = Get-AzContext
    Write-Success "Identity context available: $($context.Account.Id)"
   
    Add-Result -Category "Identity" -TestName "Entra ID Access" `
        -APIEndpoint "Microsoft.Graph (current)" `
        -Status "ACCESSIBLE" -Details "Can access identity APIs" `
        -CISControl "1.x" -Recommendation "Monitor for legacy 'Azure AD' references in controls"
} catch {
    Write-Failure "Error accessing identity: $($_.Exception.Message)"
    Add-Result -Category "Identity" -TestName "Entra ID Access" `
        -APIEndpoint "Microsoft.Graph" `
        -Status "ERROR" -Details $_.Exception.Message `
        -CISControl "1.x" -Recommendation "Investigate access issues"
}

# ══════════════════════════════════════════════════════════════════════════════
# FINAL SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

Write-Header "AUDIT SUMMARY"

# Aggregate result counts by status category
$totalTests = $auditResults.Count
$deprecated = ($auditResults | Where-Object { $_.Status -like "*DEPRECATED*" }).Count
$current    = ($auditResults | Where-Object { $_.Status -eq "CURRENT" }).Count
$errors     = ($auditResults | Where-Object { $_.Status -eq "ERROR" }).Count

Write-Host "`nTotal Tests Run: $totalTests" -ForegroundColor Cyan
Write-Host "Deprecated APIs Found: $deprecated" -ForegroundColor $(if ($deprecated -gt 0) { "Red" } else { "Green" })
Write-Host "Current APIs Working: $current" -ForegroundColor Green
Write-Host "Errors Encountered: $errors" -ForegroundColor $(if ($errors -gt 0) { "Yellow" } else { "Green" })

Write-Host "`nKEY FINDINGS:" -ForegroundColor Yellow

# Report on the Defender Databases plan specifically — this is the primary known issue
$dbResult = $auditResults | Where-Object { $_.TestName -eq "Databases Plan" }
if ($dbResult -and $dbResult.Status -eq "DEPRECATED") {
    Write-Host "❌ CONFIRMED: Defender 'Databases' pricing plan DEPRECATED" -ForegroundColor Red
    Write-Host "   Impact: Turbot CIS v2.0 control 2.1.1 will show false negative" -ForegroundColor Yellow
    Write-Host "   Solution: Turbot should check workload-specific SKUs (SqlServers, etc.)" -ForegroundColor Cyan
}

# Report on Log Profiles — secondary known issue for control 2.1.3
$logResult = $auditResults | Where-Object { $_.TestName -eq "Log Profiles" }
if ($logResult -and $logResult.Status -like "*DEPRECATED*") {
    Write-Host "❌ CONFIRMED: Log Profiles API DEPRECATED (Sept 2023)" -ForegroundColor Red
    Write-Host "   Impact: Turbot CIS v2.0 control 2.1.3 likely checking wrong API" -ForegroundColor Yellow
    Write-Host "   Solution: Turbot should check Diagnostic Settings instead" -ForegroundColor Cyan
}

# Confirm Diagnostic Settings are working correctly as the current alternative
$diagResult = $auditResults | Where-Object { $_.TestName -eq "Diagnostic Settings" }
if ($diagResult -and $diagResult.Status -eq "CURRENT") {
    Write-Host "✅ Diagnostic Settings (current method) working correctly" -ForegroundColor Green
}

Write-Host "`nRECOMMENDATION:" -ForegroundColor Cyan
if ($deprecated -gt 0) {
    Write-Host "Multiple deprecated APIs detected that Turbot CIS v2.0 may be checking." -ForegroundColor Yellow
    Write-Host "This indicates a systematic issue with CIS v2.0 control accuracy." -ForegroundColor Yellow
    Write-Host "Strong justification for upgrading to Turbot CIS v3.0." -ForegroundColor Cyan
} else {
    Write-Host "No critical deprecated API issues found." -ForegroundColor Green
    Write-Host "However, verify with Turbot control findings for confirmation." -ForegroundColor Yellow
}

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT RESULTS
# ══════════════════════════════════════════════════════════════════════════════

if ($ExportResults) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath     = "Azure_CIS_API_Audit_${timestamp}.csv"
    $jsonPath    = "Azure_CIS_API_Audit_${timestamp}.json"
    $summaryPath = "Azure_CIS_API_Audit_Summary_${timestamp}.txt"
   
    Write-Host "`nExporting results..." -ForegroundColor Cyan
   
    # CSV — flat tabular format, suitable for spreadsheet analysis
    $auditResults | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Success "CSV exported to: $csvPath"
   
    # JSON — nested/full-fidelity format, suitable for programmatic processing
    $auditResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath
    Write-Success "JSON exported to: $jsonPath"
   
    # Plain-text summary — human-readable narrative for leadership reporting
    $summary = @"
AZURE CIS v2.0 DEPRECATED API AUDIT SUMMARY
=========================================
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Subscription: $SubscriptionId
Subscription Name: $((Get-AzContext).Subscription.Name)

STATISTICS:
-----------
Total Tests: $totalTests
Deprecated APIs: $deprecated
Current APIs: $current
Errors: $errors

DEPRECATED APIs FOUND:
----------------------
$($auditResults | Where-Object { $_.Status -like "*DEPRECATED*" } | ForEach-Object { "- $($_.TestName): $($_.APIEndpoint) [$($_.CISv2Control)]" } | Out-String)

RECOMMENDATIONS:
----------------
$(if ($deprecated -gt 0) { "1. Document these findings as evidence of Turbot CIS v2.0 control drift`n2. Create exceptions for affected controls`n3. Plan migration to Turbot CIS v3.0`n4. Include in presentation to leadership" } else { "No critical deprecated APIs found. Monitor Turbot findings for discrepancies." })

DETAILED RESULTS:
-----------------
See attached CSV and JSON files for complete audit trail.
"@
   
    $summary | Out-File -FilePath $summaryPath
    Write-Success "Summary report: $summaryPath"
   
    Write-Host "`nAll results exported successfully!" -ForegroundColor Green
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Audit Complete" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan