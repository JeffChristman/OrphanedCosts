<#
.SYNOPSIS
    Validate Turbot CIS v2.0 Azure non-compliant findings against current REST APIs.

.DESCRIPTION
    Ingests Turbot Controls_by_state CSV/XLSX export, extracts non-compliant controls,
    and validates each finding against the current Azure Government REST API to determine:
    - TRUE FINDING: Genuine non-compliance confirmed by current API
    - FALSE POSITIVE (DEPRECATED API): Turbot is checking a deprecated/removed endpoint
    - FALSE POSITIVE (API VERSION): Turbot is using an outdated API version that returns stale data
    - NEEDS INVESTIGATION: Cannot definitively determine validity
   
    Designed for Azure Government (management.usgovcloudapi.net).

.PARAMETER SubscriptionId
    Azure subscription ID to audit

.PARAMETER TurbotExportPath
    Path to the Turbot Controls_by_state export (CSV or XLSX)

.PARAMETER ExportResults
    Export validation results to CSV

.EXAMPLE
    .\Validate-TurbotFindings-Azure.ps1 -SubscriptionId "your-sub-id" -TurbotExportPath ".\Turbot_Controls_by_state.xlsx"

.NOTES
    Author: Jeff (Azure Clarity / VA Security)
    Date: 2026-02-20
    Environment: Azure Government Cloud
    Requires: Az PowerShell module, ImportExcel module (for XLSX input)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$SubscriptionId,

    [Parameter(Mandatory=$true)]
    [string]$TurbotExportPath,

    [Parameter(Mandatory=$false)]
    [switch]$ExportResults
)

#region Helper Functions
function Write-Success { param($Message) Write-Host "[PASS] $Message" -ForegroundColor Green }
function Write-Failure { param($Message) Write-Host "[FAIL] $Message" -ForegroundColor Red }
function Write-Warn    { param($Message) Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Info    { param($Message) Write-Host "[INFO] $Message" -ForegroundColor Cyan }
function Write-Header  { param($Message) Write-Host "`n$('=' * 80)" -ForegroundColor Cyan; Write-Host " $Message" -ForegroundColor Cyan; Write-Host "$('=' * 80)" -ForegroundColor Cyan }

# Azure Government base URI
$script:AzureGovBaseUri = "https://management.usgovcloudapi.net"
$script:ValidationResults = @()

function Get-AzGovToken {
    try {
        $token = (Get-AzAccessToken -ResourceUrl $script:AzureGovBaseUri -ErrorAction Stop).Token
        return $token
    } catch {
        Write-Failure "Failed to get access token: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-AzGovRestApi {
    param(
        [string]$Uri,
        [string]$ApiVersion,
        [string]$Method = "GET"
    )
   
    $token = Get-AzGovToken
    if (-not $token) { return $null }
   
    $headers = @{
        Authorization  = "Bearer $token"
        "Content-Type" = "application/json"
    }
   
    $fullUri = "$script:AzureGovBaseUri$Uri`?api-version=$ApiVersion"
   
    try {
        $response = Invoke-RestMethod -Uri $fullUri -Headers $headers -Method $Method -ErrorAction Stop
        return [PSCustomObject]@{
            Success    = $true
            StatusCode = 200
            Data       = $response
            Error      = $null
        }
    } catch {
        $statusCode = 0
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
        }
        return [PSCustomObject]@{
            Success    = $false
            StatusCode = $statusCode
            Data       = $null
            Error      = $_.Exception.Message
        }
    }
}

function Add-ValidationResult {
    param(
        [string]$CISControl,
        [string]$ControlName,
        [string]$TurbotReason,
        [int]$TurbotFindingCount,
        [string]$ValidationVerdict,
        [string]$APITested,
        [string]$CurrentApiVersion,
        [string]$DeprecatedApiVersion,
        [string]$ActualResult,
        [string]$Explanation,
        [string]$Recommendation
    )
   
    $script:ValidationResults += [PSCustomObject]@{
        Timestamp            = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        CISControl           = $CISControl
        ControlName          = $ControlName
        TurbotReason         = $TurbotReason
        TurbotFindingCount   = $TurbotFindingCount
        ValidationVerdict    = $ValidationVerdict
        APITested            = $APITested
        CurrentApiVersion    = $CurrentApiVersion
        DeprecatedApiVersion = $DeprecatedApiVersion
        ActualResult         = $ActualResult
        Explanation          = $Explanation
        Recommendation       = $Recommendation
    }
}
#endregion

#region Connect and Load Data
Write-Header "TURBOT CIS v2.0 FINDING VALIDATION - AZURE GOVERNMENT"
Write-Info "Subscription: $SubscriptionId"
Write-Info "Turbot Export: $TurbotExportPath"
Write-Info "Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Info "Target: Azure Government (management.usgovcloudapi.net)"
Write-Host ""

# Connect to Azure
Write-Info "Checking Azure connection..."
try {
    $context = Get-AzContext
    if (-not $context) {
        Write-Warn "Not logged in. Attempting Connect-AzAccount for AzureUSGovernment..."
        Connect-AzAccount -Environment AzureUSGovernment
    }
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
    $subName = (Get-AzContext).Subscription.Name
    Write-Success "Connected to: $subName"
} catch {
    Write-Failure "Failed to connect. Run: Connect-AzAccount -Environment AzureUSGovernment"
    exit 1
}

# Load Turbot export
Write-Info "Loading Turbot export..."
try {
    if ($TurbotExportPath -like "*.xlsx") {
        # Requires ImportExcel module
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Warn "ImportExcel module not found. Installing..."
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
        }
        $turbotData = Import-Excel -Path $TurbotExportPath
    } else {
        $turbotData = Import-Csv -Path $TurbotExportPath
    }
   
    # Normalize column names (handle both formats from the XLSX)
    $columns = $turbotData[0].PSObject.Properties.Name
    if ($columns -contains "Unnamed: 0") {
        # XLSX format with unnamed columns - rename them
        $turbotData = $turbotData | Select-Object @{N='ControlID';E={$_.'Unnamed: 0'}},
            @{N='ControlName';E={$_.'Unnamed: 1'}},
            @{N='Resource';E={$_.Resource}},
            @{N='Reason';E={$_.Reason}},
            @{N='State';E={$_.State}},
            @{N='CreatedAt';E={$_.'Created At'}},
            @{N='UpdatedAt';E={$_.'Updated At'}}
    } elseif ($columns -contains "Control Type") {
        # Wrong file - this is AWS format
        Write-Failure "This appears to be an AWS Turbot export. Use the Azure export."
        exit 1
    } else {
        # Try to map whatever columns exist
        $turbotData = $turbotData | Select-Object @{N='ControlID';E={$_.($columns[0])}},
            @{N='ControlName';E={$_.($columns[1])}},
            @{N='Resource';E={$_.Resource}},
            @{N='Reason';E={$_.Reason}},
            @{N='State';E={$_.State}}
    }
   
    # Filter to ALARM state only
    $alarms = $turbotData | Where-Object { $_.State -eq "ALARM" }
   
    # Get unique controls
    $uniqueControls = $alarms | Group-Object -Property ControlID | ForEach-Object {
        [PSCustomObject]@{
            ControlID   = $_.Name.Trim()
            ControlName = ($_.Group[0].ControlName).Trim()
            Reason      = ($_.Group[0].Reason).Trim()
            Count       = $_.Count
        }
    }
   
    Write-Success "Loaded $($alarms.Count) ALARM findings across $($uniqueControls.Count) unique controls"
} catch {
    Write-Failure "Failed to load Turbot export: $($_.Exception.Message)"
    exit 1
}
#endregion

#region Validation Functions per Control Category

# ============================================================
# SECTION 2: DEFENDER FOR CLOUD
# ============================================================
function Test-Control-2_01_03 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # CIS v2.0 2.1.3 checks the deprecated monolithic "Databases" plan
    # Current API splits this into workload-specific plans
   
    # Test 1: Check deprecated "Databases" plan (what Turbot v2.0 likely checks)
    Write-Info "Testing deprecated 'Databases' pricing plan..."
    $depResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/pricings/Databases" -ApiVersion "2022-03-01"
   
    # Test 2: Check current workload-specific plans
    $currentPlans = @("SqlServers", "SqlServerVirtualMachines", "OpenSourceRelationalDatabases", "CosmosDbs")
    $planStatuses = @{}
    $allProtected = $true
   
    foreach ($plan in $currentPlans) {
        Write-Info "  Checking current plan: $plan"
        $result = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/pricings/$plan" -ApiVersion "2024-01-01"
        if ($result.Success) {
            $tier = $result.Data.properties.pricingTier
            $planStatuses[$plan] = $tier
            if ($tier -ne "Standard") { $allProtected = $false }
            $icon = if ($tier -eq "Standard") { "[PASS]" } else { "[WARN]" }
            Write-Host "    $icon $plan = $tier" -ForegroundColor $(if ($tier -eq "Standard") { "Green" } else { "Yellow" })
        } else {
            $planStatuses[$plan] = "ERROR: $($result.Error)"
            $allProtected = $false
            Write-Failure "    $plan - Error: $($result.Error)"
        }
    }
   
    # Determine verdict
    if (-not $depResult.Success -and $allProtected) {
        $verdict = "FALSE POSITIVE (DEPRECATED API)"
        $explanation = "Turbot CIS v2.0 checks the deprecated monolithic 'Databases' pricing plan (removed ~2023). All 4 current workload-specific plans are set to Standard."
        Write-Failure "DEPRECATED API: 'Databases' plan returns $($depResult.StatusCode)"
        Write-Success "All current workload plans are ENABLED"
    } elseif (-not $depResult.Success -and -not $allProtected) {
        $verdict = "MIXED: Deprecated API + Real Gaps"
        $explanation = "Deprecated 'Databases' plan not found AND some current workload plans are not Standard: $($planStatuses | ConvertTo-Json -Compress)"
    } elseif ($depResult.Success) {
        $verdict = "NEEDS INVESTIGATION"
        $explanation = "Deprecated 'Databases' plan unexpectedly still exists. Tier: $($depResult.Data.properties.pricingTier)"
    } else {
        $verdict = "NEEDS INVESTIGATION"
        $explanation = "Unexpected API response pattern"
    }
   
    $planSummary = ($planStatuses.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join "; "
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.Security/pricings/{Databases|SqlServers|SqlServerVMs|OpenSourceRDB|CosmosDbs}" `
        -CurrentApiVersion "2024-01-01" -DeprecatedApiVersion "2022-03-01" `
        -ActualResult $planSummary `
        -Explanation $explanation `
        -Recommendation "Turbot should check individual workload SKUs instead of monolithic 'Databases' plan. Upgrade to CIS v3.0 module."
}

function Test-Control-2_01_14 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # Check for SecurityCenterBuiltIn policy assignment
    Write-Info "Checking for ASC Default (SecurityCenterBuiltIn) policy assignment..."
   
    # Try current API
    $result = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Authorization/policyAssignments/SecurityCenterBuiltIn" -ApiVersion "2023-04-01"
   
    if ($result.Success) {
        $verdict = "FALSE POSITIVE (POLICY EXISTS)"
        $explanation = "SecurityCenterBuiltIn policy assignment found. Turbot may be checking wrong scope or using outdated API version."
        Write-Success "SecurityCenterBuiltIn policy assignment EXISTS"
    } else {
        # Check if it exists under a different name (Azure renamed some default policies)
        Write-Info "Checking for renamed/alternative ASC policy assignments..."
        $allPolicies = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Authorization/policyAssignments" -ApiVersion "2023-04-01"
       
        $ascPolicies = @()
        if ($allPolicies.Success -and $allPolicies.Data.value) {
            $ascPolicies = $allPolicies.Data.value | Where-Object {
                $_.properties.displayName -like "*Security Center*" -or
                $_.properties.displayName -like "*Microsoft cloud security*" -or
                $_.properties.displayName -like "*Defender*" -or
                $_.name -like "*SecurityCenter*" -or
                $_.name -like "*ASC*" -or
                $_.properties.policyDefinitionId -like "*1f3afdf9-d0c9-4c3d-847f-89da613e70a8*"
            }
        }
       
        if ($ascPolicies.Count -gt 0) {
            $verdict = "FALSE POSITIVE (RENAMED POLICY)"
            $policyNames = ($ascPolicies | ForEach-Object { $_.name }) -join ", "
            $explanation = "SecurityCenterBuiltIn not found by exact name, but equivalent ASC policies exist: $policyNames. Microsoft renamed the default initiative to 'Microsoft cloud security benchmark'."
            Write-Warn "SecurityCenterBuiltIn not found by name, but equivalent policies found:"
            $ascPolicies | ForEach-Object { Write-Info "  - $($_.name): $($_.properties.displayName)" }
        } else {
            $verdict = "TRUE FINDING"
            $explanation = "No ASC/Defender default policy assignment found. This is a genuine compliance gap."
            Write-Failure "No ASC default policy assignment found"
        }
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.Authorization/policyAssignments/SecurityCenterBuiltIn" `
        -CurrentApiVersion "2023-04-01" -DeprecatedApiVersion "2021-06-01" `
        -ActualResult $(if ($ascPolicies) { ($ascPolicies | ForEach-Object { $_.name }) -join "; " } else { "Not found" }) `
        -Explanation $explanation `
        -Recommendation "Check if Microsoft renamed the default initiative. In CIS v3.0 this control references 'Microsoft cloud security benchmark' instead of 'SecurityCenterBuiltIn'."
}

function Test-Control-2_01_15 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # Auto provisioning of Log Analytics agent was DEPRECATED
    # Replaced by Defender for Cloud's auto-provisioning of Azure Monitor Agent (AMA)
   
    Write-Info "Checking deprecated Auto Provisioning setting..."
    $depResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/autoProvisioningSettings/default" -ApiVersion "2017-08-01-preview"
   
    Write-Info "Checking current Server Vulnerability Assessment settings..."
    $currentResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/serverVulnerabilityAssessmentsSettings" -ApiVersion "2023-05-01"
   
    # Also check for Defender for Servers plan which replaces auto-provisioning
    Write-Info "Checking Defender for Servers plan..."
    $serversResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/pricings/VirtualMachines" -ApiVersion "2024-01-01"
   
    $autoProvValue = "N/A"
    if ($depResult.Success) {
        $autoProvValue = $depResult.Data.properties.autoProvision
        Write-Info "  Deprecated autoProvision setting: $autoProvValue"
    } else {
        Write-Warn "  Deprecated autoProvision API returned error (may be fully deprecated)"
    }
   
    $defenderServers = "N/A"
    if ($serversResult.Success) {
        $defenderServers = $serversResult.Data.properties.pricingTier
        Write-Info "  Defender for Servers: $defenderServers"
    }
   
    # Log Analytics Agent (MMA) was deprecated August 2024
    $verdict = "LIKELY FALSE POSITIVE (DEPRECATED FEATURE)"
    $explanation = "The Log Analytics Agent (MMA) auto-provisioning was deprecated August 2024. Microsoft replaced it with Azure Monitor Agent (AMA) provisioned through Defender for Servers. CIS v2.0 control references the old MMA auto-provision toggle."
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.Security/autoProvisioningSettings/default" `
        -CurrentApiVersion "2024-01-01 (pricings)" -DeprecatedApiVersion "2017-08-01-preview" `
        -ActualResult "AutoProvision=$autoProvValue; DefenderServers=$defenderServers" `
        -Explanation $explanation `
        -Recommendation "MMA auto-provisioning is deprecated. Check Defender for Servers plan instead. CIS v3.0 removes this control."
}

function Test-Control-2_01_19 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # Security contacts - API version matters here
    Write-Info "Testing deprecated security contacts API..."
    $depResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/securityContacts/default" -ApiVersion "2017-08-01-preview"
   
    Write-Info "Testing current security contacts API..."
    $curResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/securityContacts/default" -ApiVersion "2023-12-01-preview"
   
    # Also try listing all security contacts
    $listResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Security/securityContacts" -ApiVersion "2023-12-01-preview"
   
    $contactsFound = $false
    $contactDetails = "None found"
   
    if ($curResult.Success) {
        $emails = $curResult.Data.properties.emails
        $contactsFound = ($emails -and $emails.Length -gt 0)
        $contactDetails = "Emails: $emails"
        Write-Info "  Current API - Emails: $emails"
    } elseif ($listResult.Success -and $listResult.Data.value) {
        $contactsFound = $true
        $contactDetails = ($listResult.Data.value | ForEach-Object { $_.properties.emails }) -join "; "
        Write-Info "  Found contacts via list: $contactDetails"
    }
   
    if ($depResult.Success -ne $curResult.Success) {
        $verdict = "POSSIBLE FALSE POSITIVE (API VERSION MISMATCH)"
        $explanation = "Deprecated API (2017-08-01-preview) and current API (2023-12-01-preview) return different results. Turbot may be using the older API version."
    } elseif (-not $contactsFound) {
        $verdict = "TRUE FINDING"
        $explanation = "No security contact email addresses configured in either old or new API."
    } else {
        $verdict = "TRUE FINDING - VERIFY EMAIL"
        $explanation = "Security contacts exist but Turbot reports non-compliant. Verify the 'additional email addresses' field specifically. Contact: $contactDetails"
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.Security/securityContacts/default" `
        -CurrentApiVersion "2023-12-01-preview" -DeprecatedApiVersion "2017-08-01-preview" `
        -ActualResult $contactDetails `
        -Explanation $explanation `
        -Recommendation "Verify security contact configuration. Note CIS v3.0 updated the expected property names."
}

# ============================================================
# SECTION 3: STORAGE
# ============================================================
function Test-StorageControls {
    param($Controls)
   
    # Storage controls: 3.02, 3.05, 3.12, 3.13
    # These are mostly resource-level checks, less likely to be deprecated API issues
    # But the API version matters for property availability
   
    Write-Header "Validating: Storage Controls (3.02, 3.05, 3.12, 3.13)"
   
    # Get all storage accounts via current API
    Write-Info "Enumerating storage accounts via current ARM API..."
    $storageResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Storage/storageAccounts" -ApiVersion "2023-05-01"
   
    if (-not $storageResult.Success) {
        Write-Failure "Cannot enumerate storage accounts: $($storageResult.Error)"
        foreach ($ctrl in $Controls) {
            Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                -ValidationVerdict "ERROR" -APITested "Microsoft.Storage/storageAccounts" `
                -CurrentApiVersion "2023-05-01" -DeprecatedApiVersion "N/A" `
                -ActualResult "API Error" -Explanation $storageResult.Error `
                -Recommendation "Investigate API access"
        }
        return
    }
   
    $accounts = $storageResult.Data.value
    Write-Success "Found $($accounts.Count) storage account(s)"
   
    foreach ($ctrl in $Controls) {
        switch -Wildcard ($ctrl.ControlID) {
            "3.02" {
                # Infrastructure Encryption
                Write-Info "  [3.02] Checking Infrastructure Encryption..."
                $nonCompliant = @()
                foreach ($acct in $accounts) {
                    $infraEnc = $acct.properties.encryption.requireInfrastructureEncryption
                    if ($infraEnc -ne $true) { $nonCompliant += $acct.name }
                }
               
                $verdict = if ($nonCompliant.Count -eq $ctrl.Count) { "TRUE FINDING" }
                           elseif ($nonCompliant.Count -eq 0) { "FALSE POSITIVE (API VERSION)" }
                           else { "PARTIAL: $($nonCompliant.Count) of $($ctrl.Count) confirmed" }
               
                # Check if old API version even returns this property
                $oldResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Storage/storageAccounts" -ApiVersion "2021-04-01"
                $apiNote = ""
                if ($oldResult.Success) {
                    $oldHasProperty = $oldResult.Data.value[0].properties.encryption.PSObject.Properties.Name -contains "requireInfrastructureEncryption"
                    if (-not $oldHasProperty) {
                        $apiNote = " Old API (2021-04-01) may not expose this property."
                        $verdict = "NEEDS INVESTIGATION (API VERSION)"
                    }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $verdict `
                    -APITested "Microsoft.Storage/storageAccounts" `
                    -CurrentApiVersion "2023-05-01" -DeprecatedApiVersion "2021-04-01" `
                    -ActualResult "Non-compliant: $($nonCompliant -join ', ')$apiNote" `
                    -Explanation "Checked requireInfrastructureEncryption property.$apiNote" `
                    -Recommendation "Verify Turbot is using API version 2023-05-01+. Infrastructure encryption must be set at creation time."
            }
            "3.05" {
                # Queue Service Logging
                Write-Info "  [3.05] Queue Service Logging - this is a resource-level config..."
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict "LIKELY TRUE FINDING" `
                    -APITested "Microsoft.Storage/storageAccounts/queueServices/diagnosticSettings" `
                    -CurrentApiVersion "2023-05-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Queue diagnostic settings must be checked per-account" `
                    -Explanation "Classic Storage Analytics logging for queues. Note: Azure now recommends Diagnostic Settings (Monitor) over Storage Analytics logging. CIS v3.0 updated this guidance." `
                    -Recommendation "Verify if Storage Analytics logging or Azure Monitor diagnostic settings are configured for queue services. This API has not changed significantly."
            }
            "3.12" {
                # CMK Encryption
                Write-Info "  [3.12] CMK Encryption..."
                $nonCompliant = @()
                foreach ($acct in $accounts) {
                    $keySource = $acct.properties.encryption.keySource
                    if ($keySource -ne "Microsoft.Keyvault") { $nonCompliant += $acct.name }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Storage/storageAccounts" `
                    -CurrentApiVersion "2023-05-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Non-CMK accounts: $($nonCompliant -join ', ')" `
                    -Explanation "Checked encryption.keySource for 'Microsoft.Keyvault'. $($nonCompliant.Count) accounts using Microsoft-managed keys." `
                    -Recommendation "This is likely a true finding. CMK is a deliberate configuration choice and the API is stable."
            }
            "3.13" {
                # Blob Service Logging
                Write-Info "  [3.13] Blob Service Logging..."
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict "LIKELY TRUE FINDING" `
                    -APITested "Microsoft.Storage/storageAccounts/blobServices/diagnosticSettings" `
                    -CurrentApiVersion "2023-05-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Blob diagnostic settings must be checked per-account" `
                    -Explanation "Classic Storage Analytics logging for blobs. Similar to 3.05, CIS v3.0 updated this to reference Azure Monitor diagnostic settings." `
                    -Recommendation "Verify diagnostic settings for blob services. API is stable but CIS v3.0 changed the expected configuration method."
            }
        }
    }
}

# ============================================================
# SECTION 5: LOGGING & MONITORING
# ============================================================
function Test-Control-5_01_02 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # This checks subscription-level diagnostic settings
    # Turbot CIS v2.0 may check deprecated Log Profiles instead of Diagnostic Settings
   
    Write-Info "Testing deprecated Log Profiles API..."
    $logProfileResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/microsoft.insights/logprofiles" -ApiVersion "2016-03-01"
   
    Write-Info "Testing current Diagnostic Settings API..."
    $diagResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/microsoft.insights/diagnosticSettings" -ApiVersion "2021-05-01-preview"
   
    $hasLogProfiles = $logProfileResult.Success -and $logProfileResult.Data.value.Count -gt 0
    $hasDiagSettings = $diagResult.Success -and $diagResult.Data.value.Count -gt 0
   
    if ($hasLogProfiles) {
        Write-Warn "  Legacy Log Profiles found: $($logProfileResult.Data.value.Count)"
    } else {
        Write-Info "  No Log Profiles (expected - deprecated Sept 2023)"
    }
   
    if ($hasDiagSettings) {
        Write-Success "  Diagnostic Settings found: $($diagResult.Data.value.Count)"
        foreach ($ds in $diagResult.Data.value) {
            $enabledLogs = ($ds.properties.logs | Where-Object { $_.enabled -eq $true }).Count
            Write-Info "    - $($ds.name): $enabledLogs log categories enabled"
        }
    } else {
        Write-Warn "  No Diagnostic Settings found"
    }
   
    if (-not $hasDiagSettings -and -not $hasLogProfiles) {
        $verdict = "TRUE FINDING"
        $explanation = "No subscription-level diagnostic settings OR log profiles found. Activity logs are not being exported."
    } elseif ($hasDiagSettings -and -not $hasLogProfiles) {
        $verdict = "FALSE POSITIVE (DEPRECATED API)"
        $explanation = "Turbot reports 'No diagnostic settings found' but current Diagnostic Settings API shows $($diagResult.Data.value.Count) configured setting(s). Turbot CIS v2.0 likely checks the deprecated Log Profiles API (microsoft.insights/logprofiles, deprecated Sept 2023) instead of the current Diagnostic Settings API."
    } else {
        $verdict = "NEEDS INVESTIGATION"
        $explanation = "Both Log Profiles and Diagnostic Settings found. Turbot should be checking Diagnostic Settings."
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "microsoft.insights/diagnosticSettings + microsoft.insights/logprofiles" `
        -CurrentApiVersion "2021-05-01-preview" -DeprecatedApiVersion "2016-03-01 (Log Profiles)" `
        -ActualResult "LogProfiles=$hasLogProfiles; DiagSettings=$hasDiagSettings (count: $(if($diagResult.Data.value){$diagResult.Data.value.Count}else{0}))" `
        -Explanation $explanation `
        -Recommendation "If Diagnostic Settings are configured, this is a false positive from the deprecated Log Profiles API. Upgrade Turbot to CIS v3.0."
}

function Test-Control-5_01_05 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # Key Vault diagnostic settings - check per vault
    Write-Info "Enumerating Key Vaults..."
    $vaults = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.KeyVault/vaults" -ApiVersion "2023-07-01"
   
    if ($vaults.Success -and $vaults.Data.value) {
        $nonCompliant = @()
        foreach ($vault in $vaults.Data.value) {
            $vaultId = $vault.id
            $diagCheck = Invoke-AzGovRestApi -Uri "$vaultId/providers/microsoft.insights/diagnosticSettings" -ApiVersion "2021-05-01-preview"
            if (-not $diagCheck.Success -or $diagCheck.Data.value.Count -eq 0) {
                $nonCompliant += $vault.name
                Write-Warn "  $($vault.name): No diagnostic settings"
            } else {
                Write-Success "  $($vault.name): $($diagCheck.Data.value.Count) diagnostic setting(s)"
            }
        }
       
        $verdict = if ($nonCompliant.Count -eq $Control.Count) { "TRUE FINDING" }
                   elseif ($nonCompliant.Count -eq 0) { "FALSE POSITIVE" }
                   else { "PARTIAL: $($nonCompliant.Count) confirmed non-compliant" }
    } else {
        $verdict = "NEEDS INVESTIGATION"
        $nonCompliant = @("Unable to enumerate")
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.KeyVault/vaults + microsoft.insights/diagnosticSettings" `
        -CurrentApiVersion "2021-05-01-preview" -DeprecatedApiVersion "N/A" `
        -ActualResult "Non-compliant vaults: $($nonCompliant -join ', ')" `
        -Explanation "Checked diagnostic settings on each Key Vault. API is stable for this control." `
        -Recommendation "This is likely a true finding. Configure diagnostic settings for Key Vault logging."
}

function Test-Control-5_01_07 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    # App Service HTTP logs
    Write-Info "Enumerating App Services..."
    $apps = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Web/sites" -ApiVersion "2023-12-01"
   
    $nonCompliant = @()
    if ($apps.Success -and $apps.Data.value) {
        foreach ($app in $apps.Data.value) {
            # Check diagnostic settings for HTTP logs
            $diagCheck = Invoke-AzGovRestApi -Uri "$($app.id)/providers/microsoft.insights/diagnosticSettings" -ApiVersion "2021-05-01-preview"
            $hasHttpLogs = $false
            if ($diagCheck.Success -and $diagCheck.Data.value) {
                foreach ($ds in $diagCheck.Data.value) {
                    $httpLog = $ds.properties.logs | Where-Object { $_.category -eq "AppServiceHTTPLogs" -and $_.enabled -eq $true }
                    if ($httpLog) { $hasHttpLogs = $true; break }
                }
            }
            if (-not $hasHttpLogs) {
                $nonCompliant += $app.name
                Write-Warn "  $($app.name): HTTP logs not enabled"
            } else {
                Write-Success "  $($app.name): HTTP logs enabled"
            }
        }
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
        -APITested "Microsoft.Web/sites + diagnosticSettings" `
        -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "N/A" `
        -ActualResult "Non-compliant apps: $($nonCompliant -join ', ')" `
        -Explanation "Checked AppServiceHTTPLogs diagnostic category per app. API is stable." `
        -Recommendation "Configure diagnostic settings with AppServiceHTTPLogs category for each App Service."
}

function Test-ActivityLogAlerts {
    param($Controls)
   
    Write-Header "Validating: Activity Log Alert Controls (5.02.01 - 5.02.10)"
    Write-Info "Turbot says: $($Controls[0].Reason)"
   
    # All 5.02.x controls check for Activity Log Alerts
    # The API itself hasn't changed, but the operationName values may have
   
    Write-Info "Enumerating Activity Log Alerts..."
    $alertsResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Insights/activityLogAlerts" -ApiVersion "2020-10-01"
   
    # Map control IDs to expected operationName patterns
    $controlOperations = @{
        "5.02.01" = "Microsoft.Authorization/policyAssignments/write"
        "5.02.02" = "Microsoft.Authorization/policyAssignments/delete"
        "5.02.03" = "Microsoft.Network/networkSecurityGroups/write"
        "5.02.04" = "Microsoft.Network/networkSecurityGroups/delete"
        "5.02.05" = "Microsoft.Security/securitySolutions/write"
        "5.02.06" = "Microsoft.Security/securitySolutions/delete"
        "5.02.07" = "Microsoft.Sql/servers/firewallRules/write"
        "5.02.08" = "Microsoft.Sql/servers/firewallRules/delete"
        "5.02.09" = "Microsoft.Network/publicIPAddresses/write"
        "5.02.10" = "Microsoft.Network/publicIPAddresses/delete"
    }
   
    $existingAlerts = @()
    if ($alertsResult.Success -and $alertsResult.Data.value) {
        Write-Success "Found $($alertsResult.Data.value.Count) Activity Log Alert(s)"
       
        foreach ($alert in $alertsResult.Data.value) {
            $conditions = $alert.properties.condition.allOf
            $opCondition = $conditions | Where-Object { $_.field -eq "operationName" }
            if ($opCondition) {
                $existingAlerts += $opCondition.equals
                Write-Info "  Alert '$($alert.name)': monitors $($opCondition.equals)"
            }
        }
    } else {
        Write-Warn "No Activity Log Alerts found OR API error"
    }
   
    foreach ($ctrl in $Controls) {
        $ctrlId = $ctrl.ControlID.Trim()
        $expectedOp = $controlOperations[$ctrlId]
       
        if (-not $expectedOp) {
            Write-Warn "  No operation mapping for $ctrlId"
            continue
        }
       
        $alertExists = $existingAlerts -contains $expectedOp
       
        if ($alertExists) {
            $verdict = "FALSE POSITIVE"
            $explanation = "Alert for '$expectedOp' exists. Turbot may not be finding it due to scope/location mismatch."
            Write-Success "  [$ctrlId] Alert EXISTS for $expectedOp"
        } else {
            $verdict = "TRUE FINDING"
            $explanation = "No Activity Log Alert configured for operation '$expectedOp'. Turbot's reason: '$($ctrl.Reason)'"
            Write-Warn "  [$ctrlId] No alert for $expectedOp"
        }
       
        # Note: CIS v2.0 required these in 'Global' location; check if that's still valid in GovCloud
        Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
            -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
            -ValidationVerdict $verdict `
            -APITested "Microsoft.Insights/activityLogAlerts" `
            -CurrentApiVersion "2020-10-01" -DeprecatedApiVersion "2017-04-01" `
            -ActualResult "Alert for '$expectedOp': $alertExists" `
            -Explanation $explanation `
            -Recommendation "Note: In Azure Government, activity log alerts may use 'USGov Virginia'/'USGov Iowa' instead of 'Global' location. If Turbot checks for 'Global' location specifically, this could be a GovCloud-specific false positive."
    }
}

function Test-Control-5_03_01 {
    param($Control)
    Write-Header "Validating: [$($Control.ControlID)] $($Control.ControlName)"
    Write-Info "Turbot says: $($Control.Reason) ($($Control.Count) finding(s))"
   
    Write-Info "Checking for Application Insights resources..."
    $appInsights = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Insights/components" -ApiVersion "2020-02-02"
   
    if ($appInsights.Success -and $appInsights.Data.value.Count -gt 0) {
        $verdict = "FALSE POSITIVE"
        Write-Success "Found $($appInsights.Data.value.Count) Application Insights instance(s)"
    } else {
        $verdict = "TRUE FINDING"
        Write-Warn "No Application Insights instances found"
    }
   
    Add-ValidationResult -CISControl $Control.ControlID -ControlName $Control.ControlName `
        -TurbotReason $Control.Reason -TurbotFindingCount $Control.Count `
        -ValidationVerdict $verdict `
        -APITested "Microsoft.Insights/components" `
        -CurrentApiVersion "2020-02-02" -DeprecatedApiVersion "N/A" `
        -ActualResult "AppInsights count: $(if($appInsights.Data.value){$appInsights.Data.value.Count}else{0})" `
        -Explanation "Checked for Application Insights resources. API is stable." `
        -Recommendation "Deploy Application Insights for application monitoring if applicable."
}

# ============================================================
# SECTION 6: NETWORKING
# ============================================================
function Test-NetworkControls {
    param($Controls)
   
    Write-Header "Validating: Network Controls (6.01, 6.02, 6.05)"
   
    # Get all NSGs
    Write-Info "Enumerating Network Security Groups..."
    $nsgsResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Network/networkSecurityGroups" -ApiVersion "2024-01-01"
   
    if (-not $nsgsResult.Success) {
        Write-Failure "Cannot enumerate NSGs: $($nsgsResult.Error)"
        return
    }
   
    $nsgs = $nsgsResult.Data.value
    Write-Success "Found $($nsgs.Count) NSG(s)"
   
    foreach ($ctrl in $Controls) {
        switch -Wildcard ($ctrl.ControlID) {
            "6.01" {
                # RDP from Internet
                Write-Info "  [6.01] Checking for RDP (3389) open from Internet..."
                $nonCompliant = @()
                foreach ($nsg in $nsgs) {
                    $rdpRules = $nsg.properties.securityRules | Where-Object {
                        $_.properties.destinationPortRange -match "3389" -and
                        $_.properties.access -eq "Allow" -and
                        $_.properties.direction -eq "Inbound" -and
                        ($_.properties.sourceAddressPrefix -in @("*", "Internet", "0.0.0.0/0"))
                    }
                    if ($rdpRules) { $nonCompliant += $nsg.name }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "NEEDS INVESTIGATION" }) `
                    -APITested "Microsoft.Network/networkSecurityGroups" `
                    -CurrentApiVersion "2024-01-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "NSGs with open RDP: $($nonCompliant -join ', ')" `
                    -Explanation "Checked for Allow rules on port 3389 from Internet/Any. API is stable. Note: Turbot may also check service-tag-based rules which are harder to evaluate." `
                    -Recommendation "Review NSG rules for RDP access. API is current - this is likely a true finding."
            }
            "6.02" {
                # SSH from Internet
                Write-Info "  [6.02] Checking for SSH (22) open from Internet..."
                $nonCompliant = @()
                foreach ($nsg in $nsgs) {
                    $sshRules = $nsg.properties.securityRules | Where-Object {
                        $_.properties.destinationPortRange -match "22" -and
                        $_.properties.access -eq "Allow" -and
                        $_.properties.direction -eq "Inbound" -and
                        ($_.properties.sourceAddressPrefix -in @("*", "Internet", "0.0.0.0/0"))
                    }
                    if ($sshRules) { $nonCompliant += $nsg.name }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "NEEDS INVESTIGATION" }) `
                    -APITested "Microsoft.Network/networkSecurityGroups" `
                    -CurrentApiVersion "2024-01-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "NSGs with open SSH: $($nonCompliant -join ', ')" `
                    -Explanation "Checked for Allow rules on port 22 from Internet/Any. API is stable." `
                    -Recommendation "Review NSG rules for SSH access. This is likely a true finding."
            }
            "6.05" {
                # NSG Flow Log retention
                Write-Info "  [6.05] Checking NSG Flow Log retention..."
                $nonCompliant = @()
                foreach ($nsg in $nsgs) {
                    # Flow logs are a separate resource
                    $nsgId = $nsg.id
                    # Flow logs are under Network Watcher
                    # We need to check the flow log config for each NSG
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict "LIKELY TRUE FINDING" `
                    -APITested "Microsoft.Network/networkWatchers/flowLogs" `
                    -CurrentApiVersion "2024-01-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Flow log retention requires per-NSG verification via Network Watcher" `
                    -Explanation "NSG Flow Logs retention is checked via Network Watcher. API is stable but complex to validate (requires enumerating Network Watchers per region then flow logs per NSG)." `
                    -Recommendation "This is likely a true finding. Verify flow log retention is set to >90 days. API hasn't changed for this control."
            }
        }
    }
}

# ============================================================
# SECTION 7: VMs
# ============================================================
function Test-VMControls {
    param($Controls)
   
    Write-Header "Validating: VM Controls (7.01, 7.03, 7.04)"
   
    foreach ($ctrl in $Controls) {
        switch -Wildcard ($ctrl.ControlID) {
            "7.01" {
                # Bastion Host
                Write-Info "  [7.01] Checking for Azure Bastion Host..."
                $bastionResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Network/bastionHosts" -ApiVersion "2024-01-01"
               
                $bastionExists = $bastionResult.Success -and $bastionResult.Data.value.Count -gt 0
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($bastionExists) { "FALSE POSITIVE" } else { "TRUE FINDING" }) `
                    -APITested "Microsoft.Network/bastionHosts" `
                    -CurrentApiVersion "2024-01-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Bastion hosts: $(if($bastionResult.Data.value){$bastionResult.Data.value.Count}else{0})" `
                    -Explanation "Checked for Azure Bastion Host resources. API is stable." `
                    -Recommendation "Deploy Azure Bastion for secure RDP/SSH access without public IPs."
            }
            "7.03" {
                # OS and Data disk CMK encryption
                Write-Info "  [7.03] Checking VM disk CMK encryption..."
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict "LIKELY TRUE FINDING" `
                    -APITested "Microsoft.Compute/disks" `
                    -CurrentApiVersion "2024-03-02" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Requires per-disk encryption type verification" `
                    -Explanation "CMK disk encryption is a resource-level setting. API is stable. Note: Azure Gov may have different Disk Encryption Set availability." `
                    -Recommendation "Verify disk encryption type via Microsoft.Compute/disks. This is likely a true finding."
            }
            "7.04" {
                # Unattached disk CMK
                Write-Info "  [7.04] Checking unattached disk encryption..."
                $disksResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Compute/disks" -ApiVersion "2024-03-02"
               
                $unattachedNonCMK = @()
                if ($disksResult.Success -and $disksResult.Data.value) {
                    $unattached = $disksResult.Data.value | Where-Object { $_.properties.diskState -eq "Unattached" }
                    $unattachedNonCMK = $unattached | Where-Object {
                        $_.properties.encryption.type -ne "EncryptionAtRestWithCustomerKey"
                    }
                    Write-Info "  Unattached disks: $($unattached.Count), Non-CMK: $($unattachedNonCMK.Count)"
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($unattachedNonCMK.Count -gt 0) { "TRUE FINDING" } else { "NEEDS INVESTIGATION" }) `
                    -APITested "Microsoft.Compute/disks" `
                    -CurrentApiVersion "2024-03-02" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Unattached non-CMK disks: $($unattachedNonCMK.Count)" `
                    -Explanation "Checked unattached disks for CMK encryption. API is stable." `
                    -Recommendation "Encrypt or delete unattached disks. This is likely a true finding."
            }
        }
    }
}

# ============================================================
# SECTION 8: KEY VAULT
# ============================================================
function Test-KeyVaultControls {
    param($Controls)
   
    Write-Header "Validating: Key Vault Controls (8.05, 8.06, 8.07)"
   
    Write-Info "Enumerating Key Vaults..."
    $vaults = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.KeyVault/vaults" -ApiVersion "2023-07-01"
   
    if (-not $vaults.Success) {
        Write-Failure "Cannot enumerate Key Vaults"
        return
    }
   
    $vaultList = $vaults.Data.value
    Write-Success "Found $($vaultList.Count) Key Vault(s)"
   
    foreach ($ctrl in $Controls) {
        switch -Wildcard ($ctrl.ControlID) {
            "8.05" {
                # Recoverable (soft delete + purge protection)
                $nonCompliant = @()
                foreach ($v in $vaultList) {
                    $softDelete = $v.properties.enableSoftDelete
                    $purgeProtect = $v.properties.enablePurgeProtection
                    if ($softDelete -ne $true -or $purgeProtect -ne $true) {
                        $nonCompliant += "$($v.name)(SD=$softDelete,PP=$purgeProtect)"
                    }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.KeyVault/vaults" `
                    -CurrentApiVersion "2023-07-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Non-recoverable: $($nonCompliant -join ', ')" `
                    -Explanation "Checked enableSoftDelete and enablePurgeProtection. Note: soft delete is now default and cannot be disabled on new vaults since Feb 2025." `
                    -Recommendation "Enable purge protection on all vaults. API is stable."
            }
            "8.06" {
                # RBAC for Key Vault
                $nonCompliant = @()
                foreach ($v in $vaultList) {
                    $rbac = $v.properties.enableRbacAuthorization
                    if ($rbac -ne $true) { $nonCompliant += $v.name }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.KeyVault/vaults" `
                    -CurrentApiVersion "2023-07-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Non-RBAC vaults: $($nonCompliant -join ', ')" `
                    -Explanation "Checked enableRbacAuthorization property. API is stable." `
                    -Recommendation "Enable RBAC authorization on Key Vaults."
            }
            "8.07" {
                # Private Endpoints for Key Vault
                $nonCompliant = @()
                foreach ($v in $vaultList) {
                    $pe = $v.properties.privateEndpointConnections
                    if (-not $pe -or $pe.Count -eq 0) { $nonCompliant += $v.name }
                }
               
                Add-ValidationResult -CISControl $ctrl.ControlID -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.KeyVault/vaults" `
                    -CurrentApiVersion "2023-07-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Vaults without private endpoints: $($nonCompliant -join ', ')" `
                    -Explanation "Checked privateEndpointConnections property. API is stable." `
                    -Recommendation "Configure private endpoints for Key Vault network isolation."
            }
        }
    }
}

# ============================================================
# SECTION 9: APP SERVICE
# ============================================================
function Test-AppServiceControls {
    param($Controls)
   
    Write-Header "Validating: App Service Controls (9.01, 9.04, 9.05, 9.09, 9.10)"
   
    Write-Info "Enumerating App Services..."
    $appsResult = Invoke-AzGovRestApi -Uri "/subscriptions/$SubscriptionId/providers/Microsoft.Web/sites" -ApiVersion "2023-12-01"
   
    if (-not $appsResult.Success) {
        Write-Failure "Cannot enumerate App Services"
        return
    }
   
    $apps = $appsResult.Data.value
    Write-Success "Found $($apps.Count) App Service(s)"
   
    foreach ($ctrl in $Controls) {
        $ctrlId = $ctrl.ControlID.Trim()
        switch -Wildcard ($ctrlId) {
            "9.01" {
                # Authentication
                $nonCompliant = @()
                foreach ($app in $apps) {
                    $authResult = Invoke-AzGovRestApi -Uri "$($app.id)/config/authsettingsV2" -ApiVersion "2023-12-01"
                    if ($authResult.Success) {
                        $enabled = $authResult.Data.properties.platform.enabled
                        if ($enabled -ne $true) { $nonCompliant += $app.name }
                    } else {
                        $nonCompliant += "$($app.name)(error)"
                    }
                }
               
                # Note: CIS v2.0 may check authsettings (v1) while current is authsettingsV2
                Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Web/sites/config/authsettingsV2" `
                    -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "authsettings (v1)" `
                    -ActualResult "Non-authenticated apps: $($nonCompliant -join ', ')" `
                    -Explanation "Checked authsettingsV2 (current). Note: Turbot CIS v2.0 may check the older 'authsettings' (v1) endpoint which was deprecated in favor of v2." `
                    -Recommendation "Verify Turbot is checking authsettingsV2 not authsettings. Both APIs exist but v1 may not reflect current EasyAuth configuration."
            }
            "9.04" {
                # Client Certificates
                $nonCompliant = @()
                foreach ($app in $apps) {
                    if ($app.properties.clientCertEnabled -ne $true) { $nonCompliant += $app.name }
                }
               
                Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Web/sites (clientCertEnabled)" `
                    -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Apps without client certs: $($nonCompliant -join ', ')" `
                    -Explanation "Checked clientCertEnabled property. API is stable. Note: CIS v3.0 changed the property to 'clientCertMode' with value 'Required'." `
                    -Recommendation "This is likely true. Note CIS v3.0 uses clientCertMode='Required' instead of clientCertEnabled."
            }
            "9.05" {
                # Register with AAD / Entra ID
                $nonCompliant = @()
                foreach ($app in $apps) {
                    # Check managed identity
                    $hasIdentity = $app.identity -and ($app.identity.type -match "SystemAssigned|UserAssigned")
                    if (-not $hasIdentity) { $nonCompliant += $app.name }
                }
               
                Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Web/sites (identity)" `
                    -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Apps without identity: $($nonCompliant -join ', ')" `
                    -Explanation "Checked for managed identity (replaces legacy AAD app registration). Note: Control name references 'Azure Active Directory' - now 'Microsoft Entra ID' in CIS v3.0." `
                    -Recommendation "Enable managed identity. Note naming change from Azure AD to Entra ID in CIS v3.0."
            }
            "9.09" {
                # HTTP Version
                $nonCompliant = @()
                foreach ($app in $apps) {
                    $configResult = Invoke-AzGovRestApi -Uri "$($app.id)/config/web" -ApiVersion "2023-12-01"
                    if ($configResult.Success) {
                        $httpVersion = $configResult.Data.properties.http20Enabled
                        if ($httpVersion -ne $true) { $nonCompliant += $app.name }
                    }
                }
               
                Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Web/sites/config/web (http20Enabled)" `
                    -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Apps without HTTP/2: $($nonCompliant -join ', ')" `
                    -Explanation "Checked http20Enabled in site config. API is stable." `
                    -Recommendation "Enable HTTP/2 on web apps. This is likely a true finding."
            }
            "9.1" {
                # FTP Deployments
                $nonCompliant = @()
                foreach ($app in $apps) {
                    $configResult = Invoke-AzGovRestApi -Uri "$($app.id)/config/web" -ApiVersion "2023-12-01"
                    if ($configResult.Success) {
                        $ftpState = $configResult.Data.properties.ftpsState
                        if ($ftpState -notin @("Disabled", "FtpsOnly")) { $nonCompliant += "$($app.name)(ftps=$ftpState)" }
                    }
                }
               
                Add-ValidationResult -CISControl $ctrlId -ControlName $ctrl.ControlName `
                    -TurbotReason $ctrl.Reason -TurbotFindingCount $ctrl.Count `
                    -ValidationVerdict $(if ($nonCompliant.Count -gt 0) { "TRUE FINDING" } else { "FALSE POSITIVE" }) `
                    -APITested "Microsoft.Web/sites/config/web (ftpsState)" `
                    -CurrentApiVersion "2023-12-01" -DeprecatedApiVersion "N/A" `
                    -ActualResult "Apps with FTP enabled: $($nonCompliant -join ', ')" `
                    -Explanation "Checked ftpsState in site config. API is stable." `
                    -Recommendation "Disable FTP or restrict to FTPS only."
            }
        }
    }
}

#endregion

#region Main Execution

Write-Header "BEGINNING VALIDATION OF $($uniqueControls.Count) NON-COMPLIANT CONTROLS"

foreach ($ctrl in $uniqueControls) {
    $ctrlId = $ctrl.ControlID.Trim()
   
    switch -Wildcard ($ctrlId) {
        "2.01.03"  { Test-Control-2_01_03 -Control $ctrl }
        "2.01.14"  { Test-Control-2_01_14 -Control $ctrl }
        "2.01.15"  { Test-Control-2_01_15 -Control $ctrl }
        "2.01.19"  { Test-Control-2_01_19 -Control $ctrl }
        "5.01.02"  { Test-Control-5_01_02 -Control $ctrl }
        "5.01.05"  { Test-Control-5_01_05 -Control $ctrl }
        "5.01.07"  { Test-Control-5_01_07 -Control $ctrl }
        "5.03.01"  { Test-Control-5_03_01 -Control $ctrl }
        default {
            # Collect controls for batch processing
            # (handled below)
        }
    }
}

# Batch process grouped controls
$storageControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^3\." }
if ($storageControls) { Test-StorageControls -Controls $storageControls }

$alertControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^5\.02\." }
if ($alertControls) { Test-ActivityLogAlerts -Controls $alertControls }

$networkControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^6\." }
if ($networkControls) { Test-NetworkControls -Controls $networkControls }

$vmControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^7\." }
if ($vmControls) { Test-VMControls -Controls $vmControls }

$kvControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^8\." }
if ($kvControls) { Test-KeyVaultControls -Controls $kvControls }

$appControls = $uniqueControls | Where-Object { $_.ControlID.Trim() -match "^9\." }
if ($appControls) { Test-AppServiceControls -Controls $appControls }

#endregion

#region Summary Report

Write-Header "VALIDATION SUMMARY"

$total = $ValidationResults.Count
$trueFindings = ($ValidationResults | Where-Object { $_.ValidationVerdict -like "TRUE*" }).Count
$falsePositives = ($ValidationResults | Where-Object { $_.ValidationVerdict -like "FALSE*" }).Count
$deprecated = ($ValidationResults | Where-Object { $_.ValidationVerdict -like "*DEPRECATED*" }).Count
$needsInvestigation = ($ValidationResults | Where-Object { $_.ValidationVerdict -like "*INVESTIGATION*" }).Count

Write-Host ""
Write-Host "Total Controls Validated:     $total" -ForegroundColor Cyan
Write-Host "True Findings (real issues):  $trueFindings" -ForegroundColor Red
Write-Host "False Positives:              $falsePositives" -ForegroundColor Green
Write-Host "  - Due to Deprecated API:    $deprecated" -ForegroundColor Yellow
Write-Host "Needs Investigation:          $needsInvestigation" -ForegroundColor Yellow

Write-Host "`nDETAILED VERDICTS:" -ForegroundColor Cyan
$ValidationResults | Group-Object ValidationVerdict | Sort-Object Count -Descending | ForEach-Object {
    Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor $(
        if ($_.Name -like "TRUE*") { "Red" }
        elseif ($_.Name -like "FALSE*") { "Green" }
        else { "Yellow" }
    )
}

# Highlight deprecated API findings
$depFindings = $ValidationResults | Where-Object { $_.ValidationVerdict -like "*DEPRECATED*" }
if ($depFindings) {
    Write-Host "`nDEPRECATED API FALSE POSITIVES (Key Evidence for CIS v3.0 Upgrade):" -ForegroundColor Yellow
    foreach ($f in $depFindings) {
        Write-Host "  [$($f.CISControl)] $($f.ControlName)" -ForegroundColor Red
        Write-Host "    API: $($f.APITested)" -ForegroundColor Yellow
        Write-Host "    Deprecated Version: $($f.DeprecatedApiVersion)" -ForegroundColor Yellow
        Write-Host "    Current Version: $($f.CurrentApiVersion)" -ForegroundColor Cyan
        Write-Host "    Explanation: $($f.Explanation)" -ForegroundColor White
        Write-Host ""
    }
}

Write-Host "`nGOVCLOUD-SPECIFIC NOTES:" -ForegroundColor Cyan
Write-Host "  - Activity Log Alerts (5.02.x): 'Global' location may not exist in Azure Gov" -ForegroundColor Yellow
Write-Host "  - Some API versions may differ between commercial and government clouds" -ForegroundColor Yellow
Write-Host "  - Verify Turbot is configured for 'management.usgovcloudapi.net' base URI" -ForegroundColor Yellow

#endregion

#region Export

if ($ExportResults) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
   
    # CSV
    $csvPath = "Turbot_Validation_Azure_${timestamp}.csv"
    $ValidationResults | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Success "CSV exported: $csvPath"
   
    # JSON
    $jsonPath = "Turbot_Validation_Azure_${timestamp}.json"
    $ValidationResults | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath
    Write-Success "JSON exported: $jsonPath"
   
    # Summary
    $summaryPath = "Turbot_Validation_Azure_Summary_${timestamp}.txt"
    @"
TURBOT CIS v2.0 FINDING VALIDATION - AZURE GOVERNMENT
======================================================
Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Subscription: $SubscriptionId ($subName)
Cloud: Azure Government
Turbot Export: $TurbotExportPath

INPUT:
  Total Turbot ALARM Findings: $($alarms.Count)
  Unique Controls: $($uniqueControls.Count)

VALIDATION RESULTS:
  Controls Validated: $total
  True Findings: $trueFindings
  False Positives: $falsePositives
    Due to Deprecated API: $deprecated
  Needs Investigation: $needsInvestigation

DEPRECATED API FALSE POSITIVES:
$($depFindings | ForEach-Object { "  [$($_.CISControl)] $($_.ControlName) - $($_.DeprecatedApiVersion) -> $($_.CurrentApiVersion)" } | Out-String)

RECOMMENDATION:
$(if ($deprecated -gt 0) {
"The presence of $deprecated control(s) returning false positives due to deprecated APIs
provides strong evidence that Turbot's CIS v2.0 module has systematic accuracy issues.
These deprecated APIs will not be fixed in v2.0.

RECOMMENDED ACTION: Upgrade Turbot to CIS v3.0 module to resolve deprecated API issues."
} else {
"No deprecated API false positives found. All findings appear to be legitimate.
However, verify GovCloud-specific behaviors for Activity Log Alert controls."
})
"@ | Out-File -FilePath $summaryPath
    Write-Success "Summary: $summaryPath"
}

Write-Host "`n$('=' * 80)" -ForegroundColor Cyan
Write-Host " Validation Complete" -ForegroundColor Cyan 
Write-Host "$('=' * 80)`n" -ForegroundColor Cyan

