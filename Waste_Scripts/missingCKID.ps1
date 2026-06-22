#Requires -Modules Az.Compute

<#
.SYNOPSIS
    VM Billing Gap Report — identifies VMs missing CKID (primary) and EMAS tags
    with cost estimates. Designed to feed into CTS systems reports.

.DESCRIPTION
    Scans a selected subscription and outputs every VM with:
      - CKID mapping status (Mapped / UNMAPPED)
      - EMAS number if present
      - VM size, OS, power state
      - Estimated monthly cost (via Azure Retail Prices API)
      - A summary row with totals: % mapped, unmapped VM count, unmapped cost

    Subscription selection modes:
      1. Interactive picker  (default, no params)
      2. By name             (-SubscriptionName, partial match)
      3. By ID               (-SubscriptionId)
      4. VM name search      (-VMName, searches across all subs)

    Output: CSV by default. Add -AsXlsx for Excel format (requires ImportExcel module).
    No Azure CLI — Az PowerShell only.

.PARAMETER SubscriptionId
    Target a specific subscription by ID.

.PARAMETER SubscriptionName
    Target by display name (partial match supported).

.PARAMETER VMName
    Search for a specific VM across all subscriptions.

.PARAMETER OutputPath
    Directory for the output file. Defaults to current directory.

.PARAMETER AsXlsx
    Output as .xlsx instead of .csv (requires ImportExcel module).

.PARAMETER SkipPricing
    Skip pricing lookups for faster runs.

.PARAMETER UnmappedOnly
    Only include VMs missing a CKID tag.

.EXAMPLE
    .\Get-VMBillingGapReport.ps1                                            # Interactive
    .\Get-VMBillingGapReport.ps1 -SubscriptionName "CTS-Prod"              # By name
    .\Get-VMBillingGapReport.ps1 -SubscriptionName "CTS-Prod" -UnmappedOnly  # Only gaps
    .\Get-VMBillingGapReport.ps1 -SubscriptionName "CTS-Prod" -AsXlsx      # Excel output
    .\Get-VMBillingGapReport.ps1 -VMName "DVAGOV-WEB-01"                    # Find one VM
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'ById', Mandatory = $true)]
    [string]$SubscriptionId,

    [Parameter(ParameterSetName = 'ByName', Mandatory = $true)]
    [string]$SubscriptionName,

    [Parameter(ParameterSetName = 'VMSearch', Mandatory = $true)]
    [string]$VMName,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",

    [Parameter(Mandatory = $false)]
    [switch]$AsXlsx,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPricing,

    [Parameter(Mandatory = $false)]
    [switch]$UnmappedOnly
)

# ══════════════════════════════════════════════════════════════════════════════
# FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Confirm-AzGovConnection {
    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Write-Host "[*] Not connected. Logging into Azure Government..." -ForegroundColor Yellow
        Connect-AzAccount -Environment AzureUSGovernment
        $context = Get-AzContext
    }
    elseif ($context.Environment.Name -ne 'AzureUSGovernment') {
        Write-Warning "Current context is '$($context.Environment.Name)', not AzureUSGovernment."
        Write-Host "[*] Reconnecting to Azure Government..." -ForegroundColor Yellow
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Connect-AzAccount -Environment AzureUSGovernment
        $context = Get-AzContext
    }
    Write-Host "[+] Connected as: $($context.Account.Id)" -ForegroundColor Green
    Write-Host "[+] Environment:  $($context.Environment.Name)`n" -ForegroundColor Green
    return $context
}

function Select-Subscription {
    param([array]$Subscriptions)

    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Subscriptions ($($Subscriptions.Count) available)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $pageSize = 25
    $totalPages = [math]::Ceiling($Subscriptions.Count / $pageSize)
    $currentPage = 0

    for ($i = 0; $i -lt $Subscriptions.Count; $i++) {
        $num = $i + 1
        $sub = $Subscriptions[$i]
        $displayName = if ($sub.Name.Length -gt 55) { $sub.Name.Substring(0,52) + "..." } else { $sub.Name }
        Write-Host ("  [{0,3}] {1,-56} {2}" -f $num, $displayName, $sub.Id) -ForegroundColor White

        if (($i + 1) % $pageSize -eq 0 -and ($i + 1) -lt $Subscriptions.Count) {
            $currentPage++
            Write-Host "`n  ── Page $currentPage of $totalPages ── Enter to continue, or type number ──" -ForegroundColor DarkYellow
            $sel = Read-Host "  Selection"
            if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $Subscriptions.Count) {
                return $Subscriptions[[int]$sel - 1]
            }
            Write-Host ""
        }
    }

    Write-Host "`n  [F] Filter by name`n" -ForegroundColor Yellow

    while ($true) {
        $choice = Read-Host "  Enter number (1-$($Subscriptions.Count)) or [F] to filter"

        if ($choice -eq 'F' -or $choice -eq 'f') {
            $filter = Read-Host "  Search term"
            $filtered = @($Subscriptions | Where-Object { $_.Name -like "*$filter*" })
            if ($filtered.Count -eq 0) { Write-Host "  No matches." -ForegroundColor Red; continue }
            if ($filtered.Count -eq 1) {
                Write-Host "  Matched: $($filtered[0].Name)" -ForegroundColor Green
                return $filtered[0]
            }
            Write-Host ""
            for ($i = 0; $i -lt $filtered.Count; $i++) {
                Write-Host ("  [{0,3}] {1}" -f ($i+1), $filtered[$i].Name) -ForegroundColor White
            }
            $pick = Read-Host "`n  Select (1-$($filtered.Count))"
            if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $filtered.Count) {
                return $filtered[[int]$pick - 1]
            }
            Write-Host "  Invalid." -ForegroundColor Red
        }
        elseif ($choice -match '^\d+$') {
            $s = [int]$choice
            if ($s -ge 1 -and $s -le $Subscriptions.Count) { return $Subscriptions[$s - 1] }
            Write-Host "  Out of range." -ForegroundColor Red
        }
        else { Write-Host "  Invalid." -ForegroundColor Red }
    }
}

function Get-TagValue {
    param([hashtable]$Tags, [string[]]$KeyVariants)
    if (-not $Tags) { return $null }
    foreach ($key in $KeyVariants) {
        if ($Tags.ContainsKey($key)) { return $Tags[$key] }
    }
    return $null
}

function Get-VMPricing {
    param([string]$VmSize, [string]$Region)

    $filter = "armRegionName eq '$Region' and armSkuName eq '$VmSize' and priceType eq 'Consumption' and serviceName eq 'Virtual Machines'"
    $uri = "https://prices.azure.com/api/retail/prices?`$filter=$filter"

    try {
        $response = Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 15 -ErrorAction Stop

        $linux = $response.Items | Where-Object {
            $_.productName -notlike '*Windows*' -and $_.productName -notlike '*Spot*' -and
            $_.meterName -notlike '*Spot*' -and $_.meterName -notlike '*Low Priority*'
        } | Select-Object -First 1

        $windows = $response.Items | Where-Object {
            $_.productName -like '*Windows*' -and $_.productName -notlike '*Spot*' -and
            $_.meterName -notlike '*Spot*' -and $_.meterName -notlike '*Low Priority*'
        } | Select-Object -First 1

        return @{
            LinuxMonthly   = if ($linux)   { [math]::Round($linux.retailPrice * 730, 2) }   else { $null }
            WindowsMonthly = if ($windows) { [math]::Round($windows.retailPrice * 730, 2) } else { $null }
        }
    }
    catch {
        Write-Verbose "Pricing failed for $VmSize in $Region : $_"
        return @{ LinuxMonthly = $null; WindowsMonthly = $null }
    }
}

# ── Core: process VMs in a single subscription ───────────────────────────────
function Get-BillingGapData {
    param(
        [object]$Subscription,
        [string]$FilterVMName,
        [switch]$NoPricing,
        [switch]$GapsOnly
    )

    Set-AzContext -SubscriptionId $Subscription.Id -ErrorAction Stop | Out-Null

    if ($FilterVMName) {
        Write-Host "  Searching for '$FilterVMName'..." -ForegroundColor DarkGray -NoNewline
        $vms = @(Get-AzVM -Status -Name $FilterVMName -ErrorAction SilentlyContinue)
        if ($vms.Count -eq 0) { Write-Host " not found." -ForegroundColor DarkGray; return @() }
        Write-Host " FOUND!" -ForegroundColor Green
    }
    else {
        $vms = @(Get-AzVM -Status -ErrorAction SilentlyContinue)
    }

    if ($vms.Count -eq 0) { return @() }

    $pricingCache = @{}
    $inventory = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter = 0

    foreach ($vm in $vms) {
        $counter++
        if (-not $FilterVMName) {
            $pct = [math]::Round(($counter / $vms.Count) * 100)
            Write-Host ("`r  Processing: $counter / $($vms.Count) ($pct%)") -NoNewline -ForegroundColor DarkGray
        }

        # ── Tags ──
        $ckid = Get-TagValue -Tags $vm.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $emass = Get-TagValue -Tags $vm.Tags -KeyVariants @('EMASS','eMASS','emass','Emass','EMASNumber','EMASS_Number','emass_number')
        $vasi = Get-TagValue -Tags $vm.Tags -KeyVariants @('VASI','vasi','Vasi','VASINumber','VASI_Number','vasinumber')

        $ckidStatus = if ($ckid) { 'Mapped' } else { 'UNMAPPED' }

        # ── Governance Status (combined tag health) ──
        $govStatus = if ($ckid -and $emass)     { 'Fully Tagged' }
                     elseif ($ckid -and -not $emass)  { 'Billing Only' }
                     elseif (-not $ckid -and $emass)  { 'ATO Only' }
                     else                              { 'UNTAGGED' }

        # Skip mapped VMs if only showing gaps
        if ($GapsOnly -and $ckid) { continue }

        # ── Pricing ──
        $vmSize = $vm.HardwareProfile.VmSize
        $monthlyEst = $null

        if (-not $NoPricing) {
            $cacheKey = "$vmSize|$($vm.Location)"
            if (-not $pricingCache.ContainsKey($cacheKey)) {
                $pricingCache[$cacheKey] = Get-VMPricing -VmSize $vmSize -Region $vm.Location
            }
            $pricing = $pricingCache[$cacheKey]
            $isWindowsOS = $vm.StorageProfile.OsDisk.OsType -eq 'Windows'
            $monthlyEst = if ($isWindowsOS) { $pricing.WindowsMonthly } else { $pricing.LinuxMonthly }
        }

        # ── Power state (multiple fallback paths) ──
        $powerState = $null

        # Path 1: Direct PowerState property (newer Az module versions)
        if ($vm.PowerState) {
            $powerState = $vm.PowerState
        }

        # Path 2: InstanceView.Statuses (common path)
        if (-not $powerState -and $vm.InstanceView -and $vm.InstanceView.Statuses) {
            $psEntry = $vm.InstanceView.Statuses | Where-Object { $_.Code -like 'PowerState/*' }
            if ($psEntry) { $powerState = $psEntry.DisplayStatus }
        }

        # Path 3: Top-level Statuses array (older Az module versions)
        if (-not $powerState -and $vm.Statuses) {
            $psEntry = $vm.Statuses | Where-Object { $_.Code -like 'PowerState/*' }
            if ($psEntry) { $powerState = $psEntry.DisplayStatus }
        }

        # Path 4: Parse from StatusesText if available
        if (-not $powerState -and $vm.StatusesText) {
            if ($vm.StatusesText -match '"code":\s*"PowerState/(\w+)"') {
                $powerState = "VM $($Matches[1])"
            }
        }

        if (-not $powerState) { $powerState = 'Unknown' }

        $inventory.Add([PSCustomObject]@{
            SubscriptionName  = $Subscription.Name
            SubscriptionId    = $Subscription.Id
            ResourceGroup     = $vm.ResourceGroupName
            VMName            = $vm.Name
            Location          = $vm.Location
            VMSize            = $vmSize
            OSType            = $vm.StorageProfile.OsDisk.OsType
            PowerState        = $powerState
            GovernanceStatus  = $govStatus
            CKID_Status       = $ckidStatus
            CKID              = if ($ckid) { $ckid } else { '' }
            EMASS             = if ($emass) { $emass } else { '' }
            VASI              = if ($vasi) { $vasi } else { '' }
            EstMonthly_USD    = if ($monthlyEst) { $monthlyEst } else { 0 }
            CAM_CMDB_Verified = ''
            Notes             = ''
        })
    }

    if (-not $FilterVMName) { Write-Host "" }
    return $inventory
}

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

$context = Confirm-AzGovConnection
$allSubs = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
Write-Host "[+] $($allSubs.Count) enabled subscription(s)`n" -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# ── Subscription selection ────────────────────────────────────────────────────
switch ($PSCmdlet.ParameterSetName) {

    'VMSearch' {
        Write-Host "══════════════════════════════════════════════════════════" -ForegroundColor Yellow
        Write-Host "  VM Search — '$VMName' across $($allSubs.Count) subs" -ForegroundColor Yellow
        Write-Host "══════════════════════════════════════════════════════════`n" -ForegroundColor Yellow

        $sc = 0
        foreach ($sub in $allSubs) {
            $sc++
            Write-Host ("  [{0,3}/{1}] {2}" -f $sc, $allSubs.Count, $sub.Name) -ForegroundColor White -NoNewline
            $found = Get-BillingGapData -Subscription $sub -FilterVMName $VMName -NoPricing:$SkipPricing
            if ($found.Count -gt 0) {
                $found | ForEach-Object { $results.Add($_) }
                Write-Host "`n  [!] Found — stopping search.`n" -ForegroundColor Green
                break
            }
        }
        if ($results.Count -eq 0) {
            Write-Host "`n  [!] '$VMName' not found in any subscription.`n" -ForegroundColor Red
        }
    }

    'ById' {
        $target = $allSubs | Where-Object { $_.Id -eq $SubscriptionId }
        if (-not $target) { Write-Error "Subscription not found."; return }
        Write-Host "  Scanning: $($target.Name)`n" -ForegroundColor Cyan
        $found = Get-BillingGapData -Subscription $target -NoPricing:$SkipPricing -GapsOnly:$UnmappedOnly
        $found | ForEach-Object { $results.Add($_) }
    }

    'ByName' {
        $target = @($allSubs | Where-Object { $_.Name -like "*$SubscriptionName*" })
        if ($target.Count -eq 0) { Write-Error "No match for '$SubscriptionName'."; return }
        if ($target.Count -gt 1) {
            Write-Host "  Multiple matches:" -ForegroundColor Yellow
            $target = Select-Subscription -Subscriptions $target
        } else { $target = $target[0] }
        Write-Host "  Scanning: $($target.Name)`n" -ForegroundColor Cyan
        $found = Get-BillingGapData -Subscription $target -NoPricing:$SkipPricing -GapsOnly:$UnmappedOnly
        $found | ForEach-Object { $results.Add($_) }
    }

    'Interactive' {
        $selected = Select-Subscription -Subscriptions $allSubs
        Write-Host "`n  Scanning: $($selected.Name)`n" -ForegroundColor Cyan
        $found = Get-BillingGapData -Subscription $selected -NoPricing:$SkipPricing -GapsOnly:$UnmappedOnly
        $found | ForEach-Object { $results.Add($_) }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# OUTPUT + SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

if ($results.Count -eq 0) {
    Write-Host "[!] No VMs to report.`n" -ForegroundColor Yellow
    return
}

# ── Summary stats ─────────────────────────────────────────────────────────────
$totalVMs       = $results.Count
$mappedVMs      = @($results | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
$unmappedVMs    = $totalVMs - $mappedVMs
$mappedPct      = if ($totalVMs -gt 0) { [math]::Round(($mappedVMs / $totalVMs) * 100, 1) } else { 0 }
$unmappedCost   = ($results | Where-Object { $_.CKID_Status -eq 'UNMAPPED' } | Measure-Object -Property EstMonthly_USD -Sum).Sum
if (-not $unmappedCost) { $unmappedCost = 0 }
$totalCost      = ($results | Measure-Object -Property EstMonthly_USD -Sum).Sum
if (-not $totalCost) { $totalCost = 0 }
$hasEmass       = @($results | Where-Object { $_.EMASS -ne '' }).Count

# Governance breakdown
$fullyTagged    = @($results | Where-Object { $_.GovernanceStatus -eq 'Fully Tagged' }).Count
$billingOnly    = @($results | Where-Object { $_.GovernanceStatus -eq 'Billing Only' }).Count
$atoOnly        = @($results | Where-Object { $_.GovernanceStatus -eq 'ATO Only' }).Count
$untagged       = @($results | Where-Object { $_.GovernanceStatus -eq 'UNTAGGED' }).Count

# Power state breakdown
$runningVMs     = @($results | Where-Object { $_.PowerState -like '*running*' }).Count
$deallocatedVMs = @($results | Where-Object { $_.PowerState -like '*deallocated*' }).Count
$stoppedVMs     = @($results | Where-Object { $_.PowerState -like '*stopped*' -and $_.PowerState -notlike '*deallocated*' }).Count
$unknownState   = $totalVMs - $runningVMs - $deallocatedVMs - $stoppedVMs

# Decommission candidates: deallocated/stopped + no CKID
$decommCandidates = @($results | Where-Object {
    $_.CKID_Status -eq 'UNMAPPED' -and
    ($_.PowerState -like '*deallocated*' -or $_.PowerState -like '*stopped*')
}).Count
$decommCost = ($results | Where-Object {
    $_.CKID_Status -eq 'UNMAPPED' -and
    ($_.PowerState -like '*deallocated*' -or $_.PowerState -like '*stopped*')
} | Measure-Object -Property EstMonthly_USD -Sum).Sum
if (-not $decommCost) { $decommCost = 0 }

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  BILLING & GOVERNANCE GAP REPORT" -ForegroundColor Cyan
Write-Host "  NOTE: LADDER frozen — CAM CMDB migration in progress" -ForegroundColor DarkYellow
Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

Write-Host ""
Write-Host "  BILLING (CKID)" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Total VMs:              $totalVMs" -ForegroundColor White
Write-Host "  CKID Mapped:            $mappedVMs ($mappedPct%)" -ForegroundColor Green
$unmappedPctVal = 100 - $mappedPct
$ckidColor = if ($unmappedVMs -gt 0) { 'Red' } else { 'Green' }
Write-Host "  CKID UNMAPPED:          $unmappedVMs ($unmappedPctVal%)" -ForegroundColor $ckidColor

Write-Host ""
Write-Host "  GOVERNANCE STATUS" -ForegroundColor White
Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Fully Tagged (CKID+EMASS): $