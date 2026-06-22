#Requires -Modules Az.Compute, Az.Network

<#
.SYNOPSIS
    CTS Cloud Waste & Governance Report v2.1
    Identifies wasted spend across VMs, orphan disks, snapshots, and public IPs.

.DESCRIPTION
    Comprehensive cost and governance analysis per subscription:

    1. VM ANALYSIS
       - Compute cost (only if running) + Disk cost (always billed)
       - CKID / EMASS / VASI tag governance status
       - Waste flag for deallocated VMs still burning disk costs

    2. ORPHAN DISKS
       - Managed disks with no VM attached (DiskState = Unattached)
       - Leftover from deleted VMs, forgotten data disks
       - Burning storage costs with zero value

    3. ORPHAN SNAPSHOTS
       - Disk snapshots with no corresponding active disk or VM
       - Old backups nobody cleaned up
       - Age threshold configurable via -MaxSnapshotAgeDays (default 90)

    4. ORPHAN PUBLIC IPs
       - Static IPs not associated with any NIC/resource
       - Billing + security risk (reserved but unused)

    Output modes:
      - GUI  (run Get-CloudWasteReport-GUI.ps1)
      - Console with CSV/XLSX export (this script)

    Subscription selection:
      - Interactive picker (default)
      - By name (-SubscriptionName, partial match)
      - By ID (-SubscriptionId)
      - All enabled subscriptions (-AllSubscriptions)

.PARAMETER SubscriptionId
    Target a specific subscription by ID.

.PARAMETER SubscriptionName
    Target by display name (partial match).

.PARAMETER AllSubscriptions
    Scan all enabled subscriptions and combine results into one report.

.PARAMETER OutputPath
    Directory for output files. Defaults to current directory.

.PARAMETER AsXlsx
    Output as .xlsx with multiple worksheets (requires ImportExcel module).

.PARAMETER SkipPricing
    Skip VM compute pricing lookups. Disk/snapshot/IP costs still calculated.

.PARAMETER UnmappedOnly
    Only include VMs missing a CKID tag.

.PARAMETER MaxSnapshotAgeDays
    Age in days after which a snapshot is flagged as waste. Default: 90.

.EXAMPLE
    .\Get-CloudWasteReport.ps1                                        # Interactive
    .\Get-CloudWasteReport.ps1 -SubscriptionName "CTS-Prod"          # By name
    .\Get-CloudWasteReport.ps1 -SubscriptionName "CTS-Prod" -AsXlsx  # Excel multi-sheet
    .\Get-CloudWasteReport.ps1 -AllSubscriptions -AsXlsx             # All subs, Excel
#>

[CmdletBinding(DefaultParameterSetName = 'Interactive')]
param(
    [Parameter(ParameterSetName = 'ById', Mandatory = $true)]
    [string]$SubscriptionId,

    [Parameter(ParameterSetName = 'ByName', Mandatory = $true)]
    [string]$SubscriptionName,

    [Parameter(ParameterSetName = 'All', Mandatory = $true)]
    [switch]$AllSubscriptions,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".",

    [Parameter(Mandatory = $false)]
    [switch]$AsXlsx,

    [Parameter(Mandatory = $false)]
    [switch]$SkipPricing,

    [Parameter(Mandatory = $false)]
    [switch]$UnmappedOnly,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$MaxSnapshotAgeDays = 90
)

$scriptVersion = "2.1"
$reportDate = Get-Date -Format 'yyyy-MM-dd'

# ══════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
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
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Connect-AzAccount -Environment AzureUSGovernment
        $context = Get-AzContext
    }
    Write-Host "[+] Connected as: $($context.Account.Id)" -ForegroundColor Green
    Write-Host "[+] Environment:  $($context.Environment.Name)" -ForegroundColor Green
    Write-Host "[+] Report v$scriptVersion — $reportDate`n" -ForegroundColor Green
    return $context
}

function Select-Subscription {
    param([array]$Subscriptions)
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Subscriptions ($($Subscriptions.Count) available)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $pageSize = 25
    for ($i = 0; $i -lt $Subscriptions.Count; $i++) {
        $num = $i + 1
        $sub = $Subscriptions[$i]
        $displayName = if ($sub.Name.Length -gt 55) { $sub.Name.Substring(0,52) + "..." } else { $sub.Name }
        Write-Host ("  [{0,3}] {1,-56} {2}" -f $num, $displayName, $sub.Id) -ForegroundColor White
        if (($i + 1) % $pageSize -eq 0 -and ($i + 1) -lt $Subscriptions.Count) {
            $sel = Read-Host "`n  Enter to continue, or type number"
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
            if ($filtered.Count -eq 1) { Write-Host "  Matched: $($filtered[0].Name)" -ForegroundColor Green; return $filtered[0] }
            for ($i = 0; $i -lt $filtered.Count; $i++) { Write-Host ("  [{0,3}] {1}" -f ($i+1), $filtered[$i].Name) -ForegroundColor White }
            $pick = Read-Host "`n  Select (1-$($filtered.Count))"
            if ($pick -match '^\d+$' -and [int]$pick -ge 1 -and [int]$pick -le $filtered.Count) { return $filtered[[int]$pick - 1] }
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
    foreach ($key in $KeyVariants) { if ($Tags.ContainsKey($key)) { return $Tags[$key] } }
    return $null
}

function Get-PowerState {
    param([object]$VM)
    $ps = $null
    if ($VM.PowerState) { $ps = $VM.PowerState }
    if (-not $ps -and $VM.InstanceView -and $VM.InstanceView.Statuses) {
        $entry = $VM.InstanceView.Statuses | Where-Object { $_.Code -like 'PowerState/*' }
        if ($entry) { $ps = $entry.DisplayStatus }
    }
    if (-not $ps -and $VM.Statuses) {
        $entry = $VM.Statuses | Where-Object { $_.Code -like 'PowerState/*' }
        if ($entry) { $ps = $entry.DisplayStatus }
    }
    if (-not $ps -and $VM.StatusesText) {
        if ($VM.StatusesText -match '"code":\s*"PowerState/(\w+)"') { $ps = "VM $($Matches[1])" }
    }
    if (-not $ps) { $ps = 'Unknown' }
    return $ps
}

# ══════════════════════════════════════════════════════════════════════════════
# PRICING FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

# Retry wrapper for transient network/API failures
function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [int]$BaseDelaySeconds = 2
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $MaxAttempts) { throw }
            $delay = $BaseDelaySeconds * $attempt
            Write-Verbose "Attempt $attempt failed. Retrying in ${delay}s... ($_)"
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-VMPricing {
    param([string]$VmSize, [string]$Region)
    $filter = "armRegionName eq '$Region' and armSkuName eq '$VmSize' and priceType eq 'Consumption' and serviceName eq 'Virtual Machines'"
    $uri = "https://prices.azure.com/api/retail/prices?`$filter=$filter"
    try {
        $response = Invoke-WithRetry -ScriptBlock {
            Invoke-RestMethod -Uri $uri -Method Get -TimeoutSec 20 -ErrorAction Stop
        }
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
    catch { return @{ LinuxMonthly = $null; WindowsMonthly = $null } }
}

# ── Managed disk tier pricing (Gov cloud approximate rates) ───────────────────
$script:DiskTierRates = @{
    'Standard_LRS'    = @{ 4=0.48; 8=0.96; 16=1.54; 32=1.54; 64=3.01; 128=5.89; 256=11.52; 512=22.53; 1024=43.52; 2048=84.21; 4096=163.84; 8192=327.68; 16384=655.36; 32767=1310.72 }
    'StandardSSD_LRS' = @{ 4=2.40; 8=4.80; 16=9.60; 32=19.20; 64=38.40; 128=76.80; 256=153.60; 512=230.40; 1024=460.80; 2048=921.60; 4096=1843.20 }
    'Premium_LRS'     = @{ 4=5.28; 8=9.60; 16=17.92; 32=34.56; 64=66.56; 128=128.00; 256=245.76; 512=471.04; 1024=901.12; 2048=1720.32; 4096=3276.80 }
    'StandardSSD_ZRS' = @{ 4=3.00; 8=6.00; 16=12.00; 32=24.00; 64=48.00; 128=96.00; 256=192.00; 512=288.00; 1024=576.00 }
    'Premium_ZRS'     = @{ 4=6.60; 8=12.00; 16=22.40; 32=43.20; 64=83.20; 128=160.00; 256=307.20; 512=588.80; 1024=1126.40 }
}
$script:TierSizes = @(4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32767)

function Get-DiskMonthlyCost {
    param([string]$SkuName, [int]$SizeGB)

    # Treat 0-byte disks (unprovisioned) as minimum tier
    if ($SizeGB -le 0) { $SizeGB = 1 }

    $tierSize = $script:TierSizes | Where-Object { $_ -ge $SizeGB } | Select-Object -First 1
    if (-not $tierSize) { $tierSize = 32767 }

    # Try exact SKU match first, then normalize to LRS
    $skuLookup = $SkuName
    if (-not $script:DiskTierRates.ContainsKey($skuLookup)) {
        $skuLookup = $SkuName -replace '_[A-Z]+$', '_LRS'
    }

    if ($script:DiskTierRates.ContainsKey($skuLookup) -and $script:DiskTierRates[$skuLookup].ContainsKey($tierSize)) {
        return $script:DiskTierRates[$skuLookup][$tierSize]
    }

    # Fallback per-GB estimate
    $ratePerGB = switch -Wildcard ($SkuName) {
        'Premium*'     { 0.15 }
        'StandardSSD*' { 0.10 }
        'UltraSSD*'    { 0.12 }
        default        { 0.05 }
    }
    return [math]::Round($SizeGB * $ratePerGB, 2)
}

# Snapshot pricing: roughly per-GB stored
function Get-SnapshotMonthlyCost {
    param([string]$SkuName, [int]$SizeGB)
    if ($SizeGB -le 0) { $SizeGB = 1 }
    $ratePerGB = switch -Wildcard ($SkuName) {
        'Premium*' { 0.10 }
        default    { 0.05 }
    }
    return [math]::Round($SizeGB * $ratePerGB, 2)
}

# Static Public IP: ~$3.65/mo in Gov cloud
$script:StaticIPMonthlyCost = 3.65

# Shared pricing cache — persists across multiple subscription scans
$script:PricingCache = @{}

# ══════════════════════════════════════════════════════════════════════════════
# DATA COLLECTION
# ══════════════════════════════════════════════════════════════════════════════

function Get-AllSubscriptionData {
    param(
        [object]$Subscription,
        [switch]$NoPricing,
        [switch]$UnmappedOnly,
        [int]$MaxSnapshotAgeDays = 90
    )

    Set-AzContext -SubscriptionId $Subscription.Id -ErrorAction Stop | Out-Null
    $subName = $Subscription.Name

    # ── Phase 1: Pull ALL disks in one call and index by ID ───────────────
    Write-Progress -Activity "[$subName] Loading resources" -Status "Managed disks..." -PercentComplete 10
    $allDisks = @(Get-AzDisk -ErrorAction SilentlyContinue)
    $diskIndex = @{}
    foreach ($d in $allDisks) { $diskIndex[$d.Id] = $d }
    Write-Host "  [1/5] Disks:       $($allDisks.Count)" -ForegroundColor DarkGray

    # ── Phase 2: Pull ALL snapshots ───────────────────────────────────────
    Write-Progress -Activity "[$subName] Loading resources" -Status "Snapshots..." -PercentComplete 25
    $allSnapshots = @(Get-AzSnapshot -ErrorAction SilentlyContinue)
    Write-Host "  [2/5] Snapshots:   $($allSnapshots.Count)" -ForegroundColor DarkGray

    # ── Phase 3: Pull ALL public IPs ──────────────────────────────────────
    Write-Progress -Activity "[$subName] Loading resources" -Status "Public IPs..." -PercentComplete 40
    $allPIPs = @(Get-AzPublicIpAddress -ErrorAction SilentlyContinue)
    Write-Host "  [3/5] Public IPs:  $($allPIPs.Count)" -ForegroundColor DarkGray

    # ── Phase 4: Pull ALL VMs ─────────────────────────────────────────────
    Write-Progress -Activity "[$subName] Loading resources" -Status "VMs with status..." -PercentComplete 55
    $allVMs = @(Get-AzVM -Status -ErrorAction SilentlyContinue)
    Write-Host "  [4/5] VMs:         $($allVMs.Count)" -ForegroundColor DarkGray

    # ── Phase 5: Process everything ───────────────────────────────────────
    Write-Host "  [5/5] Analyzing..." -ForegroundColor DarkGray

    # Track which disk IDs are claimed by any VM (running or deallocated)
    $attachedDiskIds = @{}

    # ────────────────────── VM PROCESSING ──────────────────────────────────
    $vmResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $vmCounter = 0
    $vmTotal   = $allVMs.Count

    foreach ($vm in $allVMs) {
        $vmCounter++
        $pct = if ($vmTotal -gt 0) { [math]::Round(($vmCounter / $vmTotal) * 100) } else { 100 }
        Write-Progress -Activity "[$subName] Analyzing VMs" -Status "$vmCounter / $vmTotal" -PercentComplete $pct

        # Tags
        $ckid  = Get-TagValue -Tags $vm.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $emass = Get-TagValue -Tags $vm.Tags -KeyVariants @('EMASS','eMASS','emass','Emass','EMASNumber','EMASS_Number','emass_number')
        $vasi  = Get-TagValue -Tags $vm.Tags -KeyVariants @('VASI','vasi','Vasi','VASINumber','VASI_Number','vasinumber')

        $ckidStatus = if ($ckid) { 'Mapped' } else { 'UNMAPPED' }
        $govStatus = if ($ckid -and $emass)          { 'Fully Tagged' }
                     elseif ($ckid -and -not $emass)  { 'Billing Only' }
                     elseif (-not $ckid -and $emass)  { 'ATO Only' }
                     else                              { 'UNTAGGED' }

        if ($UnmappedOnly -and $ckid) { continue }

        # Compute pricing — use shared cross-subscription cache
        $vmSize = $vm.HardwareProfile.VmSize
        $computeCost = 0
        if (-not $NoPricing) {
            $cacheKey = "$vmSize|$($vm.Location)"
            if (-not $script:PricingCache.ContainsKey($cacheKey)) {
                $script:PricingCache[$cacheKey] = Get-VMPricing -VmSize $vmSize -Region $vm.Location
            }
            $pricing = $script:PricingCache[$cacheKey]
            $isWindowsOS = $vm.StorageProfile.OsDisk.OsType -eq 'Windows'
            $price = if ($isWindowsOS) { $pricing.WindowsMonthly } else { $pricing.LinuxMonthly }
            if ($price) { $computeCost = $price }
        }

        # Disk costs from pre-loaded index (no extra API calls)
        $diskCost = 0
        $diskCount = 0
        $totalDiskGB = 0
        $diskDetailList = [System.Collections.Generic.List[string]]::new()

        # OS Disk
        $osDiskId = $vm.StorageProfile.OsDisk.ManagedDisk.Id
        if ($osDiskId) {
            $attachedDiskIds[$osDiskId] = $true
            $osDisk = $diskIndex[$osDiskId]
            if ($osDisk) {
                $diskCount++
                $totalDiskGB += $osDisk.DiskSizeGB
                $cost = Get-DiskMonthlyCost -SkuName $osDisk.Sku.Name -SizeGB $osDisk.DiskSizeGB
                $diskCost += $cost
                $diskDetailList.Add("OS:$($osDisk.Sku.Name):$($osDisk.DiskSizeGB)GB=`$$cost")
            }
        }

        # Data Disks
        foreach ($dd in $vm.StorageProfile.DataDisks) {
            if ($dd.ManagedDisk.Id) {
                $attachedDiskIds[$dd.ManagedDisk.Id] = $true
                $dataDisk = $diskIndex[$dd.ManagedDisk.Id]
                if ($dataDisk) {
                    $diskCount++
                    $totalDiskGB += $dataDisk.DiskSizeGB
                    $cost = Get-DiskMonthlyCost -SkuName $dataDisk.Sku.Name -SizeGB $dataDisk.DiskSizeGB
                    $diskCost += $cost
                    $diskDetailList.Add("Data:$($dataDisk.Sku.Name):$($dataDisk.DiskSizeGB)GB=`$$cost")
                }
            }
        }

        # Power state + cost classification
        $powerState = Get-PowerState -VM $vm
        $isRunning = $powerState -like '*running*'
        $isDeallocated = $powerState -like '*deallocated*' -or $powerState -like '*stopped*'
        $actualCost = if ($isRunning) { $computeCost + $diskCost } else { $diskCost }
        $wasteFlag = if ($isDeallocated) { 'WASTE' } elseif (-not $isRunning -and -not $isDeallocated) { 'CHECK' } else { '' }

        $vmResults.Add([PSCustomObject]@{
            ResourceType      = 'VM'
            SubscriptionName  = $subName
            ResourceGroup     = $vm.ResourceGroupName
            Name              = $vm.Name
            Location          = $vm.Location
            VMSize            = $vmSize
            OSType            = $vm.StorageProfile.OsDisk.OsType
            PowerState        = $powerState
            GovernanceStatus  = $govStatus
            CKID_Status       = $ckidStatus
            CKID              = if ($ckid) { $ckid } else { '' }
            EMASS             = if ($emass) { $emass } else { '' }
            VASI              = if ($vasi) { $vasi } else { '' }
            DiskCount         = $diskCount
            TotalDiskGB       = $totalDiskGB
            ComputeCost_Mo    = [math]::Round($computeCost, 2)
            DiskCost_Mo       = [math]::Round($diskCost, 2)
            ActualCost_Mo     = [math]::Round($actualCost, 2)
            WasteFlag         = $wasteFlag
            Details           = ($diskDetailList -join ' | ')
        })
    }
    Write-Progress -Activity "[$subName] Analyzing VMs" -Completed

    # ────────────────────── ORPHAN DISKS ───────────────────────────────────
    Write-Progress -Activity "[$subName] Identifying orphans" -Status "Orphan disks..." -PercentComplete 33
    $orphanDiskResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($disk in $allDisks) {
        # Skip disks claimed by any VM (running or deallocated)
        if ($attachedDiskIds.ContainsKey($disk.Id)) { continue }
        # Skip disks the API still considers attached or reserved (belt-and-suspenders)
        if ($disk.DiskState -eq 'Attached' -or $disk.DiskState -eq 'Reserved') { continue }

        $cost = Get-DiskMonthlyCost -SkuName $disk.Sku.Name -SizeGB $disk.DiskSizeGB
        $ckid = Get-TagValue -Tags $disk.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')

        $createdDate = $disk.TimeCreated
        $ageInDays = if ($createdDate) { [math]::Round(((Get-Date) - $createdDate).TotalDays) } else { -1 }

        $orphanDiskResults.Add([PSCustomObject]@{
            ResourceType     = 'Orphan Disk'
            SubscriptionName = $subName
            ResourceGroup    = $disk.ResourceGroupName
            Name             = $disk.Name
            Location         = $disk.Location
            DiskSku          = $disk.Sku.Name
            DiskSizeGB       = $disk.DiskSizeGB
            DiskState        = $disk.DiskState
            CKID             = if ($ckid) { $ckid } else { '' }
            AgeDays          = $ageInDays
            EstCost_Mo       = $cost
            WasteFlag        = 'WASTE'
            Details          = "Unattached $($disk.Sku.Name) disk, $($disk.DiskSizeGB)GB, $ageInDays days old"
        })
    }
    Write-Host "        Orphan disks:  $($orphanDiskResults.Count)" -ForegroundColor $(if ($orphanDiskResults.Count -gt 0) { 'Yellow' } else { 'DarkGray' })

    # ────────────────────── ORPHAN SNAPSHOTS ───────────────────────────────
    Write-Progress -Activity "[$subName] Identifying orphans" -Status "Snapshots..." -PercentComplete 66
    $orphanSnapResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    # Build set of active disk IDs for cross-reference
    $activeDiskIds = @{}
    foreach ($d in $allDisks) { $activeDiskIds[$d.Id] = $true }

    foreach ($snap in $allSnapshots) {
        $cost = Get-SnapshotMonthlyCost -SkuName $snap.Sku.Name -SizeGB $snap.DiskSizeGB

        $createdDate = $snap.TimeCreated
        $ageInDays = if ($createdDate) { [math]::Round(((Get-Date) - $createdDate).TotalDays) } else { -1 }

        # Check if the source disk still exists
        $sourceDiskId = $snap.CreationData.SourceResourceId
        $sourceDiskExists = if ($sourceDiskId) { $activeDiskIds.ContainsKey($sourceDiskId) } else { $false }

        $ckid = Get-TagValue -Tags $snap.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')

        $isWaste = (-not $sourceDiskExists) -or ($ageInDays -ge 0 -and $ageInDays -gt $MaxSnapshotAgeDays)

        $orphanSnapResults.Add([PSCustomObject]@{
            ResourceType      = 'Snapshot'
            SubscriptionName  = $subName
            ResourceGroup     = $snap.ResourceGroupName
            Name              = $snap.Name
            Location          = $snap.Location
            SnapshotSku       = $snap.Sku.Name
            DiskSizeGB        = $snap.DiskSizeGB
            SourceDiskExists  = $sourceDiskExists
            CKID              = if ($ckid) { $ckid } else { '' }
            AgeDays           = $ageInDays
            EstCost_Mo        = $cost
            WasteFlag         = if ($isWaste) { 'WASTE' } else { 'REVIEW' }
            Details           = "Snapshot $($snap.DiskSizeGB)GB, $ageInDays days old$(if (-not $sourceDiskExists) { ', SOURCE DISK DELETED' })"
        })
    }
    $snapWasteCount = @($orphanSnapResults | Where-Object { $_.WasteFlag -eq 'WASTE' }).Count
    Write-Host "        Snapshots:     $($orphanSnapResults.Count) total, $snapWasteCount waste" -ForegroundColor $(if ($orphanSnapResults.Count -gt 0) { 'Yellow' } else { 'DarkGray' })

    # ────────────────────── ORPHAN PUBLIC IPs ──────────────────────────────
    Write-Progress -Activity "[$subName] Identifying orphans" -Status "Public IPs..." -PercentComplete 90
    $orphanIPResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($pip in $allPIPs) {
        # Orphaned = no NIC association and not attached to a NAT gateway
        $isOrphaned = -not $pip.IpConfiguration -and -not $pip.NatGateway

        if (-not $isOrphaned) { continue }

        $cost = if ($pip.PublicIpAllocationMethod -eq 'Static') { $script:StaticIPMonthlyCost } else { 0 }
        $ckid = Get-TagValue -Tags $pip.Tag -KeyVariants @('CKID','ckid','CkId','Ckid')

        $orphanIPResults.Add([PSCustomObject]@{
            ResourceType      = 'Public IP'
            SubscriptionName  = $subName
            ResourceGroup     = $pip.ResourceGroupName
            Name              = $pip.Name
            Location          = $pip.Location
            IPAddress         = if ($pip.IpAddress) { $pip.IpAddress } else { 'Not assigned' }
            AllocationMethod  = $pip.PublicIpAllocationMethod
            Sku               = $pip.Sku.Name
            CKID              = if ($ckid) { $ckid } else { '' }
            EstCost_Mo        = $cost
            WasteFlag         = if ($cost -gt 0) { 'WASTE' } else { 'REVIEW' }
            Details           = "$($pip.PublicIpAllocationMethod) $($pip.Sku.Name) IP, not associated to any resource"
        })
    }
    Write-Host "        Orphan IPs:    $($orphanIPResults.Count)" -ForegroundColor $(if ($orphanIPResults.Count -gt 0) { 'Yellow' } else { 'DarkGray' })

    Write-Progress -Activity "[$subName] Identifying orphans" -Completed

    return @{
        VMs           = $vmResults
        OrphanDisks   = $orphanDiskResults
        Snapshots     = $orphanSnapResults
        OrphanIPs     = $orphanIPResults
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY HELPER
# ══════════════════════════════════════════════════════════════════════════════

function Write-WasteSummary {
    param(
        [object[]]$Vms,
        [object[]]$OrphanDisks,
        [object[]]$Snapshots,
        [object[]]$OrphanIPs,
        [string]$ScopeName
    )

    $totalVMs      = $Vms.Count
    $mappedVMs     = @($Vms | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
    $unmappedVMs   = $totalVMs - $mappedVMs
    $mappedPct     = if ($totalVMs -gt 0) { [math]::Round(($mappedVMs / $totalVMs) * 100, 1) } else { 0 }

    $fullyTagged   = @($Vms | Where-Object { $_.GovernanceStatus -eq 'Fully Tagged' }).Count
    $billingOnly   = @($Vms | Where-Object { $_.GovernanceStatus -eq 'Billing Only' }).Count
    $atoOnly       = @($Vms | Where-Object { $_.GovernanceStatus -eq 'ATO Only' }).Count
    $untaggedVMs   = @($Vms | Where-Object { $_.GovernanceStatus -eq 'UNTAGGED' }).Count

    $runningVMs    = @($Vms | Where-Object { $_.PowerState -like '*running*' }).Count
    $deallocVMs    = @($Vms | Where-Object { $_.PowerState -like '*deallocated*' }).Count

    $vmActualCost  = [double](($Vms | Measure-Object -Property ActualCost_Mo -Sum).Sum)
    $vmComputeCost = [double](($Vms | Measure-Object -Property ComputeCost_Mo -Sum).Sum)
    $vmDiskCost    = [double](($Vms | Measure-Object -Property DiskCost_Mo -Sum).Sum)

    $vmWasteVMs    = @($Vms | Where-Object { $_.WasteFlag -eq 'WASTE' })
    $vmWasteCost   = [double](($vmWasteVMs | Measure-Object -Property ActualCost_Mo -Sum).Sum)

    $orphanDiskCost = [double](($OrphanDisks | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $snapWasteItems = @($Snapshots | Where-Object { $_.WasteFlag -eq 'WASTE' })
    $snapWasteCost  = [double](($snapWasteItems | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $snapTotalCost  = [double](($Snapshots | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $orphanIPCost   = [double](($OrphanIPs | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum)

    $totalWaste    = $vmWasteCost + $orphanDiskCost + $snapWasteCost + $orphanIPCost
    $totalSpend    = $vmActualCost + $orphanDiskCost + $snapTotalCost + $orphanIPCost

    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  CTS CLOUD WASTE & GOVERNANCE REPORT v$script:scriptVersion" -ForegroundColor Cyan
    Write-Host "  $ScopeName" -ForegroundColor White
    Write-Host "  $script:reportDate  |  LADDER FROZEN — CAM CMDB Migration" -ForegroundColor DarkYellow
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    Write-Host "`n  VM GOVERNANCE ($totalVMs VMs)" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  CKID Mapped:            $mappedVMs ($mappedPct%)" -ForegroundColor Green
    Write-Host "  CKID UNMAPPED:          $unmappedVMs" -ForegroundColor $(if ($unmappedVMs -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Fully Tagged:           $fullyTagged  |  Billing Only: $billingOnly" -ForegroundColor DarkGray
    Write-Host "  ATO Only:               $atoOnly  |  UNTAGGED: $untaggedVMs" -ForegroundColor DarkGray

    Write-Host "`n  POWER STATE" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Running: $runningVMs  |  Deallocated: $deallocVMs  |  Other: $($totalVMs - $runningVMs - $deallocVMs)" -ForegroundColor DarkGray

    Write-Host "`n  VM COSTS" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ("  Compute (running):      `${0:N2}" -f $vmComputeCost) -ForegroundColor White
    Write-Host ("  Disk (all VMs):         `${0:N2}" -f $vmDiskCost) -ForegroundColor White
    Write-Host ("  VM Total Actual:        `${0:N2}" -f $vmActualCost) -ForegroundColor White
    Write-Host ("  VM Waste (dealloc):     `${0:N2}  ({1} VMs)" -f $vmWasteCost, $vmWasteVMs.Count) -ForegroundColor $(if ($vmWasteCost -gt 0) { 'Red' } else { 'Green' })

    Write-Host "`n  ORPHAN RESOURCES" -ForegroundColor White
    Write-Host "  ─────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ("  Orphan Disks:           {0} disks  (`${1:N2}/mo)" -f $OrphanDisks.Count, $orphanDiskCost) -ForegroundColor $(if ($OrphanDisks.Count -gt 0) { 'Red' } else { 'Green' })
    Write-Host ("  Snapshots (waste):      {0} of {1}  (`${2:N2}/mo)" -f $snapWasteItems.Count, $Snapshots.Count, $snapWasteCost) -ForegroundColor $(if ($snapWasteItems.Count -gt 0) { 'Yellow' } else { 'Green' })
    Write-Host ("  Orphan Public IPs:      {0}  (`${1:N2}/mo)" -f $OrphanIPs.Count, $orphanIPCost) -ForegroundColor $(if ($OrphanIPs.Count -gt 0) { 'Yellow' } else { 'Green' })

    Write-Host "`n  ═════════════════════════════════════════" -ForegroundColor Red
    Write-Host ("  TOTAL MONTHLY WASTE:    `${0:N2}" -f $totalWaste) -ForegroundColor $(if ($totalWaste -gt 0) { 'Red' } else { 'Green' })
    Write-Host ("  TOTAL MONTHLY SPEND:    `${0:N2}" -f $totalSpend) -ForegroundColor White
    Write-Host "  ═════════════════════════════════════════`n" -ForegroundColor Red

    # Top waste items
    if ($vmWasteVMs.Count -gt 0) {
        Write-Host "  Top Wasted VMs (deallocated, burning disk cost):" -ForegroundColor Red
        $vmWasteVMs | Sort-Object ActualCost_Mo -Descending | Select-Object -First 10 |
            Format-Table -AutoSize -Property Name, VMSize, CKID_Status, DiskCount, @{
                Name='DiskGB'; Expression={ $_.TotalDiskGB }
            }, @{
                Name='Waste$/Mo'; Expression={ '$' + ('{0:N2}' -f $_.DiskCost_Mo) }
            }, ResourceGroup
    }

    if ($OrphanDisks.Count -gt 0) {
        Write-Host "  Top Orphan Disks (unattached, no VM):" -ForegroundColor Red
        $OrphanDisks | Sort-Object EstCost_Mo -Descending | Select-Object -First 10 |
            Format-Table -AutoSize -Property Name, DiskSku, DiskSizeGB, AgeDays, CKID, @{
                Name='Waste$/Mo'; Expression={ '$' + ('{0:N2}' -f $_.EstCost_Mo) }
            }, ResourceGroup
    }

    return @{
        TotalWaste  = $totalWaste
        TotalSpend  = $totalSpend
        VMCount     = $totalVMs
        OrphanCount = $OrphanDisks.Count + $snapWasteItems.Count + $OrphanIPs.Count
    }
}

# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

$context = Confirm-AzGovConnection
$allSubs = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
Write-Host "[+] $($allSubs.Count) enabled subscription(s)`n" -ForegroundColor Cyan

# Subscription selection
$targetSubs = @()
switch ($PSCmdlet.ParameterSetName) {
    'ById' {
        $found = $allSubs | Where-Object { $_.Id -eq $SubscriptionId }
        if (-not $found) { Write-Error "Subscription not found."; return }
        $targetSubs = @($found)
    }
    'ByName' {
        $matches = @($allSubs | Where-Object { $_.Name -like "*$SubscriptionName*" })
        if ($matches.Count -eq 0) { Write-Error "No match for '$SubscriptionName'."; return }
        if ($matches.Count -gt 1) { $targetSubs = @(Select-Subscription -Subscriptions $matches) }
        else { $targetSubs = $matches }
    }
    'All' {
        $targetSubs = $allSubs
        Write-Host "[*] Scanning all $($allSubs.Count) subscriptions`n" -ForegroundColor Yellow
    }
    'Interactive' {
        $targetSubs = @(Select-Subscription -Subscriptions $allSubs)
    }
}

# Accumulate results across all target subscriptions
$allVms        = [System.Collections.Generic.List[PSCustomObject]]::new()
$allOrphanDisks = [System.Collections.Generic.List[PSCustomObject]]::new()
$allSnapshots  = [System.Collections.Generic.List[PSCustomObject]]::new()
$allOrphanIPs  = [System.Collections.Generic.List[PSCustomObject]]::new()

$subIndex = 0
foreach ($targetSub in $targetSubs) {
    $subIndex++
    Write-Host "`n  [$subIndex/$($targetSubs.Count)] Scanning: $($targetSub.Name)" -ForegroundColor Cyan
    Write-Host "  ═══════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

    $data = Get-AllSubscriptionData `
        -Subscription      $targetSub `
        -NoPricing:        $SkipPricing `
        -UnmappedOnly:     $UnmappedOnly `
        -MaxSnapshotAgeDays $MaxSnapshotAgeDays

    foreach ($r in $data.VMs)         { $allVms.Add($r) }
    foreach ($r in $data.OrphanDisks) { $allOrphanDisks.Add($r) }
    foreach ($r in $data.Snapshots)   { $allSnapshots.Add($r) }
    foreach ($r in $data.OrphanIPs)   { $allOrphanIPs.Add($r) }
}

# ══════════════════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

$scopeName = if ($targetSubs.Count -eq 1) { $targetSubs[0].Name } else { "All Subscriptions ($($targetSubs.Count))" }

$summary = Write-WasteSummary `
    -Vms         $allVms `
    -OrphanDisks $allOrphanDisks `
    -Snapshots   $allSnapshots `
    -OrphanIPs   $allOrphanIPs `
    -ScopeName   $scopeName

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════

$subLabel = ($targetSubs | Select-Object -First 1).Name -replace '[^a-zA-Z0-9\-]', '_'
if ($targetSubs.Count -gt 1) { $subLabel = "AllSubscriptions" }

if ($AsXlsx) {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Warning "ImportExcel module not found. Falling back to CSV."
        $AsXlsx = $false
    }
}

if ($AsXlsx) {
    $outFile = Join-Path $OutputPath "CloudWaste_${subLabel}_${reportDate}.xlsx"

    if ($allVms.Count -gt 0) {
        $allVms | Export-Excel -Path $outFile -WorksheetName "VMs" -AutoSize -AutoFilter -FreezeTopRow -ConditionalText @(
            $(New-ConditionalText -Text "WASTE"       -BackgroundColor '#FFCCCC' -ConditionalTextColor '#CC0000')
            $(New-ConditionalText -Text "UNTAGGED"    -BackgroundColor '#FFCCCC' -ConditionalTextColor '#CC0000')
            $(New-ConditionalText -Text "Fully Tagged" -BackgroundColor '#D4EDDA' -ConditionalTextColor '#155724')
        )
    }
    if ($allOrphanDisks.Count -gt 0) {
        $allOrphanDisks | Export-Excel -Path $outFile -WorksheetName "Orphan Disks" -AutoSize -AutoFilter -FreezeTopRow
    }
    if ($allSnapshots.Count -gt 0) {
        $allSnapshots | Export-Excel -Path $outFile -WorksheetName "Snapshots" -AutoSize -AutoFilter -FreezeTopRow
    }
    if ($allOrphanIPs.Count -gt 0) {
        $allOrphanIPs | Export-Excel -Path $outFile -WorksheetName "Orphan IPs" -AutoSize -AutoFilter -FreezeTopRow
    }

    # Summary sheet — one row per subscription scanned
    $summaryRows = foreach ($sub in $targetSubs) {
        $subVms  = @($allVms        | Where-Object { $_.SubscriptionName -eq $sub.Name })
        $subDisks = @($allOrphanDisks | Where-Object { $_.SubscriptionName -eq $sub.Name })
        $subSnaps = @($allSnapshots  | Where-Object { $_.SubscriptionName -eq $sub.Name })
        $subIPs   = @($allOrphanIPs  | Where-Object { $_.SubscriptionName -eq $sub.Name })
        [PSCustomObject]@{
            Subscription      = $sub.Name
            ReportDate        = $reportDate
            Version           = "v$scriptVersion"
            TotalVMs          = $subVms.Count
            CKID_Mapped       = @($subVms | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
            CKID_Unmapped     = @($subVms | Where-Object { $_.CKID_Status -eq 'UNMAPPED' }).Count
            FullyTagged       = @($subVms | Where-Object { $_.GovernanceStatus -eq 'Fully Tagged' }).Count
            Running           = @($subVms | Where-Object { $_.PowerState -like '*running*' }).Count
            Deallocated       = @($subVms | Where-Object { $_.PowerState -like '*deallocated*' }).Count
            OrphanDisks       = $subDisks.Count
            Snapshots_Waste   = @($subSnaps | Where-Object { $_.WasteFlag -eq 'WASTE' }).Count
            OrphanIPs         = $subIPs.Count
            VM_ActualCost     = [double](($subVms  | Measure-Object -Property ActualCost_Mo -Sum).Sum)
            VM_WasteCost      = [double](($subVms | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property ActualCost_Mo -Sum).Sum)
            OrphanDiskCost    = [double](($subDisks | Measure-Object -Property EstCost_Mo -Sum).Sum)
            SnapshotWasteCost = [double](($subSnaps | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum)
            OrphanIPCost      = [double](($subIPs   | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum)
        }
    }
    $summaryRows | Export-Excel -Path $outFile -WorksheetName "Summary" -AutoSize

    Write-Host "[+] Exported: $outFile" -ForegroundColor Green
    Write-Host "    Sheets: VMs, Orphan Disks, Snapshots, Orphan IPs, Summary`n" -ForegroundColor DarkGray
}
else {
    $baseFile = Join-Path $OutputPath "CloudWaste_${subLabel}_${reportDate}"

    if ($allVms.Count -gt 0) {
        $vmFile = "${baseFile}_VMs.csv"
        $allVms | Select-Object * -ExcludeProperty ResourceType | Export-Csv -Path $vmFile -NoTypeInformation -Encoding UTF8
        Write-Host "[+] VMs:          $vmFile" -ForegroundColor Green
    }
    if ($allOrphanDisks.Count -gt 0) {
        $diskFile = "${baseFile}_OrphanDisks.csv"
        $allOrphanDisks | Select-Object * -ExcludeProperty ResourceType | Export-Csv -Path $diskFile -NoTypeInformation -Encoding UTF8
        Write-Host "[+] Orphan Disks: $diskFile" -ForegroundColor Green
    }
    if ($allSnapshots.Count -gt 0) {
        $snapFile = "${baseFile}_Snapshots.csv"
        $allSnapshots | Select-Object * -ExcludeProperty ResourceType | Export-Csv -Path $snapFile -NoTypeInformation -Encoding UTF8
        Write-Host "[+] Snapshots:    $snapFile" -ForegroundColor Green
    }
    if ($allOrphanIPs.Count -gt 0) {
        $ipFile = "${baseFile}_OrphanIPs.csv"
        $allOrphanIPs | Select-Object * -ExcludeProperty ResourceType | Export-Csv -Path $ipFile -NoTypeInformation -Encoding UTF8
        Write-Host "[+] Orphan IPs:   $ipFile" -ForegroundColor Green
    }

    Write-Host ("`n    Total: {0} VMs | {1} orphan disks | {2} snapshots | {3} orphan IPs" -f `
        $allVms.Count, $allOrphanDisks.Count, $allSnapshots.Count, $allOrphanIPs.Count) -ForegroundColor DarkGray
    Write-Host ("    Monthly waste: `${0:N2}`n" -f $summary.TotalWaste) -ForegroundColor DarkGray
}

# Return all data for pipeline use
return @{
    VMs          = $allVms
    OrphanDisks  = $allOrphanDisks
    Snapshots    = $allSnapshots
    OrphanIPs    = $allOrphanIPs
    Summary      = $summary
}

