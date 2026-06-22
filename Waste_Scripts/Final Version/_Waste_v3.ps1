#Requires -Modules Az.Compute, Az.Network

<#
.SYNOPSIS
    CTS Cloud Waste & Governance Report v2.1 — GUI Edition.
    Identifies cloud waste and governance gaps across Azure Government subscriptions.

.DESCRIPTION
    Launches a Windows Forms GUI that scans a selected Azure Government subscription
    and reports on four categories of cloud waste and governance:

        - VMs:          All virtual machines with power state, governance tagging
                        status (CKID/EMASS/VASI), and estimated monthly cost.
        - Orphan Disks: Managed disks not attached to any VM, with estimated cost.
        - Snapshots:    Disk snapshots that are older than the configured age
                        threshold or whose source disk no longer exists.
        - Orphan IPs:   Public IP addresses not associated with any NIC or NAT
                        gateway, with cost flagged for Static allocations.

    Pricing is retrieved live from the Azure Retail Prices API and cached per
    VM size/region pair. Disk costs are estimated from built-in tier rate tables.
    A summary panel shows totals and a breakdown of estimated monthly waste.

    Results can be filtered in-grid and exported to CSV (one file per category).

.PARAMETER SkipPricing
    When specified, skips all calls to the Azure Retail Prices API. VM compute
    costs will be reported as $0; disk costs are still estimated from local rate
    tables. Useful for faster scans when cost accuracy is not required.

.PARAMETER MaxSnapshotAgeDays
    The age threshold (in days) used to flag snapshots as waste. Any snapshot
    older than this value — or whose source disk no longer exists — is marked
    as WASTE. Must be between 1 and 3650. Defaults to 90.

.EXAMPLE
    .\Get-CloudWasteReport-GUI.ps1
    Launches the GUI with default settings (pricing enabled, 90-day snapshot threshold).

.EXAMPLE
    .\Get-CloudWasteReport-GUI.ps1 -SkipPricing
    Launches the GUI without querying the Azure Retail Prices API.

.EXAMPLE
    .\Get-CloudWasteReport-GUI.ps1 -MaxSnapshotAgeDays 60
    Launches the GUI and flags snapshots older than 60 days as waste.

.NOTES
    Version  : 2.1
    Requires : Az.Compute, Az.Network PowerShell modules
    Target   : Azure US Government (AzureUSGovernment environment)

    Governance tag keys recognised (case-insensitive variants):
        CKID  — CKID, ckid, CkId, Ckid
        EMASS — EMASS, va_emass_id, emass, EMASNumber, EMASS_Number, emass_number
        VASI  — VASI, vasi, va_vasi_id, VASINumber, VASI_Number, vasinumber

    Actual monthly cost is calculated as:
        Running VM    = Compute cost + Disk cost
        Deallocated VM = Disk cost only (compute is not billed when deallocated)

    Static orphan IP cost is fixed at $3.65/month per Azure Government pricing.
    Disk costs use pre-loaded tier rate tables; snapshot costs use a per-GB rate.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$SkipPricing,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 3650)]
    [int]$MaxSnapshotAgeDays = 90
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptVersion = "2.1"

# ══════════════════════════════════════════════════════════════════════════════
# CORE FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

function Confirm-AzGovConnection {
    <#
    .SYNOPSIS
        Ensures the current Az session is connected to Azure US Government.

    .DESCRIPTION
        Checks for an existing Az context. If none exists, or if the current
        context is not targeting the AzureUSGovernment environment, the function
        disconnects the existing session (if any) and prompts for re-authentication
        against AzureUSGovernment.

    .OUTPUTS
        Microsoft.Azure.Commands.Profile.Models.Core.PSAzureContext
        Returns the active Azure context after ensuring correct environment.
    #>
    $context = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Connect-AzAccount -Environment AzureUSGovernment
        $context = Get-AzContext
    }
    elseif ($context.Environment.Name -ne 'AzureUSGovernment') {
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Connect-AzAccount -Environment AzureUSGovernment
        $context = Get-AzContext
    }
    return $context
}

function Get-TagValue {
    <#
    .SYNOPSIS
        Retrieves a tag value from a resource tag hashtable using multiple key variants.

    .DESCRIPTION
        Iterates over a list of possible key names (to handle inconsistent tag casing
        or naming conventions) and returns the value of the first matching key found
        in the provided tag hashtable. Returns $null if no matching key is found or
        if the hashtable is empty.

    .PARAMETER Tags
        The tag hashtable from an Azure resource object (e.g. $resource.Tags).

    .PARAMETER KeyVariants
        An array of strings representing the possible key names to search for,
        in priority order.

    .OUTPUTS
        System.String or $null
    #>
    param([hashtable]$Tags, [string[]]$KeyVariants)
    if (-not $Tags) { return $null }
    foreach ($key in $KeyVariants) { if ($Tags.ContainsKey($key)) { return $Tags[$key] } }
    return $null
}

function Get-PowerState {
    <#
    .SYNOPSIS
        Extracts the human-readable power state from a VM object.

    .DESCRIPTION
        Azure VM power state can surface in several different locations depending
        on how the VM object was retrieved (Get-AzVM -Status, instance view, etc.).
        This function checks each known location in order of preference and returns
        the first power state string found. Returns 'Unknown' if no state can be
        determined.

        Locations checked (in order):
            1. $VM.PowerState             — populated by Get-AzVM -Status
            2. $VM.InstanceView.Statuses  — instance view statuses array
            3. $VM.Statuses               — direct statuses property
            4. $VM.StatusesText           — raw JSON text fallback (regex parsed)

    .PARAMETER VM
        The VM object returned by Get-AzVM (with or without -Status).

    .OUTPUTS
        System.String
        A display string such as 'VM running', 'VM deallocated', or 'Unknown'.
    #>
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

function Invoke-WithRetry {
    <#
    .SYNOPSIS
        Executes a script block with automatic retry on failure.

    .DESCRIPTION
        Runs the provided script block up to MaxAttempts times. If the script block
        throws an exception, the function waits for an exponentially increasing delay
        (BaseDelaySeconds * attempt number) before retrying. If all attempts fail,
        the final exception is re-thrown to the caller.

        Intended for use with transient network or Azure API failures.

    .PARAMETER ScriptBlock
        The script block to execute.

    .PARAMETER MaxAttempts
        The maximum number of execution attempts. Defaults to 3.

    .PARAMETER BaseDelaySeconds
        The base number of seconds to wait between attempts. The actual delay is
        BaseDelaySeconds multiplied by the current attempt number. Defaults to 2.

    .OUTPUTS
        The return value of the script block on success.
    #>
    param(
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 3,
        [int]$BaseDelaySeconds = 2
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try { return & $ScriptBlock }
        catch {
            if ($attempt -eq $MaxAttempts) { throw }
            Start-Sleep -Seconds ($BaseDelaySeconds * $attempt)
        }
    }
}

function Get-VMPricing {
    <#
    .SYNOPSIS
        Retrieves the retail monthly price for a VM size from the Azure Pricing API.

    .DESCRIPTION
        Queries the Azure Retail Prices REST API for the given VM size and region,
        filtering to Consumption (pay-as-you-go) pricing only, excluding Spot and
        Low Priority tiers. Returns separate estimated monthly costs for Linux and
        Windows OS types. Monthly cost is calculated as hourly retail price * 730.

        The API call is wrapped in Invoke-WithRetry to handle transient failures.
        Returns $null costs if the API call fails or no matching price is found.

    .PARAMETER VmSize
        The Azure VM size string (e.g. 'Standard_D2s_v3').

    .PARAMETER Region
        The Azure region name as used in ARM (e.g. 'usgovvirginia').

    .OUTPUTS
        Hashtable with keys:
            LinuxMonthly   [decimal or $null] — estimated monthly cost for Linux
            WindowsMonthly [decimal or $null] — estimated monthly cost for Windows
    #>
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

# Disk tier monthly cost lookup tables, keyed by SKU name then by provisioned size in GB.
# Sizes follow Azure managed disk tier boundaries; costs are in USD per month.
$script:DiskTierRates = @{
    'Standard_LRS'    = @{ 4=0.48; 8=0.96; 16=1.54; 32=1.54; 64=3.01; 128=5.89; 256=11.52; 512=22.53; 1024=43.52; 2048=84.21; 4096=163.84; 8192=327.68; 16384=655.36; 32767=1310.72 }
    'StandardSSD_LRS' = @{ 4=2.40; 8=4.80; 16=9.60; 32=19.20; 64=38.40; 128=76.80; 256=153.60; 512=230.40; 1024=460.80; 2048=921.60; 4096=1843.20 }
    'Premium_LRS'     = @{ 4=5.28; 8=9.60; 16=17.92; 32=34.56; 64=66.56; 128=128.00; 256=245.76; 512=471.04; 1024=901.12; 2048=1720.32; 4096=3276.80 }
    'StandardSSD_ZRS' = @{ 4=3.00; 8=6.00; 16=12.00; 32=24.00; 64=48.00; 128=96.00; 256=192.00; 512=288.00; 1024=576.00 }
    'Premium_ZRS'     = @{ 4=6.60; 8=12.00; 16=22.40; 32=43.20; 64=83.20; 128=160.00; 256=307.20; 512=588.80; 1024=1126.40 }
}

# Ordered list of Azure managed disk tier sizes in GB, used to map a disk's actual
# provisioned size to the next available billing tier boundary.
$script:TierSizes = @(4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32767)

function Get-DiskMonthlyCost {
    <#
    .SYNOPSIS
        Estimates the monthly cost of a managed disk based on its SKU and size.

    .DESCRIPTION
        Maps the disk's provisioned size to the next Azure billing tier boundary,
        then looks up the cost from the $script:DiskTierRates table. If the SKU is
        not found in the table, attempts to fall back to the equivalent _LRS variant.
        If no table match is found, applies a flat per-GB rate based on the SKU type:
            Premium*      — $0.15/GB
            StandardSSD*  — $0.10/GB
            UltraSSD*     — $0.12/GB
            All others    — $0.05/GB

    .PARAMETER SkuName
        The managed disk SKU name (e.g. 'Premium_LRS', 'StandardSSD_ZRS').

    .PARAMETER SizeGB
        The provisioned disk size in GB. Values <= 0 are treated as 1 GB.

    .OUTPUTS
        System.Decimal — estimated monthly cost in USD.
    #>
    param([string]$SkuName, [int]$SizeGB)
    if ($SizeGB -le 0) { $SizeGB = 1 }
    $tierSize = $script:TierSizes | Where-Object { $_ -ge $SizeGB } | Select-Object -First 1
    if (-not $tierSize) { $tierSize = 32767 }
    $skuLookup = $SkuName
    if (-not $script:DiskTierRates.ContainsKey($skuLookup)) { $skuLookup = $SkuName -replace '_[A-Z]+$', '_LRS' }
    if ($script:DiskTierRates.ContainsKey($skuLookup) -and $script:DiskTierRates[$skuLookup].ContainsKey($tierSize)) {
        return $script:DiskTierRates[$skuLookup][$tierSize]
    }
    $ratePerGB = switch -Wildcard ($SkuName) { 'Premium*' { 0.15 }; 'StandardSSD*' { 0.10 }; 'UltraSSD*' { 0.12 }; default { 0.05 } }
    return [math]::Round($SizeGB * $ratePerGB, 2)
}

function Get-SnapshotMonthlyCost {
    <#
    .SYNOPSIS
        Estimates the monthly cost of a managed disk snapshot.

    .DESCRIPTION
        Calculates snapshot cost using a flat per-GB rate based on the snapshot SKU:
            Premium*   — $0.10/GB/month
            All others — $0.05/GB/month

        Unlike managed disks, snapshots are billed by actual used data size rather
        than provisioned tier, so no tier boundary lookup is applied.

    .PARAMETER SkuName
        The snapshot SKU name (e.g. 'Premium_LRS', 'Standard_LRS').

    .PARAMETER SizeGB
        The snapshot size in GB. Values <= 0 are treated as 1 GB.

    .OUTPUTS
        System.Decimal — estimated monthly cost in USD.
    #>
    param([string]$SkuName, [int]$SizeGB)
    if ($SizeGB -le 0) { $SizeGB = 1 }
    $ratePerGB = switch -Wildcard ($SkuName) { 'Premium*' { 0.10 }; default { 0.05 } }
    return [math]::Round($SizeGB * $ratePerGB, 2)
}

# Fixed monthly cost for an unattached static public IP address (Azure Government rate).
$script:StaticIPMonthlyCost = 3.65

function Get-AllSubscriptionData {
    <#
    .SYNOPSIS
        Scans an Azure subscription and returns waste and governance data for all
        VMs, orphan disks, snapshots, and orphan public IPs.

    .DESCRIPTION
        Sets the Az context to the specified subscription, then collects:
            - All managed disks (used to identify orphans and build cost indexes)
            - All snapshots (flagged by age and source disk existence)
            - All public IP addresses (unattached ones flagged as orphans)
            - All VMs with instance view (power state and cost calculation)

        For each VM, the function resolves governance tags (CKID, EMASS, VASI),
        calculates estimated compute and disk costs, determines power state, and
        assigns a WasteFlag of 'WASTE' for deallocated VMs or 'CHECK' for
        indeterminate states.

        Orphan disks are identified as managed disks not referenced by any VM's
        OS or data disk list, and not in an Attached or Reserved state.

        Snapshots are flagged as WASTE if their source disk no longer exists or
        if they exceed the MaxSnapshotAgeDays threshold; otherwise flagged REVIEW.

        Orphan IPs are public IPs with no associated IP configuration or NAT gateway.
        Static orphan IPs are flagged WASTE; dynamic ones are flagged REVIEW.

        The StatusLabel parameter is updated throughout to provide progress feedback
        in the GUI status bar.

    .PARAMETER Subscription
        The Azure subscription object (from Get-AzSubscription) to scan.

    .PARAMETER NoPricing
        When specified, skips VM pricing API calls. Compute costs will be $0.

    .PARAMETER MaxSnapshotAgeDays
        Age threshold in days for flagging snapshots as waste. Defaults to 30.

    .PARAMETER StatusLabel
        Optional. A ToolStripStatusLabel control to update with progress messages
        during the scan. If omitted, no UI updates are made.

    .OUTPUTS
        Hashtable with four keys:
            VMs         — List of PSCustomObject (VM governance and cost data)
            OrphanDisks — List of PSCustomObject (unattached managed disks)
            Snapshots   — List of PSCustomObject (all snapshots with waste flags)
            OrphanIPs   — List of PSCustomObject (unattached public IPs)
    #>
    param(
        [object]$Subscription,
        [switch]$NoPricing,
        [int]$MaxSnapshotAgeDays = 30,
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )

    Set-AzContext -SubscriptionId $Subscription.Id -ErrorAction Stop | Out-Null

    if ($StatusLabel) { $StatusLabel.Text = "Loading disks..."; [System.Windows.Forms.Application]::DoEvents() }
    $allDisks = @(Get-AzDisk -ErrorAction SilentlyContinue)
    $diskIndex = @{}; foreach ($d in $allDisks) { $diskIndex[$d.Id] = $d }

    if ($StatusLabel) { $StatusLabel.Text = "Loading snapshots..."; [System.Windows.Forms.Application]::DoEvents() }
    $allSnapshots = @(Get-AzSnapshot -ErrorAction SilentlyContinue)

    if ($StatusLabel) { $StatusLabel.Text = "Loading public IPs..."; [System.Windows.Forms.Application]::DoEvents() }
    $allPIPs = @(Get-AzPublicIpAddress -ErrorAction SilentlyContinue)

    if ($StatusLabel) { $StatusLabel.Text = "Loading VMs..."; [System.Windows.Forms.Application]::DoEvents() }
    $allVMs = @(Get-AzVM -Status -ErrorAction SilentlyContinue)

    $attachedDiskIds = @{}
    $pricingCache    = @{}
    $vmResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter   = 0
    $vmTotal   = $allVMs.Count

    foreach ($vm in $allVMs) {
        $counter++
        if ($StatusLabel) {
            $pct = if ($vmTotal -gt 0) { [math]::Round(($counter / $vmTotal) * 100) } else { 100 }
            $StatusLabel.Text = "VM $counter/$vmTotal ($pct%) — $($vm.Name)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        $ckid  = Get-TagValue -Tags $vm.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $emass = Get-TagValue -Tags $vm.Tags -KeyVariants @('EMASS','va_emass_id','emass','Emass','EMASNumber','EMASS_Number','emass_number')
        $vasi  = Get-TagValue -Tags $vm.Tags -KeyVariants @('VASI','vasi','va_vasi_id','VASINumber','VASI_Number','vasinumber')

        # Governance status based on presence of CKID (billing) and EMASS (ATO) tags.
        $ckidStatus = if ($ckid) { 'Mapped' } else { 'UNMAPPED' }
        $govStatus  = if ($ckid -and $emass)         { 'Fully Tagged' }
                      elseif ($ckid -and -not $emass) { 'Billing Only' }
                      elseif (-not $ckid -and $emass) { 'ATO Only' }
                      else                             { 'UNTAGGED' }

        $vmSize = $vm.HardwareProfile.VmSize
        $computeCost = 0
        if (-not $NoPricing) {
            # Cache pricing per size+region pair to avoid redundant API calls.
            $cacheKey = "$vmSize|$($vm.Location)"
            if (-not $pricingCache.ContainsKey($cacheKey)) { $pricingCache[$cacheKey] = Get-VMPricing -VmSize $vmSize -Region $vm.Location }
            $pricing = $pricingCache[$cacheKey]
            $isWindowsOS = $vm.StorageProfile.OsDisk.OsType -eq 'Windows'
            $price = if ($isWindowsOS) { $pricing.WindowsMonthly } else { $pricing.LinuxMonthly }
            if ($price) { $computeCost = $price }
        }

        # Accumulate disk costs and track all disk IDs as attached to skip them in the orphan scan.
        $diskCost = 0; $diskCount = 0; $totalDiskGB = 0
        $osDiskId = $vm.StorageProfile.OsDisk.ManagedDisk.Id
        if ($osDiskId) {
            $attachedDiskIds[$osDiskId] = $true
            $osDisk = $diskIndex[$osDiskId]
            if ($osDisk) { $diskCount++; $totalDiskGB += $osDisk.DiskSizeGB; $diskCost += (Get-DiskMonthlyCost -SkuName $osDisk.Sku.Name -SizeGB $osDisk.DiskSizeGB) }
        }
        foreach ($dd in $vm.StorageProfile.DataDisks) {
            if ($dd.ManagedDisk.Id) {
                $attachedDiskIds[$dd.ManagedDisk.Id] = $true
                $dataDisk = $diskIndex[$dd.ManagedDisk.Id]
                if ($dataDisk) { $diskCount++; $totalDiskGB += $dataDisk.DiskSizeGB; $diskCost += (Get-DiskMonthlyCost -SkuName $dataDisk.Sku.Name -SizeGB $dataDisk.DiskSizeGB) }
            }
        }

        $powerState    = Get-PowerState -VM $vm
        $isRunning     = $powerState -like '*running*'
        $isDeallocated = $powerState -like '*deallocated*' -or $powerState -like '*stopped*'

        # Deallocated VMs still incur disk costs; compute is only billed when running.
        $actualCost    = if ($isRunning) { $computeCost + $diskCost } else { $diskCost }
        $wasteFlag     = if ($isDeallocated) { 'WASTE' } elseif (-not $isRunning -and -not $isDeallocated) { 'CHECK' } else { '' }

        $vmResults.Add([PSCustomObject]@{
            Name             = $vm.Name
            ResourceGroup    = $vm.ResourceGroupName
            Location         = $vm.Location
            VMSize           = $vmSize
            OSType           = $vm.StorageProfile.OsDisk.OsType
            PowerState       = $powerState
            GovernanceStatus = $govStatus
            CKID_Status      = $ckidStatus
            CKID             = if ($ckid)  { $ckid }  else { '' }
            EMASS            = if ($emass) { $emass } else { '' }
            VASI             = if ($vasi)  { $vasi }  else { '' }
            Disks            = $diskCount
            DiskGB           = $totalDiskGB
            Compute_Mo       = [math]::Round($computeCost, 2)
            Disk_Mo          = [math]::Round($diskCost, 2)
            Actual_Mo        = [math]::Round($actualCost, 2)
            WasteFlag        = $wasteFlag
        })
    }

    # ── Orphan Disks ──────────────────────────────────────────────────────────
    if ($StatusLabel) { $StatusLabel.Text = "Identifying orphan disks..."; [System.Windows.Forms.Application]::DoEvents() }
    $orphanDisks = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($disk in $allDisks) {
        # Skip disks claimed by any VM or still considered attached/reserved by the API.
        if ($attachedDiskIds.ContainsKey($disk.Id)) { continue }
        if ($disk.DiskState -eq 'Attached' -or $disk.DiskState -eq 'Reserved') { continue }
        $cost = Get-DiskMonthlyCost -SkuName $disk.Sku.Name -SizeGB $disk.DiskSizeGB
        $ckid = Get-TagValue -Tags $disk.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $age  = if ($disk.TimeCreated) { [math]::Round(((Get-Date) - $disk.TimeCreated).TotalDays) } else { -1 }
        $orphanDisks.Add([PSCustomObject]@{
            Name          = $disk.Name
            ResourceGroup = $disk.ResourceGroupName
            Location      = $disk.Location
            DiskSku       = $disk.Sku.Name
            SizeGB        = $disk.DiskSizeGB
            DiskState     = $disk.DiskState
            CKID          = if ($ckid) { $ckid } else { '' }
            AgeDays       = $age
            EstCost_Mo    = $cost
            WasteFlag     = 'WASTE'
        })
    }

    # ── Snapshots ─────────────────────────────────────────────────────────────
    if ($StatusLabel) { $StatusLabel.Text = "Identifying snapshots..."; [System.Windows.Forms.Application]::DoEvents() }
    $activeDiskIds = @{}; foreach ($d in $allDisks) { $activeDiskIds[$d.Id] = $true }
    $snapResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($snap in $allSnapshots) {
        $cost      = Get-SnapshotMonthlyCost -SkuName $snap.Sku.Name -SizeGB $snap.DiskSizeGB
        $age       = if ($snap.TimeCreated) { [math]::Round(((Get-Date) - $snap.TimeCreated).TotalDays) } else { -1 }
        $srcExists = if ($snap.CreationData.SourceResourceId) { $activeDiskIds.ContainsKey($snap.CreationData.SourceResourceId) } else { $false }
        $ckid      = Get-TagValue -Tags $snap.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $isWaste   = (-not $srcExists) -or ($age -ge 0 -and $age -gt $MaxSnapshotAgeDays)
        $snapResults.Add([PSCustomObject]@{
            Name          = $snap.Name
            ResourceGroup = $snap.ResourceGroupName
            Location      = $snap.Location
            SnapshotSku   = $snap.Sku.Name
            SizeGB        = $snap.DiskSizeGB
            SourceExists  = $srcExists
            CKID          = if ($ckid) { $ckid } else { '' }
            AgeDays       = $age
            EstCost_Mo    = $cost
            WasteFlag     = if ($isWaste) { 'WASTE' } else { 'REVIEW' }
        })
    }

    # ── Orphan Public IPs ─────────────────────────────────────────────────────
    if ($StatusLabel) { $StatusLabel.Text = "Identifying orphan IPs..."; [System.Windows.Forms.Application]::DoEvents() }
    $ipResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($pip in $allPIPs) {
        # Skip IPs associated with a NIC (IpConfiguration) or NAT gateway.
        if ($pip.IpConfiguration -or $pip.NatGateway) { continue }
        $cost = if ($pip.PublicIpAllocationMethod -eq 'Static') { $script:StaticIPMonthlyCost } else { 0 }
        $ckid = Get-TagValue -Tags $pip.Tag -KeyVariants @('CKID','ckid','CkId','Ckid')
        $ipResults.Add([PSCustomObject]@{
            Name          = $pip.Name
            ResourceGroup = $pip.ResourceGroupName
            Location      = $pip.Location
            IPAddress     = if ($pip.IpAddress) { $pip.IpAddress } else { 'N/A' }
            Allocation    = $pip.PublicIpAllocationMethod
            Sku           = $pip.Sku.Name
            CKID          = if ($ckid) { $ckid } else { '' }
            EstCost_Mo    = $cost
            WasteFlag     = if ($cost -gt 0) { 'WASTE' } else { 'REVIEW' }
        })
    }

    return @{ VMs = $vmResults; OrphanDisks = $orphanDisks; Snapshots = $snapResults; OrphanIPs = $ipResults }
}

# ══════════════════════════════════════════════════════════════════════════════
# CONNECT
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "[*] Connecting to Azure Government..." -ForegroundColor Cyan
$context = Confirm-AzGovConnection
Write-Host "[+] Connected as: $($context.Account.Id)" -ForegroundColor Green

$allSubs = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
Write-Host "[+] Found $($allSubs.Count) subscriptions" -ForegroundColor Green

# ══════════════════════════════════════════════════════════════════════════════
# COLORS
# ══════════════════════════════════════════════════════════════════════════════

$bgDark   = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel  = [System.Drawing.Color]::FromArgb(45, 45, 48)
$bgAlt    = [System.Drawing.Color]::FromArgb(38, 38, 42)
$bgHeader = [System.Drawing.Color]::FromArgb(50, 50, 55)
$fgAccent = [System.Drawing.Color]::FromArgb(0, 200, 150)
$fgWhite  = [System.Drawing.Color]::White
$fgWarn   = [System.Drawing.Color]::FromArgb(255, 200, 50)
$fgRed    = [System.Drawing.Color]::FromArgb(255, 80, 80)
$fgGreen  = [System.Drawing.Color]::FromArgb(100, 255, 160)
$fgYellow = [System.Drawing.Color]::FromArgb(255, 220, 100)
$bgInput  = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgSelect = [System.Drawing.Color]::FromArgb(0, 100, 80)
$bgWaste  = [System.Drawing.Color]::FromArgb(60, 20, 20)
$bgBtn    = [System.Drawing.Color]::FromArgb(0, 150, 110)
$bgStatus = [System.Drawing.Color]::FromArgb(0, 122, 90)

# ══════════════════════════════════════════════════════════════════════════════
# FORM
# ══════════════════════════════════════════════════════════════════════════════

$form = New-Object System.Windows.Forms.Form
$form.Text          = "CTS Cloud Waste & Governance Report v$scriptVersion"
$form.Size          = New-Object System.Drawing.Size(1500, 900)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor     = $bgDark
$form.ForeColor     = $fgWhite

# ── Top Panel ─────────────────────────────────────────────────────────────────
$topPanel           = New-Object System.Windows.Forms.Panel
$topPanel.Dock      = 'Top'
$topPanel.Height    = 120
$topPanel.BackColor = $bgPanel

$titleLabel           = New-Object System.Windows.Forms.Label
$titleLabel.Text      = "CTS CLOUD WASTE & GOVERNANCE REPORT v$scriptVersion"
$titleLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $fgAccent
$titleLabel.Location  = New-Object System.Drawing.Point(12, 5)
$titleLabel.AutoSize  = $true
$topPanel.Controls.Add($titleLabel)

$ladderLabel           = New-Object System.Windows.Forms.Label
$ladderLabel.Text      = "LADDER FROZEN — CAM CMDB Migration  |  Actual Cost = Compute (if running) + Disk (always billed)"
$ladderLabel.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$ladderLabel.ForeColor = $fgWarn
$ladderLabel.Location  = New-Object System.Drawing.Point(14, 30)
$ladderLabel.AutoSize  = $true
$topPanel.Controls.Add($ladderLabel)

# Row 2 — Subscription selector and action buttons
$subLbl           = New-Object System.Windows.Forms.Label
$subLbl.Text      = "Subscription:"
$subLbl.Location  = New-Object System.Drawing.Point(12, 58)
$subLbl.AutoSize  = $true
$subLbl.ForeColor = $fgWhite
$topPanel.Controls.Add($subLbl)

$subCombo                = New-Object System.Windows.Forms.ComboBox
$subCombo.Location       = New-Object System.Drawing.Point(110, 55)
$subCombo.Size           = New-Object System.Drawing.Size(450, 25)
$subCombo.DropDownStyle  = 'DropDownList'
$subCombo.BackColor      = $bgInput
$subCombo.ForeColor      = $fgWhite
$subCombo.FlatStyle      = 'Flat'
foreach ($sub in $allSubs) { $subCombo.Items.Add("$($sub.Name)  |  $($sub.Id)") | Out-Null }
if ($subCombo.Items.Count -gt 0) { $subCombo.SelectedIndex = 0 }
$topPanel.Controls.Add($subCombo)

$scanBtn           = New-Object System.Windows.Forms.Button
$scanBtn.Text      = "SCAN"
$scanBtn.Location  = New-Object System.Drawing.Point(575, 53)
$scanBtn.Size      = New-Object System.Drawing.Size(90, 28)
$scanBtn.BackColor = $bgBtn
$scanBtn.ForeColor = $fgWhite
$scanBtn.FlatStyle = 'Flat'
$scanBtn.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$scanBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$topPanel.Controls.Add($scanBtn)

$exportBtn           = New-Object System.Windows.Forms.Button
$exportBtn.Text      = "EXPORT CSV"
$exportBtn.Location  = New-Object System.Drawing.Point(680, 53)
$exportBtn.Size      = New-Object System.Drawing.Size(110, 28)
$exportBtn.BackColor = $bgInput
$exportBtn.ForeColor = $fgWhite
$exportBtn.FlatStyle = 'Flat'
$exportBtn.Enabled   = $false
$exportBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
$topPanel.Controls.Add($exportBtn)

# Row 3 — Scan options (pricing toggle and snapshot age threshold)
$skipPricingChk           = New-Object System.Windows.Forms.CheckBox
$skipPricingChk.Text      = "Skip Pricing"
$skipPricingChk.Location  = New-Object System.Drawing.Point(12, 90)
$skipPricingChk.AutoSize  = $true
$skipPricingChk.ForeColor = $fgWhite
$skipPricingChk.Checked   = $SkipPricing.IsPresent
$topPanel.Controls.Add($skipPricingChk)

$snapAgeLbl           = New-Object System.Windows.Forms.Label
$snapAgeLbl.Text      = "Snapshot waste age (days):"
$snapAgeLbl.Location  = New-Object System.Drawing.Point(140, 92)
$snapAgeLbl.AutoSize  = $true
$snapAgeLbl.ForeColor = $fgWhite
$topPanel.Controls.Add($snapAgeLbl)

$snapAgeSpinner               = New-Object System.Windows.Forms.NumericUpDown
$snapAgeSpinner.Location      = New-Object System.Drawing.Point(315, 89)
$snapAgeSpinner.Size          = New-Object System.Drawing.Size(65, 22)
$snapAgeSpinner.Minimum       = 1
$snapAgeSpinner.Maximum       = 3650
$snapAgeSpinner.Value         = $MaxSnapshotAgeDays
$snapAgeSpinner.BackColor     = $bgInput
$snapAgeSpinner.ForeColor     = $fgWhite
$topPanel.Controls.Add($snapAgeSpinner)

# ── Summary Panel ─────────────────────────────────────────────────────────────
$summaryPanel           = New-Object System.Windows.Forms.Panel
$summaryPanel.Dock      = 'Right'
$summaryPanel.Width     = 280
$summaryPanel.BackColor = $bgAlt

$summaryTitle           = New-Object System.Windows.Forms.Label
$summaryTitle.Text      = "SUMMARY"
$summaryTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$summaryTitle.ForeColor = $fgAccent
$summaryTitle.Location  = New-Object System.Drawing.Point(15, 10)
$summaryTitle.AutoSize  = $true
$summaryPanel.Controls.Add($summaryTitle)

$summaryText           = New-Object System.Windows.Forms.Label
$summaryText.Text      = "Select a subscription`nand click SCAN"
$summaryText.Font      = New-Object System.Drawing.Font("Consolas", 9)
$summaryText.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$summaryText.Location  = New-Object System.Drawing.Point(15, 40)
$summaryText.Size      = New-Object System.Drawing.Size(250, 650)
$summaryPanel.Controls.Add($summaryText)

# ── Tab Control ───────────────────────────────────────────────────────────────
$tabControl      = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

function New-StyledGrid {
    <#
    .SYNOPSIS
        Creates and returns a pre-styled DataGridView for use in the report tabs.

    .DESCRIPTION
        Builds a DataGridView with the dark theme colours, read-only full-row
        selection, alternating row backgrounds, styled column headers, and
        auto-fill column sizing consistent with the rest of the GUI.

    .OUTPUTS
        System.Windows.Forms.DataGridView
    #>
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock                             = 'Fill'
    $g.BackgroundColor                  = $bgDark
    $g.GridColor                        = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $g.BorderStyle                      = 'None'
    $g.CellBorderStyle                  = 'SingleHorizontal'
    $g.ColumnHeadersBorderStyle         = 'Single'
    $g.EnableHeadersVisualStyles        = $false
    $g.AutoSizeColumnsMode              = 'Fill'
    $g.AllowUserToAddRows               = $false
    $g.AllowUserToDeleteRows            = $false
    $g.ReadOnly                         = $true
    $g.SelectionMode                    = 'FullRowSelect'
    $g.RowHeadersVisible                = $false
    $g.AllowUserToResizeRows            = $false
    $g.AllowUserToOrderColumns          = $true
    $g.ColumnHeadersDefaultCellStyle.BackColor        = $bgHeader
    $g.ColumnHeadersDefaultCellStyle.ForeColor        = $fgAccent
    $g.ColumnHeadersDefaultCellStyle.Font             = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.SelectionBackColor = $bgHeader
    $g.ColumnHeadersHeight              = 32
    $g.DefaultCellStyle.BackColor       = $bgDark
    $g.DefaultCellStyle.ForeColor       = $fgWhite
    $g.DefaultCellStyle.SelectionBackColor = $bgSelect
    $g.DefaultCellStyle.SelectionForeColor = $fgWhite
    $g.DefaultCellStyle.Font            = New-Object System.Drawing.Font("Consolas", 9)
    $g.AlternatingRowsDefaultCellStyle.BackColor = $bgAlt
    return $g
}

function New-TabWithSearch {
    <#
    .SYNOPSIS
        Creates a TabPage containing a filter bar and a styled DataGridView.

    .DESCRIPTION
        Builds a tab page with a top panel containing a filter text box and a
        clear button, and a DataGridView below it. Returns a hashtable with
        references to each control so the caller can wire up events and bind data.

    .PARAMETER TabText
        The display text for the tab page header.

    .OUTPUTS
        Hashtable with keys:
            Tab       — System.Windows.Forms.TabPage
            Grid      — System.Windows.Forms.DataGridView
            SearchBox — System.Windows.Forms.TextBox
            ClearBtn  — System.Windows.Forms.Button
    #>
    param([string]$TabText)
    $tab           = New-Object System.Windows.Forms.TabPage
    $tab.Text      = $TabText
    $tab.BackColor = $bgDark

    $searchPanel           = New-Object System.Windows.Forms.Panel
    $searchPanel.Dock      = 'Top'
    $searchPanel.Height    = 34
    $searchPanel.BackColor = $bgPanel

    $searchLbl           = New-Object System.Windows.Forms.Label
    $searchLbl.Text      = "Filter:"
    $searchLbl.Location  = New-Object System.Drawing.Point(8, 8)
    $searchLbl.AutoSize  = $true
    $searchLbl.ForeColor = $fgWhite
    $searchPanel.Controls.Add($searchLbl)

    $searchBox           = New-Object System.Windows.Forms.TextBox
    $searchBox.Location  = New-Object System.Drawing.Point(50, 5)
    $searchBox.Size      = New-Object System.Drawing.Size(350, 22)
    $searchBox.BackColor = $bgInput
    $searchBox.ForeColor = $fgWhite
    $searchBox.BorderStyle = 'FixedSingle'
    $searchPanel.Controls.Add($searchBox)

    $clearBtn           = New-Object System.Windows.Forms.Button
    $clearBtn.Text      = "✕"
    $clearBtn.Location  = New-Object System.Drawing.Point(408, 3)
    $clearBtn.Size      = New-Object System.Drawing.Size(26, 26)
    $clearBtn.BackColor = $bgInput
    $clearBtn.ForeColor = $fgWhite
    $clearBtn.FlatStyle = 'Flat'
    $clearBtn.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $searchPanel.Controls.Add($clearBtn)

    $grid = New-StyledGrid

    $tab.Controls.Add($grid)
    $tab.Controls.Add($searchPanel)

    return @{ Tab = $tab; Grid = $grid; SearchBox = $searchBox; ClearBtn = $clearBtn }
}

function Set-GridFilter {
    <#
    .SYNOPSIS
        Applies a case-insensitive text filter to all string columns in a grid's DataTable.

    .DESCRIPTION
        Sets the DefaultView.RowFilter on the DataTable bound to the given grid.
        The filter matches rows where any string column contains the search text
        (SQL LIKE '%text%' across all string columns, OR'd together). Passing an
        empty or whitespace-only string clears the filter.

        Non-string columns (e.g. decimal, int) are excluded from the filter to
        avoid DataTable expression type errors.

    .PARAMETER Grid
        The DataGridView whose bound DataTable should be filtered.

    .PARAMETER Text
        The search string to filter by. Empty or whitespace clears the filter.
    #>
    param($Grid, [string]$Text)
    $dt = $Grid.DataSource -as [System.Data.DataTable]
    if (-not $dt) { return }
    if ([string]::IsNullOrWhiteSpace($Text)) { $dt.DefaultView.RowFilter = ''; return }
    $safe = $Text.Trim() -replace "'", "''"
    $conditions = $dt.Columns |
        Where-Object { $_.DataType -eq [string] } |
        ForEach-Object { "$($_.ColumnName) LIKE '%$safe%'" }
    $dt.DefaultView.RowFilter = if ($conditions) { $conditions -join ' OR ' } else { '' }
}

# Create the four report tabs
$vmTab    = New-TabWithSearch -TabText "  VMs  "
$diskTab  = New-TabWithSearch -TabText "  Orphan Disks  "
$snapTab  = New-TabWithSearch -TabText "  Snapshots  "
$ipTab    = New-TabWithSearch -TabText "  Orphan IPs  "

$tabVMs   = $vmTab.Tab;   $gridVMs   = $vmTab.Grid
$tabDisks = $diskTab.Tab; $gridDisks = $diskTab.Grid
$tabSnaps = $snapTab.Tab; $gridSnaps = $snapTab.Grid
$tabIPs   = $ipTab.Tab;   $gridIPs   = $ipTab.Grid

$tabControl.TabPages.Add($tabVMs)
$tabControl.TabPages.Add($tabDisks)
$tabControl.TabPages.Add($tabSnaps)
$tabControl.TabPages.Add($tabIPs)

# Wire up search box TextChanged events and clear buttons for each tab
$vmTab.SearchBox.Add_TextChanged({   Set-GridFilter -Grid $gridVMs   -Text $vmTab.SearchBox.Text })
$diskTab.SearchBox.Add_TextChanged({ Set-GridFilter -Grid $gridDisks -Text $diskTab.SearchBox.Text })
$snapTab.SearchBox.Add_TextChanged({ Set-GridFilter -Grid $gridSnaps -Text $snapTab.SearchBox.Text })
$ipTab.SearchBox.Add_TextChanged({   Set-GridFilter -Grid $gridIPs   -Text $ipTab.SearchBox.Text })

$vmTab.ClearBtn.Add_Click({   $vmTab.SearchBox.Clear() })
$diskTab.ClearBtn.Add_Click({ $diskTab.SearchBox.Clear() })
$snapTab.ClearBtn.Add_Click({ $snapTab.SearchBox.Clear() })
$ipTab.ClearBtn.Add_Click({   $ipTab.SearchBox.Clear() })

# ── Status Bar ────────────────────────────────────────────────────────────────
$statusBar       = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = $bgStatus
$statusLabel     = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text    = "Ready — select a subscription and click SCAN"
$statusLabel.ForeColor = $fgWhite
$statusBar.Items.Add($statusLabel) | Out-Null

# Add controls in reverse dock order (Bottom/Fill dock order is sensitive to add sequence)
$form.Controls.Add($statusBar)
$form.Controls.Add($tabControl)
$form.Controls.Add($summaryPanel)
$form.Controls.Add($topPanel)

# ══════════════════════════════════════════════════════════════════════════════
# GRID LOADERS
# ══════════════════════════════════════════════════════════════════════════════

$script:scanData = $null

function Load-VMGrid {
    <#
    .SYNOPSIS
        Populates the VMs DataGridView from scan result data.

    .DESCRIPTION
        Creates a typed DataTable from the VM result objects, binds it to the
        VMs grid, formats cost columns as currency, and applies row-level colour
        coding: WASTE rows are highlighted red/dark-red; non-waste rows are
        coloured by governance status (green = Fully Tagged, yellow = Billing Only,
        red = UNTAGGED).

    .PARAMETER Data
        Array of VM PSCustomObjects returned by Get-AllSubscriptionData.
    #>
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    foreach ($col in @('Name','ResourceGroup','Location','VMSize','OSType','PowerState','GovernanceStatus','CKID_Status','CKID','EMASS','VASI','WasteFlag')) {
        $dt.Columns.Add($col, [string]) | Out-Null
    }
    $dt.Columns.Add('Disks',      [int])     | Out-Null
    $dt.Columns.Add('DiskGB',     [int])     | Out-Null
    $dt.Columns.Add('Compute_Mo', [decimal]) | Out-Null
    $dt.Columns.Add('Disk_Mo',    [decimal]) | Out-Null
    $dt.Columns.Add('Actual_Mo',  [decimal]) | Out-Null

    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name']             = $r.Name
        $row['ResourceGroup']    = $r.ResourceGroup
        $row['Location']         = $r.Location
        $row['VMSize']           = $r.VMSize
        $row['OSType']           = $r.OSType
        $row['PowerState']       = $r.PowerState
        $row['GovernanceStatus'] = $r.GovernanceStatus
        $row['CKID_Status']      = $r.CKID_Status
        $row['CKID']             = $r.CKID
        $row['EMASS']            = $r.EMASS
        $row['VASI']             = $r.VASI
        $row['WasteFlag']        = $r.WasteFlag
        $row['Disks']            = $r.Disks
        $row['DiskGB']           = $r.DiskGB
        $row['Compute_Mo']       = [decimal]$r.Compute_Mo
        $row['Disk_Mo']          = [decimal]$r.Disk_Mo
        $row['Actual_Mo']        = [decimal]$r.Actual_Mo
        $dt.Rows.Add($row)
    }
    $gridVMs.DataSource = $dt
    foreach ($col in @('Compute_Mo','Disk_Mo','Actual_Mo')) {
        if ($gridVMs.Columns.Contains($col)) {
            $gridVMs.Columns[$col].DefaultCellStyle.Format    = 'C2'
            $gridVMs.Columns[$col].DefaultCellStyle.Alignment = 'MiddleRight'
        }
    }
    foreach ($gridRow in $gridVMs.Rows) {
        $w = $gridRow.Cells['WasteFlag'].Value
        if ($w -eq 'WASTE') {
            $gridRow.DefaultCellStyle.ForeColor = $fgRed
            $gridRow.DefaultCellStyle.BackColor = $bgWaste
        } else {
            switch ($gridRow.Cells['GovernanceStatus'].Value) {
                'UNTAGGED'     { $gridRow.DefaultCellStyle.ForeColor = $fgRed }
                'Billing Only' { $gridRow.DefaultCellStyle.ForeColor = $fgYellow }
                'Fully Tagged' { $gridRow.DefaultCellStyle.ForeColor = $fgGreen }
            }
        }
    }
}

function Load-OrphanDiskGrid {
    <#
    .SYNOPSIS
        Populates the Orphan Disks DataGridView from scan result data.

    .DESCRIPTION
        Creates a typed DataTable from orphan disk result objects, binds it to
        the Orphan Disks grid, formats the cost column as currency, and colours
        all rows red to indicate every result is flagged as waste.

    .PARAMETER Data
        Array of orphan disk PSCustomObjects returned by Get-AllSubscriptionData.
    #>
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    foreach ($col in @('Name','ResourceGroup','Location','DiskSku','DiskState','CKID','WasteFlag')) {
        $dt.Columns.Add($col, [string]) | Out-Null
    }
    $dt.Columns.Add('SizeGB',     [int])     | Out-Null
    $dt.Columns.Add('AgeDays',    [int])     | Out-Null
    $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null

    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name']          = $r.Name
        $row['ResourceGroup'] = $r.ResourceGroup
        $row['Location']      = $r.Location
        $row['DiskSku']       = $r.DiskSku
        $row['SizeGB']        = $r.SizeGB
        $row['DiskState']     = $r.DiskState
        $row['CKID']          = $r.CKID
        $row['AgeDays']       = $r.AgeDays
        $row['EstCost_Mo']    = [decimal]$r.EstCost_Mo
        $row['WasteFlag']     = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridDisks.DataSource = $dt
    if ($gridDisks.Columns.Contains('EstCost_Mo')) {
        $gridDisks.Columns['EstCost_Mo'].DefaultCellStyle.Format    = 'C2'
        $gridDisks.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight'
    }
    foreach ($gridRow in $gridDisks.Rows) { $gridRow.DefaultCellStyle.ForeColor = $fgRed }
}

function Load-SnapshotGrid {
    <#
    .SYNOPSIS
        Populates the Snapshots DataGridView from scan result data.

    .DESCRIPTION
        Creates a typed DataTable from snapshot result objects, binds it to the
        Snapshots grid, formats the cost column as currency, and colours rows red
        for WASTE and yellow for REVIEW.

    .PARAMETER Data
        Array of snapshot PSCustomObjects returned by Get-AllSubscriptionData.
    #>
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    foreach ($col in @('Name','ResourceGroup','Location','SnapshotSku','SourceExists','CKID','WasteFlag')) {
        $dt.Columns.Add($col, [string]) | Out-Null
    }
    $dt.Columns.Add('SizeGB',     [int])     | Out-Null
    $dt.Columns.Add('AgeDays',    [int])     | Out-Null
    $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null

    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name']          = $r.Name
        $row['ResourceGroup'] = $r.ResourceGroup
        $row['Location']      = $r.Location
        $row['SnapshotSku']   = $r.SnapshotSku
        $row['SizeGB']        = $r.SizeGB
        $row['SourceExists']  = $r.SourceExists
        $row['CKID']          = $r.CKID
        $row['AgeDays']       = $r.AgeDays
        $row['EstCost_Mo']    = [decimal]$r.EstCost_Mo
        $row['WasteFlag']     = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridSnaps.DataSource = $dt
    if ($gridSnaps.Columns.Contains('EstCost_Mo')) {
        $gridSnaps.Columns['EstCost_Mo'].DefaultCellStyle.Format    = 'C2'
        $gridSnaps.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight'
    }
    foreach ($gridRow in $gridSnaps.Rows) {
        $gridRow.DefaultCellStyle.ForeColor = if ($gridRow.Cells['WasteFlag'].Value -eq 'WASTE') { $fgRed } else { $fgYellow }
    }
}

function Load-IPGrid {
    <#
    .SYNOPSIS
        Populates the Orphan IPs DataGridView from scan result data.

    .DESCRIPTION
        Creates a typed DataTable from orphan IP result objects, binds it to the
        Orphan IPs grid, formats the cost column as currency, and colours all
        rows red to indicate every result requires attention.

    .PARAMETER Data
        Array of orphan IP PSCustomObjects returned by Get-AllSubscriptionData.
    #>
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    foreach ($col in @('Name','ResourceGroup','Location','IPAddress','Allocation','Sku','CKID','WasteFlag')) {
        $dt.Columns.Add($col, [string]) | Out-Null
    }
    $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null

    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name']          = $r.Name
        $row['ResourceGroup'] = $r.ResourceGroup
        $row['Location']      = $r.Location
        $row['IPAddress']     = $r.IPAddress
        $row['Allocation']    = $r.Allocation
        $row['Sku']           = $r.Sku
        $row['CKID']          = $r.CKID
        $row['EstCost_Mo']    = [decimal]$r.EstCost_Mo
        $row['WasteFlag']     = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridIPs.DataSource = $dt
    if ($gridIPs.Columns.Contains('EstCost_Mo')) {
        $gridIPs.Columns['EstCost_Mo'].DefaultCellStyle.Format    = 'C2'
        $gridIPs.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight'
    }
    foreach ($gridRow in $gridIPs.Rows) { $gridRow.DefaultCellStyle.ForeColor = $fgRed }
}

function Update-Summary {
    <#
    .SYNOPSIS
        Refreshes the summary panel and tab header counts after a scan completes.

    .DESCRIPTION
        Reads from $script:scanData to compute aggregate totals for VMs, orphan
        disks, snapshots, and orphan IPs. Updates the summary label with VM counts
        by power state and CKID mapping status, estimated monthly costs broken down
        by compute, disk, and waste category, and a total estimated monthly waste
        figure. Also updates each tab's header text to show the item count.
    #>
    $vms = $script:scanData.VMs
    $od  = $script:scanData.OrphanDisks
    $sn  = $script:scanData.Snapshots
    $ips = $script:scanData.OrphanIPs

    $total   = $vms.Count
    $mapped  = @($vms | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
    $pct     = if ($total -gt 0) { [math]::Round(($mapped / $total) * 100, 1) } else { 0 }
    $running = @($vms | Where-Object { $_.PowerState -like '*running*' }).Count
    $dealloc = @($vms | Where-Object { $_.PowerState -like '*deallocated*' }).Count

    $vmActual  = [double](($vms | Measure-Object -Property Actual_Mo  -Sum).Sum)
    $vmCompute = [double](($vms | Measure-Object -Property Compute_Mo -Sum).Sum)
    $vmDisk    = [double](($vms | Measure-Object -Property Disk_Mo    -Sum).Sum)
    $vmWaste   = [double](($vms | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property Actual_Mo -Sum).Sum)
    $vmWasteN  = @($vms | Where-Object { $_.WasteFlag -eq 'WASTE' }).Count

    $odCost     = [double](($od  | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $snCost     = [double](($sn  | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $ipCost     = [double](($ips | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum)
    $totalWaste = $vmWaste + $odCost + $snCost + $ipCost

    # Update tab header text to show current result counts
    $tabVMs.Text   = "  VMs ($total)  "
    $tabDisks.Text = "  Orphan Disks ($($od.Count))  "
    $tabSnaps.Text = "  Snapshots ($($sn.Count))  "
    $tabIPs.Text   = "  Orphan IPs ($($ips.Count))  "

    $summaryText.Text = @"
VMs: $total
─────────────────────
CKID Mapped:  $mapped ($pct%)
Unmapped:     $($total - $mapped)
Running:      $running
Deallocated:  $dealloc

VM COSTS
─────────────────────
Compute/Mo: $("{0:C2}" -f $vmCompute)
Disk/Mo:    $("{0:C2}" -f $vmDisk)
Actual/Mo:  $("{0:C2}" -f $vmActual)
VM Waste:   $("{0:C2}" -f $vmWaste)
            ($vmWasteN VMs)

ORPHANS
─────────────────────
Disks: $($od.Count)  $("{0:C2}" -f $odCost)/mo
Snaps: $($sn.Count)  $("{0:C2}" -f $snCost)/mo
IPs:   $($ips.Count)  $("{0:C2}" -f $ipCost)/mo

═════════════════════
TOTAL WASTE:
$("{0:C2}" -f $totalWaste)/mo
═════════════════════
"@
}

# ══════════════════════════════════════════════════════════════════════════════
# EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$scanBtn.Add_Click({
    # Validate selection, disable controls, and start the scan for the chosen subscription.
    $idx = $subCombo.SelectedIndex
    if ($idx -lt 0) { return }
    $sub = $allSubs[$idx]

    $scanBtn.Enabled = $false
    $scanBtn.Text    = "Scanning..."
    $statusLabel.Text = "Scanning $($sub.Name)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:scanData = Get-AllSubscriptionData `
            -Subscription       $sub `
            -NoPricing:         $skipPricingChk.Checked `
            -MaxSnapshotAgeDays ([int]$snapAgeSpinner.Value) `
            -StatusLabel        $statusLabel

        # Clear any active filters before reloading grids
        foreach ($sb in @($vmTab.SearchBox, $diskTab.SearchBox, $snapTab.SearchBox, $ipTab.SearchBox)) { $sb.Clear() }

        Load-VMGrid         -Data $script:scanData.VMs
        Load-OrphanDiskGrid -Data $script:scanData.OrphanDisks
        Load-SnapshotGrid   -Data $script:scanData.Snapshots
        Load-IPGrid         -Data $script:scanData.OrphanIPs
        Update-Summary

        $exportBtn.Enabled = $true
        $statusLabel.Text  = "Scan complete — $($script:scanData.VMs.Count) VMs, $($script:scanData.OrphanDisks.Count) orphan disks, $($script:scanData.Snapshots.Count) snapshots, $($script:scanData.OrphanIPs.Count) orphan IPs"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error:`n$($_.Exception.Message)", "Scan Error", 'OK', 'Error')
        $statusLabel.Text = "Error during scan"
    }
    finally {
        $scanBtn.Enabled = $true
        $scanBtn.Text    = "SCAN"
    }
})

$exportBtn.Add_Click({
    # Prompt the user for a save location and export one CSV per data category.
    # The filename chosen for VMs is used as the base path; category suffixes are appended.
    if (-not $script:scanData) { return }
    $saveDialog            = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter     = "CSV Files (*.csv)|*.csv"
    $saveDialog.DefaultExt = "csv"
    $subName               = $allSubs[$subCombo.SelectedIndex].Name -replace '[^a-zA-Z0-9\-]', '_'
    $saveDialog.FileName   = "CloudWaste_${subName}_$(Get-Date -Format 'yyyy-MM-dd')_VMs.csv"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        $basePath = $saveDialog.FileName -replace '_VMs\.csv$', ''
        if ($script:scanData.VMs.Count -gt 0)         { $script:scanData.VMs         | Export-Csv -Path "${basePath}_VMs.csv"         -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.OrphanDisks.Count -gt 0) { $script:scanData.OrphanDisks | Export-Csv -Path "${basePath}_OrphanDisks.csv" -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.Snapshots.Count -gt 0)   { $script:scanData.Snapshots   | Export-Csv -Path "${basePath}_Snapshots.csv"   -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.OrphanIPs.Count -gt 0)   { $script:scanData.OrphanIPs   | Export-Csv -Path "${basePath}_OrphanIPs.csv"   -NoTypeInformation -Encoding UTF8 }
        $statusLabel.Text = "Exported to: $basePath*.csv"
        [System.Windows.Forms.MessageBox]::Show("Exported to:`n$basePath*.csv", "Export Complete", 'OK', 'Information')
    }
})

# ══════════════════════════════════════════════════════════════════════════════
# LAUNCH
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "[+] Launching GUI v$scriptVersion..." -ForegroundColor Green
[void]$form.ShowDialog()
$form.Dispose()