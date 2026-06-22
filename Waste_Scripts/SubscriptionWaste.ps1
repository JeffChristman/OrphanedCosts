#Requires -Modules Az.Compute, Az.Network

<#
.SYNOPSIS
    CTS Cloud Waste & Governance Report v2.0 — GUI Edition
    Tabbed interface: VMs | Orphan Disks | Snapshots | Orphan IPs

.EXAMPLE
    .\Get-CloudWasteReport-GUI.ps1
    .\Get-CloudWasteReport-GUI.ps1 -SkipPricing
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$SkipPricing
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$scriptVersion = "2.0"

# ══════════════════════════════════════════════════════════════════════════════
# CORE FUNCTIONS (same as console v2)
# ══════════════════════════════════════════════════════════════════════════════

function Confirm-AzGovConnection {
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
    catch { return @{ LinuxMonthly = $null; WindowsMonthly = $null } }
}

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
    param([string]$SkuName, [int]$SizeGB)
    $ratePerGB = switch -Wildcard ($SkuName) { 'Premium*' { 0.10 }; default { 0.05 } }
    return [math]::Round($SizeGB * $ratePerGB, 2)
}

$script:StaticIPMonthlyCost = 3.65

function Get-AllSubscriptionData {
    param(
        [object]$Subscription,
        [switch]$NoPricing,
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )

    Set-AzContext -SubscriptionId $Subscription.Id -ErrorAction Stop | Out-Null
    $subName = $Subscription.Name

    # Phase 1-3: Bulk loads
    if ($StatusLabel) { $StatusLabel.Text = "Loading disks..."; [System.Windows.Forms.Application]::DoEvents() }
    $allDisks = @(Get-AzDisk -ErrorAction SilentlyContinue)
    $diskIndex = @{}; foreach ($d in $allDisks) { $diskIndex[$d.Id] = $d }

    if ($StatusLabel) { $StatusLabel.Text = "Loading snapshots..."; [System.Windows.Forms.Application]::DoEvents() }
    $allSnapshots = @(Get-AzSnapshot -ErrorAction SilentlyContinue)

    if ($StatusLabel) { $StatusLabel.Text = "Loading public IPs..."; [System.Windows.Forms.Application]::DoEvents() }
    $allPIPs = @(Get-AzPublicIpAddress -ErrorAction SilentlyContinue)

    if ($StatusLabel) { $StatusLabel.Text = "Loading VMs..."; [System.Windows.Forms.Application]::DoEvents() }
    $allVMs = @(Get-AzVM -Status -ErrorAction SilentlyContinue)

    # Process VMs
    $attachedDiskIds = @{}
    $pricingCache = @{}
    $vmResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter = 0

    foreach ($vm in $allVMs) {
        $counter++
        if ($StatusLabel) {
            $pct = [math]::Round(($counter / $allVMs.Count) * 100)
            $StatusLabel.Text = "VM $counter/$($allVMs.Count) ($pct%) — $($vm.Name)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        $ckid  = Get-TagValue -Tags $vm.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $emass = Get-TagValue -Tags $vm.Tags -KeyVariants @('EMASS','eMASS','emass','Emass','EMASNumber','EMASS_Number','emass_number')
        $vasi  = Get-TagValue -Tags $vm.Tags -KeyVariants @('VASI','vasi','Vasi','VASINumber','VASI_Number','vasinumber')

        $ckidStatus = if ($ckid) { 'Mapped' } else { 'UNMAPPED' }
        $govStatus = if ($ckid -and $emass) { 'Fully Tagged' }
                     elseif ($ckid -and -not $emass) { 'Billing Only' }
                     elseif (-not $ckid -and $emass) { 'ATO Only' }
                     else { 'UNTAGGED' }

        $vmSize = $vm.HardwareProfile.VmSize
        $computeCost = 0
        if (-not $NoPricing) {
            $cacheKey = "$vmSize|$($vm.Location)"
            if (-not $pricingCache.ContainsKey($cacheKey)) { $pricingCache[$cacheKey] = Get-VMPricing -VmSize $vmSize -Region $vm.Location }
            $pricing = $pricingCache[$cacheKey]
            $isWindowsOS = $vm.StorageProfile.OsDisk.OsType -eq 'Windows'
            $price = if ($isWindowsOS) { $pricing.WindowsMonthly } else { $pricing.LinuxMonthly }
            if ($price) { $computeCost = $price }
        }

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

        $powerState = Get-PowerState -VM $vm
        $isRunning = $powerState -like '*running*'
        $isDeallocated = $powerState -like '*deallocated*' -or $powerState -like '*stopped*'
        $actualCost = if ($isRunning) { $computeCost + $diskCost } else { $diskCost }
        $wasteFlag = if ($isDeallocated) { 'WASTE' } elseif (-not $isRunning -and -not $isDeallocated) { 'CHECK' } else { '' }

        $vmResults.Add([PSCustomObject]@{
            Name = $vm.Name; ResourceGroup = $vm.ResourceGroupName; Location = $vm.Location
            VMSize = $vmSize; OSType = $vm.StorageProfile.OsDisk.OsType; PowerState = $powerState
            GovernanceStatus = $govStatus; CKID_Status = $ckidStatus
            CKID = if ($ckid) { $ckid } else { '' }; EMASS = if ($emass) { $emass } else { '' }; VASI = if ($vasi) { $vasi } else { '' }
            Disks = $diskCount; DiskGB = $totalDiskGB
            Compute_Mo = [math]::Round($computeCost, 2); Disk_Mo = [math]::Round($diskCost, 2)
            Actual_Mo = [math]::Round($actualCost, 2); WasteFlag = $wasteFlag
        })
    }

    # Orphan Disks
    if ($StatusLabel) { $StatusLabel.Text = "Identifying orphan disks..."; [System.Windows.Forms.Application]::DoEvents() }
    $orphanDisks = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($disk in $allDisks) {
        if ($attachedDiskIds.ContainsKey($disk.Id) -or $disk.DiskState -eq 'Attached') { continue }
        $cost = Get-DiskMonthlyCost -SkuName $disk.Sku.Name -SizeGB $disk.DiskSizeGB
        $ckid = Get-TagValue -Tags $disk.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $age = if ($disk.TimeCreated) { [math]::Round(((Get-Date) - $disk.TimeCreated).TotalDays) } else { -1 }
        $orphanDisks.Add([PSCustomObject]@{
            Name = $disk.Name; ResourceGroup = $disk.ResourceGroupName; Location = $disk.Location
            DiskSku = $disk.Sku.Name; SizeGB = $disk.DiskSizeGB; DiskState = $disk.DiskState
            CKID = if ($ckid) { $ckid } else { '' }; AgeDays = $age
            EstCost_Mo = $cost; WasteFlag = 'WASTE'
        })
    }

    # Snapshots
    if ($StatusLabel) { $StatusLabel.Text = "Identifying snapshots..."; [System.Windows.Forms.Application]::DoEvents() }
    $activeDiskIds = @{}; foreach ($d in $allDisks) { $activeDiskIds[$d.Id] = $true }
    $snapResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($snap in $allSnapshots) {
        $cost = Get-SnapshotMonthlyCost -SkuName $snap.Sku.Name -SizeGB $snap.DiskSizeGB
        $age = if ($snap.TimeCreated) { [math]::Round(((Get-Date) - $snap.TimeCreated).TotalDays) } else { -1 }
        $srcExists = if ($snap.CreationData.SourceResourceId) { $activeDiskIds.ContainsKey($snap.CreationData.SourceResourceId) } else { $false }
        $ckid = Get-TagValue -Tags $snap.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $wf = if (-not $srcExists -or $age -gt 90) { 'WASTE' } else { 'REVIEW' }
        $snapResults.Add([PSCustomObject]@{
            Name = $snap.Name; ResourceGroup = $snap.ResourceGroupName; Location = $snap.Location
            SnapshotSku = $snap.Sku.Name; SizeGB = $snap.DiskSizeGB; SourceExists = $srcExists
            CKID = if ($ckid) { $ckid } else { '' }; AgeDays = $age
            EstCost_Mo = $cost; WasteFlag = $wf
        })
    }

    # Orphan IPs
    if ($StatusLabel) { $StatusLabel.Text = "Identifying orphan IPs..."; [System.Windows.Forms.Application]::DoEvents() }
    $ipResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($pip in $allPIPs) {
        if ($pip.IpConfiguration -or $pip.NatGateway) { continue }
        $cost = if ($pip.PublicIpAllocationMethod -eq 'Static') { $script:StaticIPMonthlyCost } else { 0 }
        $ckid = Get-TagValue -Tags $pip.Tag -KeyVariants @('CKID','ckid','CkId','Ckid')
        $ipResults.Add([PSCustomObject]@{
            Name = $pip.Name; ResourceGroup = $pip.ResourceGroupName; Location = $pip.Location
            IPAddress = if ($pip.IpAddress) { $pip.IpAddress } else { 'N/A' }
            Allocation = $pip.PublicIpAllocationMethod; Sku = $pip.Sku.Name
            CKID = if ($ckid) { $ckid } else { '' }
            EstCost_Mo = $cost; WasteFlag = if ($cost -gt 0) { 'WASTE' } else { 'REVIEW' }
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
# BUILD GUI
# ══════════════════════════════════════════════════════════════════════════════

# Colors
$bgDark    = [System.Drawing.Color]::FromArgb(30, 30, 30)
$bgPanel   = [System.Drawing.Color]::FromArgb(45, 45, 48)
$bgAlt     = [System.Drawing.Color]::FromArgb(38, 38, 42)
$bgHeader  = [System.Drawing.Color]::FromArgb(50, 50, 55)
$fgAccent  = [System.Drawing.Color]::FromArgb(0, 200, 150)
$fgWhite   = [System.Drawing.Color]::White
$fgWarn    = [System.Drawing.Color]::FromArgb(255, 200, 50)
$fgRed     = [System.Drawing.Color]::FromArgb(255, 80, 80)
$fgGreen   = [System.Drawing.Color]::FromArgb(100, 255, 160)
$fgYellow  = [System.Drawing.Color]::FromArgb(255, 220, 100)
$bgInput   = [System.Drawing.Color]::FromArgb(60, 60, 60)
$bgSelect  = [System.Drawing.Color]::FromArgb(0, 100, 80)
$bgWaste   = [System.Drawing.Color]::FromArgb(60, 20, 20)
$bgBtn     = [System.Drawing.Color]::FromArgb(0, 150, 110)
$bgStatus  = [System.Drawing.Color]::FromArgb(0, 122, 90)

# Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "CTS Cloud Waste & Governance Report v$scriptVersion"
$form.Size = New-Object System.Drawing.Size(1500, 900)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $bgDark
$form.ForeColor = $fgWhite

# ── Top Panel ─────────────────────────────────────────────────────────────────
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'; $topPanel.Height = 95; $topPanel.BackColor = $bgPanel

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "CTS CLOUD WASTE & GOVERNANCE REPORT v$scriptVersion"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $fgAccent; $titleLabel.Location = New-Object System.Drawing.Point(12, 5); $titleLabel.AutoSize = $true
$topPanel.Controls.Add($titleLabel)

$ladderLabel = New-Object System.Windows.Forms.Label
$ladderLabel.Text = "LADDER FROZEN — CAM CMDB Migration  |  Actual Cost = Compute (if running) + Disk (always billed)"
$ladderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$ladderLabel.ForeColor = $fgWarn; $ladderLabel.Location = New-Object System.Drawing.Point(14, 30); $ladderLabel.AutoSize = $true
$topPanel.Controls.Add($ladderLabel)

# Subscription
$subLbl = New-Object System.Windows.Forms.Label
$subLbl.Text = "Subscription:"; $subLbl.Location = New-Object System.Drawing.Point(12, 58); $subLbl.AutoSize = $true; $subLbl.ForeColor = $fgWhite
$topPanel.Controls.Add($subLbl)

$subCombo = New-Object System.Windows.Forms.ComboBox
$subCombo.Location = New-Object System.Drawing.Point(110, 55); $subCombo.Size = New-Object System.Drawing.Size(450, 25)
$subCombo.DropDownStyle = 'DropDownList'; $subCombo.BackColor = $bgInput; $subCombo.ForeColor = $fgWhite; $subCombo.FlatStyle = 'Flat'
foreach ($sub in $allSubs) { $subCombo.Items.Add("$($sub.Name)  |  $($sub.Id)") | Out-Null }
if ($subCombo.Items.Count -gt 0) { $subCombo.SelectedIndex = 0 }
$topPanel.Controls.Add($subCombo)

$scanBtn = New-Object System.Windows.Forms.Button
$scanBtn.Text = "SCAN"; $scanBtn.Location = New-Object System.Drawing.Point(575, 53); $scanBtn.Size = New-Object System.Drawing.Size(90, 28)
$scanBtn.BackColor = $bgBtn; $scanBtn.ForeColor = $fgWhite; $scanBtn.FlatStyle = 'Flat'
$scanBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $scanBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$topPanel.Controls.Add($scanBtn)

$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "EXPORT CSV"; $exportBtn.Location = New-Object System.Drawing.Point(680, 53); $exportBtn.Size = New-Object System.Drawing.Size(110, 28)
$exportBtn.BackColor = $bgInput; $exportBtn.ForeColor = $fgWhite; $exportBtn.FlatStyle = 'Flat'; $exportBtn.Enabled = $false
$exportBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$topPanel.Controls.Add($exportBtn)

# ── Summary Panel ─────────────────────────────────────────────────────────────
$summaryPanel = New-Object System.Windows.Forms.Panel
$summaryPanel.Dock = 'Right'; $summaryPanel.Width = 280; $summaryPanel.BackColor = $bgAlt

$summaryTitle = New-Object System.Windows.Forms.Label
$summaryTitle.Text = "SUMMARY"; $summaryTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$summaryTitle.ForeColor = $fgAccent; $summaryTitle.Location = New-Object System.Drawing.Point(15, 10); $summaryTitle.AutoSize = $true
$summaryPanel.Controls.Add($summaryTitle)

$summaryText = New-Object System.Windows.Forms.Label
$summaryText.Text = "Select a subscription`nand click SCAN"
$summaryText.Font = New-Object System.Drawing.Font("Consolas", 9)
$summaryText.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$summaryText.Location = New-Object System.Drawing.Point(15, 40); $summaryText.Size = New-Object System.Drawing.Size(250, 650)
$summaryPanel.Controls.Add($summaryText)

# ── Tab Control ───────────────────────────────────────────────────────────────
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

# Helper to create a styled DataGridView
function New-StyledGrid {
    $g = New-Object System.Windows.Forms.DataGridView
    $g.Dock = 'Fill'; $g.BackgroundColor = $bgDark; $g.GridColor = [System.Drawing.Color]::FromArgb(60,60,60)
    $g.BorderStyle = 'None'; $g.CellBorderStyle = 'SingleHorizontal'; $g.ColumnHeadersBorderStyle = 'Single'
    $g.EnableHeadersVisualStyles = $false; $g.AutoSizeColumnsMode = 'Fill'
    $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false; $g.ReadOnly = $true
    $g.SelectionMode = 'FullRowSelect'; $g.RowHeadersVisible = $false; $g.AllowUserToResizeRows = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = $bgHeader; $g.ColumnHeadersDefaultCellStyle.ForeColor = $fgAccent
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.SelectionBackColor = $bgHeader; $g.ColumnHeadersHeight = 32
    $g.DefaultCellStyle.BackColor = $bgDark; $g.DefaultCellStyle.ForeColor = $fgWhite
    $g.DefaultCellStyle.SelectionBackColor = $bgSelect; $g.DefaultCellStyle.SelectionForeColor = $fgWhite
    $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
    $g.AlternatingRowsDefaultCellStyle.BackColor = $bgAlt
    return $g
}

# VM Tab
$tabVMs = New-Object System.Windows.Forms.TabPage; $tabVMs.Text = "  VMs  "; $tabVMs.BackColor = $bgDark
$gridVMs = New-StyledGrid; $tabVMs.Controls.Add($gridVMs)
$tabControl.TabPages.Add($tabVMs)

# Orphan Disks Tab
$tabDisks = New-Object System.Windows.Forms.TabPage; $tabDisks.Text = "  Orphan Disks  "; $tabDisks.BackColor = $bgDark
$gridDisks = New-StyledGrid; $tabDisks.Controls.Add($gridDisks)
$tabControl.TabPages.Add($tabDisks)

# Snapshots Tab
$tabSnaps = New-Object System.Windows.Forms.TabPage; $tabSnaps.Text = "  Snapshots  "; $tabSnaps.BackColor = $bgDark
$gridSnaps = New-StyledGrid; $tabSnaps.Controls.Add($gridSnaps)
$tabControl.TabPages.Add($tabSnaps)

# Orphan IPs Tab
$tabIPs = New-Object System.Windows.Forms.TabPage; $tabIPs.Text = "  Orphan IPs  "; $tabIPs.BackColor = $bgDark
$gridIPs = New-StyledGrid; $tabIPs.Controls.Add($gridIPs)
$tabControl.TabPages.Add($tabIPs)

# ── Status Bar ────────────────────────────────────────────────────────────────
$statusBar = New-Object System.Windows.Forms.StatusStrip; $statusBar.BackColor = $bgStatus
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready — select a subscription and click SCAN"; $statusLabel.ForeColor = $fgWhite
$statusBar.Items.Add($statusLabel) | Out-Null

# ── Add controls (reverse dock order) ─────────────────────────────────────────
$form.Controls.Add($statusBar)
$form.Controls.Add($tabControl)
$form.Controls.Add($summaryPanel)
$form.Controls.Add($topPanel)

# ══════════════════════════════════════════════════════════════════════════════
# DATA + GRID HELPERS
# ══════════════════════════════════════════════════════════════════════════════

$script:scanData = $null

function Load-VMGrid {
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    @('Name','ResourceGroup','VMSize','OSType','PowerState','GovernanceStatus','CKID_Status','CKID','EMASS','WasteFlag') |
        ForEach-Object { $dt.Columns.Add($_, [string]) | Out-Null }
    $dt.Columns.Add('Disks', [int]) | Out-Null
    $dt.Columns.Add('DiskGB', [int]) | Out-Null
    $dt.Columns.Add('Compute_Mo', [decimal]) | Out-Null
    $dt.Columns.Add('Disk_Mo', [decimal]) | Out-Null
    $dt.Columns.Add('Actual_Mo', [decimal]) | Out-Null

    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name'] = $r.Name; $row['ResourceGroup'] = $r.ResourceGroup; $row['VMSize'] = $r.VMSize
        $row['OSType'] = $r.OSType; $row['PowerState'] = $r.PowerState
        $row['GovernanceStatus'] = $r.GovernanceStatus; $row['CKID_Status'] = $r.CKID_Status
        $row['CKID'] = $r.CKID; $row['EMASS'] = $r.EMASS; $row['WasteFlag'] = $r.WasteFlag
        $row['Disks'] = $r.Disks; $row['DiskGB'] = $r.DiskGB
        $row['Compute_Mo'] = [decimal]$r.Compute_Mo; $row['Disk_Mo'] = [decimal]$r.Disk_Mo; $row['Actual_Mo'] = [decimal]$r.Actual_Mo
        $dt.Rows.Add($row)
    }
    $gridVMs.DataSource = $dt
    foreach ($col in @('Compute_Mo','Disk_Mo','Actual_Mo')) {
        if ($gridVMs.Columns.Contains($col)) { $gridVMs.Columns[$col].DefaultCellStyle.Format = 'C2'; $gridVMs.Columns[$col].DefaultCellStyle.Alignment = 'MiddleRight' }
    }
    foreach ($gridRow in $gridVMs.Rows) {
        $w = $gridRow.Cells['WasteFlag'].Value
        if ($w -eq 'WASTE') { $gridRow.DefaultCellStyle.ForeColor = $fgRed; $gridRow.DefaultCellStyle.BackColor = $bgWaste }
        else {
            $g = $gridRow.Cells['GovernanceStatus'].Value
            switch ($g) { 'UNTAGGED' { $gridRow.DefaultCellStyle.ForeColor = $fgRed }; 'Billing Only' { $gridRow.DefaultCellStyle.ForeColor = $fgYellow }; 'Fully Tagged' { $gridRow.DefaultCellStyle.ForeColor = $fgGreen } }
        }
    }
}

function Load-OrphanDiskGrid {
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    @('Name','ResourceGroup','Location','DiskSku','DiskState','CKID','WasteFlag') | ForEach-Object { $dt.Columns.Add($_, [string]) | Out-Null }
    $dt.Columns.Add('SizeGB', [int]) | Out-Null; $dt.Columns.Add('AgeDays', [int]) | Out-Null; $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null
    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name'] = $r.Name; $row['ResourceGroup'] = $r.ResourceGroup; $row['Location'] = $r.Location
        $row['DiskSku'] = $r.DiskSku; $row['SizeGB'] = $r.SizeGB; $row['DiskState'] = $r.DiskState
        $row['CKID'] = $r.CKID; $row['AgeDays'] = $r.AgeDays; $row['EstCost_Mo'] = [decimal]$r.EstCost_Mo; $row['WasteFlag'] = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridDisks.DataSource = $dt
    if ($gridDisks.Columns.Contains('EstCost_Mo')) { $gridDisks.Columns['EstCost_Mo'].DefaultCellStyle.Format = 'C2'; $gridDisks.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight' }
    foreach ($gridRow in $gridDisks.Rows) { $gridRow.DefaultCellStyle.ForeColor = $fgRed }
}

function Load-SnapshotGrid {
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    @('Name','ResourceGroup','Location','SnapshotSku','SourceExists','CKID','WasteFlag') | ForEach-Object { $dt.Columns.Add($_, [string]) | Out-Null }
    $dt.Columns.Add('SizeGB', [int]) | Out-Null; $dt.Columns.Add('AgeDays', [int]) | Out-Null; $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null
    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name'] = $r.Name; $row['ResourceGroup'] = $r.ResourceGroup; $row['Location'] = $r.Location
        $row['SnapshotSku'] = $r.SnapshotSku; $row['SizeGB'] = $r.SizeGB; $row['SourceExists'] = $r.SourceExists
        $row['CKID'] = $r.CKID; $row['AgeDays'] = $r.AgeDays; $row['EstCost_Mo'] = [decimal]$r.EstCost_Mo; $row['WasteFlag'] = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridSnaps.DataSource = $dt
    if ($gridSnaps.Columns.Contains('EstCost_Mo')) { $gridSnaps.Columns['EstCost_Mo'].DefaultCellStyle.Format = 'C2'; $gridSnaps.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight' }
    foreach ($gridRow in $gridSnaps.Rows) {
        $w = $gridRow.Cells['WasteFlag'].Value
        if ($w -eq 'WASTE') { $gridRow.DefaultCellStyle.ForeColor = $fgRed } else { $gridRow.DefaultCellStyle.ForeColor = $fgYellow }
    }
}

function Load-IPGrid {
    param([array]$Data)
    $dt = New-Object System.Data.DataTable
    @('Name','ResourceGroup','Location','IPAddress','Allocation','Sku','CKID','WasteFlag') | ForEach-Object { $dt.Columns.Add($_, [string]) | Out-Null }
    $dt.Columns.Add('EstCost_Mo', [decimal]) | Out-Null
    foreach ($r in $Data) {
        $row = $dt.NewRow()
        $row['Name'] = $r.Name; $row['ResourceGroup'] = $r.ResourceGroup; $row['Location'] = $r.Location
        $row['IPAddress'] = $r.IPAddress; $row['Allocation'] = $r.Allocation; $row['Sku'] = $r.Sku
        $row['CKID'] = $r.CKID; $row['EstCost_Mo'] = [decimal]$r.EstCost_Mo; $row['WasteFlag'] = $r.WasteFlag
        $dt.Rows.Add($row)
    }
    $gridIPs.DataSource = $dt
    if ($gridIPs.Columns.Contains('EstCost_Mo')) { $gridIPs.Columns['EstCost_Mo'].DefaultCellStyle.Format = 'C2'; $gridIPs.Columns['EstCost_Mo'].DefaultCellStyle.Alignment = 'MiddleRight' }
    foreach ($gridRow in $gridIPs.Rows) { $gridRow.DefaultCellStyle.ForeColor = $fgRed }
}

function Update-Summary {
    $vms = $script:scanData.VMs
    $od  = $script:scanData.OrphanDisks
    $sn  = $script:scanData.Snapshots
    $ips = $script:scanData.OrphanIPs

    $total = $vms.Count; $mapped = @($vms | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
    $pct = if ($total -gt 0) { [math]::Round(($mapped / $total) * 100, 1) } else { 0 }
    $running = @($vms | Where-Object { $_.PowerState -like '*running*' }).Count
    $dealloc = @($vms | Where-Object { $_.PowerState -like '*deallocated*' }).Count

    $vmActual  = ($vms | Measure-Object -Property Actual_Mo -Sum).Sum; if (-not $vmActual) { $vmActual = 0 }
    $vmCompute = ($vms | Measure-Object -Property Compute_Mo -Sum).Sum; if (-not $vmCompute) { $vmCompute = 0 }
    $vmDisk    = ($vms | Measure-Object -Property Disk_Mo -Sum).Sum; if (-not $vmDisk) { $vmDisk = 0 }
    $vmWaste   = ($vms | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property Actual_Mo -Sum).Sum; if (-not $vmWaste) { $vmWaste = 0 }
    $vmWasteN  = @($vms | Where-Object { $_.WasteFlag -eq 'WASTE' }).Count

    $odCost = ($od | Measure-Object -Property EstCost_Mo -Sum).Sum; if (-not $odCost) { $odCost = 0 }
    $snCost = ($sn | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum; if (-not $snCost) { $snCost = 0 }
    $ipCost = ($ips | Where-Object { $_.WasteFlag -eq 'WASTE' } | Measure-Object -Property EstCost_Mo -Sum).Sum; if (-not $ipCost) { $ipCost = 0 }
    $totalWaste = $vmWaste + $odCost + $snCost + $ipCost

    # Update tab labels with counts
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
Compute/Mo: $('${0:N2}' -f $vmCompute)
Disk/Mo:    $('${0:N2}' -f $vmDisk)
Actual/Mo:  $('${0:N2}' -f $vmActual)
VM Waste:   $('${0:N2}' -f $vmWaste)
            ($vmWasteN VMs)

ORPHANS
─────────────────────
Disks:    $($od.Count)  $('${0:N2}' -f $odCost)/mo
Snaps:    $($sn.Count)  $('${0:N2}' -f $snCost)/mo
IPs:      $($ips.Count)  $('${0:N2}' -f $ipCost)/mo

═════════════════════
TOTAL WASTE:
$('${0:N2}' -f $totalWaste)/mo
═════════════════════
"@
}

# ══════════════════════════════════════════════════════════════════════════════
# EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

$scanBtn.Add_Click({
    $idx = $subCombo.SelectedIndex; if ($idx -lt 0) { return }
    $sub = $allSubs[$idx]
    $scanBtn.Enabled = $false; $scanBtn.Text = "Scanning..."
    $statusLabel.Text = "Scanning $($sub.Name)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:scanData = Get-AllSubscriptionData -Subscription $sub -NoPricing:$SkipPricing -StatusLabel $statusLabel
        Load-VMGrid -Data $script:scanData.VMs
        Load-OrphanDiskGrid -Data $script:scanData.OrphanDisks
        Load-SnapshotGrid -Data $script:scanData.Snapshots
        Load-IPGrid -Data $script:scanData.OrphanIPs
        Update-Summary
        $exportBtn.Enabled = $true
        $statusLabel.Text = "Scan complete — $($script:scanData.VMs.Count) VMs, $($script:scanData.OrphanDisks.Count) orphan disks, $($script:scanData.Snapshots.Count) snapshots, $($script:scanData.OrphanIPs.Count) orphan IPs"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error:`n$($_.Exception.Message)", "Scan Error", 'OK', 'Error')
        $statusLabel.Text = "Error during scan"
    }
    finally { $scanBtn.Enabled = $true; $scanBtn.Text = "SCAN" }
})

$exportBtn.Add_Click({
    if (-not $script:scanData) { return }
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"; $saveDialog.DefaultExt = "csv"
    $subName = $allSubs[$subCombo.SelectedIndex].Name -replace '[^a-zA-Z0-9\-]', '_'
    $saveDialog.FileName = "CloudWaste_${subName}_$(Get-Date -Format 'yyyy-MM-dd')_VMs.csv"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        $basePath = $saveDialog.FileName -replace '_VMs\.csv$', ''
        if ($script:scanData.VMs.Count -gt 0) { $script:scanData.VMs | Export-Csv -Path "${basePath}_VMs.csv" -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.OrphanDisks.Count -gt 0) { $script:scanData.OrphanDisks | Export-Csv -Path "${basePath}_OrphanDisks.csv" -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.Snapshots.Count -gt 0) { $script:scanData.Snapshots | Export-Csv -Path "${basePath}_Snapshots.csv" -NoTypeInformation -Encoding UTF8 }
        if ($script:scanData.OrphanIPs.Count -gt 0) { $script:scanData.OrphanIPs | Export-Csv -Path "${basePath}_OrphanIPs.csv" -NoTypeInformation -Encoding UTF8 }
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

