#Requires -Modules Az.Compute

<#
.SYNOPSIS
    VM Billing Gap Report — GUI Edition
    Displays results in a sortable/filterable DataGridView window.

.DESCRIPTION
    Same core logic as the console version but outputs to a Windows Forms GUI.
    Features:
      - Subscription picker dropdown
      - Sortable columns (click headers)
      - Filter by Governance Status
      - Color-coded rows (red=UNTAGGED, yellow=partial, green=Fully Tagged)
      - Export to CSV button
      - Summary stats panel

.PARAMETER SkipPricing
    Skip pricing lookups for faster runs.

.EXAMPLE
    .\Get-VMBillingGapReport-GUI.ps1
    .\Get-VMBillingGapReport-GUI.ps1 -SkipPricing
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$SkipPricing
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
# CORE FUNCTIONS (same logic as console version)
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
    catch { return @{ LinuxMonthly = $null; WindowsMonthly = $null } }
}

function Get-BillingGapData {
    param(
        [object]$Subscription,
        [switch]$NoPricing,
        [System.Windows.Forms.ToolStripStatusLabel]$StatusLabel
    )

    Set-AzContext -SubscriptionId $Subscription.Id -ErrorAction Stop | Out-Null
    $vms = @(Get-AzVM -Status -ErrorAction SilentlyContinue)
    if ($vms.Count -eq 0) { return @() }

    $pricingCache = @{}
    $inventory = [System.Collections.Generic.List[PSCustomObject]]::new()
    $counter = 0

    foreach ($vm in $vms) {
        $counter++
        if ($StatusLabel) {
            $pct = [math]::Round(($counter / $vms.Count) * 100)
            $StatusLabel.Text = "Processing VM $counter of $($vms.Count) ($pct%) — $($vm.Name)"
            [System.Windows.Forms.Application]::DoEvents()
        }

        $ckid  = Get-TagValue -Tags $vm.Tags -KeyVariants @('CKID','ckid','CkId','Ckid')
        $emass = Get-TagValue -Tags $vm.Tags -KeyVariants @('EMASS','eMASS','emass','Emass','EMASNumber','EMASS_Number','emass_number')
        $vasi  = Get-TagValue -Tags $vm.Tags -KeyVariants @('VASI','vasi','Vasi','VASINumber','VASI_Number','vasinumber')

        $ckidStatus = if ($ckid) { 'Mapped' } else { 'UNMAPPED' }
        $govStatus = if ($ckid -and $emass)          { 'Fully Tagged' }
                     elseif ($ckid -and -not $emass)  { 'Billing Only' }
                     elseif (-not $ckid -and $emass)  { 'ATO Only' }
                     else                              { 'UNTAGGED' }

        $vmSize = $vm.HardwareProfile.VmSize
        $monthlyEst = 0

        if (-not $NoPricing) {
            $cacheKey = "$vmSize|$($vm.Location)"
            if (-not $pricingCache.ContainsKey($cacheKey)) {
                $pricingCache[$cacheKey] = Get-VMPricing -VmSize $vmSize -Region $vm.Location
            }
            $pricing = $pricingCache[$cacheKey]
            $isWindowsOS = $vm.StorageProfile.OsDisk.OsType -eq 'Windows'
            $price = if ($isWindowsOS) { $pricing.WindowsMonthly } else { $pricing.LinuxMonthly }
            if ($price) { $monthlyEst = $price }
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
            VMName            = $vm.Name
            ResourceGroup     = $vm.ResourceGroupName
            Location          = $vm.Location
            VMSize            = $vmSize
            OSType            = $vm.StorageProfile.OsDisk.OsType
            PowerState        = $powerState
            GovernanceStatus  = $govStatus
            CKID_Status       = $ckidStatus
            CKID              = if ($ckid) { $ckid } else { '' }
            EMASS             = if ($emass) { $emass } else { '' }
            VASI              = if ($vasi) { $vasi } else { '' }
            EstMonthly_USD    = $monthlyEst
        })
    }
    return $inventory
}

# ══════════════════════════════════════════════════════════════════════════════
# CONNECT + GET SUBSCRIPTIONS
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "[*] Connecting to Azure Government..." -ForegroundColor Cyan
$context = Confirm-AzGovConnection
Write-Host "[+] Connected as: $($context.Account.Id)" -ForegroundColor Green

$allSubs = @(Get-AzSubscription -ErrorAction Stop |
    Where-Object { $_.State -eq 'Enabled' } | Sort-Object Name)
Write-Host "[+] Found $($allSubs.Count) subscriptions" -ForegroundColor Green

# ══════════════════════════════════════════════════════════════════════════════
# BUILD THE GUI
# ══════════════════════════════════════════════════════════════════════════════

# ── Main Form ─────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text = "CTS VM Billing & Governance Gap Report"
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

# ── Top Panel (controls) ─────────────────────────────────────────────────────
$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'
$topPanel.Height = 95
$topPanel.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
$topPanel.Padding = New-Object System.Windows.Forms.Padding(10)

# Title
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "VM BILLING & GOVERNANCE GAP REPORT"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 150)
$titleLabel.Location = New-Object System.Drawing.Point(12, 5)
$titleLabel.AutoSize = $true
$topPanel.Controls.Add($titleLabel)

# LADDER notice
$ladderLabel = New-Object System.Windows.Forms.Label
$ladderLabel.Text = "LADDER FROZEN — CAM CMDB Migration In Progress"
$ladderLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$ladderLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 200, 50)
$ladderLabel.Location = New-Object System.Drawing.Point(14, 30)
$ladderLabel.AutoSize = $true
$topPanel.Controls.Add($ladderLabel)

# Subscription label
$subLabel = New-Object System.Windows.Forms.Label
$subLabel.Text = "Subscription:"
$subLabel.Location = New-Object System.Drawing.Point(12, 58)
$subLabel.AutoSize = $true
$subLabel.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($subLabel)

# Subscription dropdown
$subCombo = New-Object System.Windows.Forms.ComboBox
$subCombo.Location = New-Object System.Drawing.Point(110, 55)
$subCombo.Size = New-Object System.Drawing.Size(500, 25)
$subCombo.DropDownStyle = 'DropDownList'
$subCombo.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$subCombo.ForeColor = [System.Drawing.Color]::White
$subCombo.FlatStyle = 'Flat'
foreach ($sub in $allSubs) {
    $subCombo.Items.Add("$($sub.Name)  |  $($sub.Id)") | Out-Null
}
if ($subCombo.Items.Count -gt 0) { $subCombo.SelectedIndex = 0 }
$topPanel.Controls.Add($subCombo)

# Scan button
$scanBtn = New-Object System.Windows.Forms.Button
$scanBtn.Text = "SCAN"
$scanBtn.Location = New-Object System.Drawing.Point(625, 53)
$scanBtn.Size = New-Object System.Drawing.Size(90, 28)
$scanBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 150, 110)
$scanBtn.ForeColor = [System.Drawing.Color]::White
$scanBtn.FlatStyle = 'Flat'
$scanBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$scanBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$topPanel.Controls.Add($scanBtn)

# Filter label
$filterLabel = New-Object System.Windows.Forms.Label
$filterLabel.Text = "Filter:"
$filterLabel.Location = New-Object System.Drawing.Point(735, 58)
$filterLabel.AutoSize = $true
$filterLabel.ForeColor = [System.Drawing.Color]::White
$topPanel.Controls.Add($filterLabel)

# Filter dropdown
$filterCombo = New-Object System.Windows.Forms.ComboBox
$filterCombo.Location = New-Object System.Drawing.Point(780, 55)
$filterCombo.Size = New-Object System.Drawing.Size(170, 25)
$filterCombo.DropDownStyle = 'DropDownList'
$filterCombo.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$filterCombo.ForeColor = [System.Drawing.Color]::White
$filterCombo.FlatStyle = 'Flat'
@('All VMs', 'UNTAGGED', 'Billing Only', 'ATO Only', 'Fully Tagged', 'UNMAPPED (no CKID)') |
    ForEach-Object { $filterCombo.Items.Add($_) | Out-Null }
$filterCombo.SelectedIndex = 0
$topPanel.Controls.Add($filterCombo)

# Export button
$exportBtn = New-Object System.Windows.Forms.Button
$exportBtn.Text = "EXPORT CSV"
$exportBtn.Location = New-Object System.Drawing.Point(970, 53)
$exportBtn.Size = New-Object System.Drawing.Size(110, 28)
$exportBtn.BackColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$exportBtn.ForeColor = [System.Drawing.Color]::White
$exportBtn.FlatStyle = 'Flat'
$exportBtn.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$exportBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
$exportBtn.Enabled = $false
$topPanel.Controls.Add($exportBtn)

$form.Controls.Add($topPanel)

# ── Summary Panel (right side stats) ──────────────────────────────────────────
$summaryPanel = New-Object System.Windows.Forms.Panel
$summaryPanel.Dock = 'Right'
$summaryPanel.Width = 280
$summaryPanel.BackColor = [System.Drawing.Color]::FromArgb(38, 38, 42)
$summaryPanel.Padding = New-Object System.Windows.Forms.Padding(15, 10, 15, 10)

$summaryTitle = New-Object System.Windows.Forms.Label
$summaryTitle.Text = "SUMMARY"
$summaryTitle.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$summaryTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 150)
$summaryTitle.Location = New-Object System.Drawing.Point(15, 10)
$summaryTitle.AutoSize = $true
$summaryPanel.Controls.Add($summaryTitle)

$summaryText = New-Object System.Windows.Forms.Label
$summaryText.Text = "Select a subscription and click SCAN"
$summaryText.Font = New-Object System.Drawing.Font("Consolas", 9)
$summaryText.ForeColor = [System.Drawing.Color]::FromArgb(200, 200, 200)
$summaryText.Location = New-Object System.Drawing.Point(15, 40)
$summaryText.Size = New-Object System.Drawing.Size(250, 500)
$summaryPanel.Controls.Add($summaryText)

# ── DataGridView ──────────────────────────────────────────────────────────────
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$grid.GridColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$grid.BorderStyle = 'None'
$grid.CellBorderStyle = 'SingleHorizontal'
$grid.ColumnHeadersBorderStyle = 'Single'
$grid.EnableHeadersVisualStyles = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.RowHeadersVisible = $false
$grid.AllowUserToResizeRows = $false

# Header style
$grid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 55)
$grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 150)
$grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(50, 50, 55)
$grid.ColumnHeadersDefaultCellStyle.Alignment = 'MiddleLeft'
$grid.ColumnHeadersHeight = 32

# Row style
$grid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$grid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$grid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 100, 80)
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)

$grid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(38, 38, 42)

# ── Status Bar ────────────────────────────────────────────────────────────────
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 90)
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready — select a subscription and click SCAN"
$statusLabel.ForeColor = [System.Drawing.Color]::White
$statusBar.Items.Add($statusLabel) | Out-Null

# ── Add controls in REVERSE dock order (Fill added first, edges last) ─────────
# WinForms docking rule: last control added docks first.
# So we add: grid (Fill) → summaryPanel (Right) → topPanel (Top) → statusBar (Bottom)
$form.Controls.Add($statusBar)
$form.Controls.Add($grid)
$form.Controls.Add($summaryPanel)
$form.Controls.Add($topPanel)

# ══════════════════════════════════════════════════════════════════════════════
# DATA STORE + HELPERS
# ══════════════════════════════════════════════════════════════════════════════

$script:allResults = @()

function Update-Grid {
    param([string]$Filter)

    $filtered = switch ($Filter) {
        'UNTAGGED'          { $script:allResults | Where-Object { $_.GovernanceStatus -eq 'UNTAGGED' } }
        'Billing Only'      { $script:allResults | Where-Object { $_.GovernanceStatus -eq 'Billing Only' } }
        'ATO Only'          { $script:allResults | Where-Object { $_.GovernanceStatus -eq 'ATO Only' } }
        'Fully Tagged'      { $script:allResults | Where-Object { $_.GovernanceStatus -eq 'Fully Tagged' } }
        'UNMAPPED (no CKID)' { $script:allResults | Where-Object { $_.CKID_Status -eq 'UNMAPPED' } }
        default             { $script:allResults }
    }

    $dt = New-Object System.Data.DataTable
    @('VMName','ResourceGroup','Location','VMSize','OSType','PowerState',
      'GovernanceStatus','CKID_Status','CKID','EMASS','VASI') |
        ForEach-Object { $dt.Columns.Add($_, [string]) | Out-Null }
    $dt.Columns.Add('EstMonthly_USD', [decimal]) | Out-Null

    foreach ($r in $filtered) {
        $row = $dt.NewRow()
        $row['VMName']           = $r.VMName
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
        $row['EstMonthly_USD']   = [decimal]$r.EstMonthly_USD
        $dt.Rows.Add($row)
    }

    $grid.DataSource = $dt

    # Format cost column
    if ($grid.Columns.Contains('EstMonthly_USD')) {
        $grid.Columns['EstMonthly_USD'].DefaultCellStyle.Format = 'C2'
        $grid.Columns['EstMonthly_USD'].DefaultCellStyle.Alignment = 'MiddleRight'
    }

    # Color-code rows by GovernanceStatus
    foreach ($gridRow in $grid.Rows) {
        $gov = $gridRow.Cells['GovernanceStatus'].Value
        switch ($gov) {
            'UNTAGGED' {
                $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 100, 100)
            }
            'Billing Only' {
                $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 220, 100)
            }
            'ATO Only' {
                $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(255, 180, 80)
            }
            'Fully Tagged' {
                $gridRow.DefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(100, 255, 160)
            }
        }
    }

    $statusLabel.Text = "Showing $($filtered.Count) of $($script:allResults.Count) VMs"
}

function Update-Summary {
    $total      = $script:allResults.Count
    $mapped     = @($script:allResults | Where-Object { $_.CKID_Status -eq 'Mapped' }).Count
    $unmapped   = $total - $mapped
    $mappedPct  = if ($total -gt 0) { [math]::Round(($mapped / $total) * 100, 1) } else { 0 }
    $hasEmass   = @($script:allResults | Where-Object { $_.EMASS -ne '' }).Count

    $fully      = @($script:allResults | Where-Object { $_.GovernanceStatus -eq 'Fully Tagged' }).Count
    $billOnly   = @($script:allResults | Where-Object { $_.GovernanceStatus -eq 'Billing Only' }).Count
    $atoOnly    = @($script:allResults | Where-Object { $_.GovernanceStatus -eq 'ATO Only' }).Count
    $none       = @($script:allResults | Where-Object { $_.GovernanceStatus -eq 'UNTAGGED' }).Count

    $running    = @($script:allResults | Where-Object { $_.PowerState -like '*running*' }).Count
    $dealloc    = @($script:allResults | Where-Object { $_.PowerState -like '*deallocated*' }).Count
    $stopped    = @($script:allResults | Where-Object { $_.PowerState -like '*stopped*' -and $_.PowerState -notlike '*deallocated*' }).Count

    $totalCost  = ($script:allResults | Measure-Object -Property EstMonthly_USD -Sum).Sum
    if (-not $totalCost) { $totalCost = 0 }
    $unmapCost  = ($script:allResults | Where-Object { $_.CKID_Status -eq 'UNMAPPED' } | Measure-Object -Property EstMonthly_USD -Sum).Sum
    if (-not $unmapCost) { $unmapCost = 0 }
    $decommCount = @($script:allResults | Where-Object {
        $_.CKID_Status -eq 'UNMAPPED' -and
        ($_.PowerState -like '*deallocated*' -or $_.PowerState -like '*stopped*')
    }).Count

    $summaryText.Text = @"
BILLING (CKID)
───────────────────
Total VMs:       $total
CKID Mapped:     $mapped ($mappedPct%)
CKID Unmapped:   $unmapped

GOVERNANCE
───────────────────
Fully Tagged:    $fully
Billing Only:    $billOnly
ATO Only:        $atoOnly
UNTAGGED:        $none
EMASS Present:   $hasEmass

POWER STATE
───────────────────
Running:         $running
Deallocated:     $dealloc
Stopped:         $stopped

COST
───────────────────
Total/Mo:    $('${0:N2}' -f $totalCost)
Unmapped/Mo: $('${0:N2}' -f $unmapCost)

Decomm Candidates: $decommCount
(stopped + no CKID)
"@
}

# ══════════════════════════════════════════════════════════════════════════════
# EVENT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

# ── SCAN button ───────────────────────────────────────────────────────────────
$scanBtn.Add_Click({
    $selectedIndex = $subCombo.SelectedIndex
    if ($selectedIndex -lt 0) { return }

    $sub = $allSubs[$selectedIndex]
    $scanBtn.Enabled = $false
    $scanBtn.Text = "Scanning..."
    $statusLabel.Text = "Scanning $($sub.Name)..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:allResults = @(Get-BillingGapData -Subscription $sub -NoPricing:$SkipPricing -StatusLabel $statusLabel)
        Update-Summary
        Update-Grid -Filter $filterCombo.SelectedItem.ToString()
        $exportBtn.Enabled = ($script:allResults.Count -gt 0)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error scanning subscription:`n$($_.Exception.Message)",
            "Scan Error",
            'OK', 'Error'
        )
        $statusLabel.Text = "Error during scan"
    }
    finally {
        $scanBtn.Enabled = $true
        $scanBtn.Text = "SCAN"
    }
})

# ── Filter dropdown ───────────────────────────────────────────────────────────
$filterCombo.Add_SelectedIndexChanged({
    if ($script:allResults.Count -gt 0) {
        Update-Grid -Filter $filterCombo.SelectedItem.ToString()
    }
})

# ── Export button ─────────────────────────────────────────────────────────────
$exportBtn.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveDialog.DefaultExt = "csv"

    $subName = if ($script:allResults.Count -gt 0) {
        $allSubs[$subCombo.SelectedIndex].Name -replace '[^a-zA-Z0-9\-]', '_'
    } else { 'export' }
    $saveDialog.FileName = "BillingGap_${subName}_$(Get-Date -Format 'yyyy-MM-dd').csv"

    if ($saveDialog.ShowDialog() -eq 'OK') {
        $script:allResults | Select-Object `
            VMName, ResourceGroup, Location, VMSize, OSType, PowerState,
            GovernanceStatus, CKID_Status, CKID, EMASS, VASI,
            @{ Name='EstMonthly_USD'; Expression={ '{0:N2}' -f $_.EstMonthly_USD } } |
            Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8

        $statusLabel.Text = "Exported to: $($saveDialog.FileName)"
        [System.Windows.Forms.MessageBox]::Show(
            "Exported $($script:allResults.Count) VMs to:`n$($saveDialog.FileName)",
            "Export Complete", 'OK', 'Information'
        )
    }
})

# ══════════════════════════════════════════════════════════════════════════════
# LAUNCH
# ══════════════════════════════════════════════════════════════════════════════

Write-Host "[+] Launching GUI..." -ForegroundColor Green
[void]$form.ShowDialog()
$form.Dispose()
