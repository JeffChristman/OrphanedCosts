<#
.SYNOPSIS
    Compares two MDC assessment exports (baseline vs. refresh) and writes a
    four-sheet comparison workbook: Assessment, Summary, Comparison,
    Comparison_Summary.

.DESCRIPTION
    Both input files must use the original flat schema:
        SubscriptionName, SubscriptionId, SettingType, ExpectedState,
        ActualState, Compliance, Recommendation

    Each row is keyed on SubscriptionId + SettingType and classified as
    Improved / Regressed / Unchanged / New / Removed. A compliant row is one
    whose Compliance cell contains the check-mark glyph; anything else counts
    as non-compliant.

.PARAMETER NewPath
    The refreshed assessment export (the current run).

.PARAMETER BaselinePath
    The prior assessment export to compare against
    (e.g. DefenderForCloud_Assessment_20250707_0926.xlsx).

.NOTES
    Requires: ImportExcel   (Install-Module ImportExcel -Scope CurrentUser)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$NewPath,
    [Parameter(Mandatory)][string]$BaselinePath,
    [string]$NewSheet      = 'Assessment',
    [string]$BaselineSheet = 'Assessment',
    [string]$OutputFolder  = (Get-Location).Path
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable ImportExcel)) {
    throw "ImportExcel not installed. Run: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel

foreach ($p in $NewPath,$BaselinePath) {
    if (-not (Test-Path $p)) { throw "File not found: $p" }
}

$PASS = [char]0x2714   # check mark (ignore any trailing variation selector)
function Is-Pass($c) { "$c" -like "*$PASS*" }

$new  = Import-Excel -Path $NewPath      -WorksheetName $NewSheet
$base = Import-Excel -Path $BaselinePath -WorksheetName $BaselineSheet

$key = { param($r) "$($r.SubscriptionId)|$($r.SettingType)" }
$baseMap = @{}; foreach ($b in $base) { $baseMap[(& $key $b)] = $b }
$newMap  = @{}; foreach ($n in $new)  { $newMap[(& $key $n)]  = $n }

$comparison = foreach ($k in ($newMap.Keys + $baseMap.Keys | Select-Object -Unique)) {
    $n = $newMap[$k]; $b = $baseMap[$k]
    if     ($n -and -not $b) { $change = 'New' }
    elseif ($b -and -not $n) { $change = 'Removed' }
    else {
        $bp = Is-Pass $b.Compliance; $np = Is-Pass $n.Compliance
        $change = if ($bp -eq $np) {'Unchanged'} elseif (-not $bp -and $np) {'Improved'} else {'Regressed'}
    }
    $ref = if ($n) {$n} else {$b}
    [pscustomobject]@{
        SubscriptionName = $ref.SubscriptionName
        SettingType      = $ref.SettingType
        Baseline_State   = if ($b) {$b.ActualState} else {'(absent)'}
        New_State        = if ($n) {$n.ActualState} else {'(absent)'}
        Change           = $change
    }
}

$baseNC = ($base | Where-Object { -not (Is-Pass $_.Compliance) }).Count
$newNC  = ($new  | Where-Object { -not (Is-Pass $_.Compliance) }).Count
$cc     = $comparison | Group-Object Change -AsHashTable -AsString

$compSummary = [pscustomobject]@{
    BaselineFile          = Split-Path $BaselinePath -Leaf
    RefreshFile           = Split-Path $NewPath -Leaf
    Baseline_NonCompliant = $baseNC
    New_NonCompliant      = $newNC
    NetChange             = $newNC - $baseNC
    Improved              = ($cc['Improved']  | Measure-Object).Count
    Regressed             = ($cc['Regressed'] | Measure-Object).Count
    Unchanged             = ($cc['Unchanged'] | Measure-Object).Count
    New_Settings          = ($cc['New']       | Measure-Object).Count
    Removed_Settings      = ($cc['Removed']   | Measure-Object).Count
}

$planFree = $new | Where-Object { $_.SettingType -like 'Defender Plan*' -and $_.ActualState -eq 'Free' }
$summary = [pscustomobject]@{
    SubscriptionsAssessed = ($new | Select-Object -Unique SubscriptionId).Count
    TotalRows             = $new.Count
    NonCompliantTotal     = $newNC
    DefenderPlans_Free    = $planFree.Count
    CloudPosture_NotStd   = ($new | Where-Object {$_.SettingType -eq 'Defender Plan - CloudPosture' -and $_.ActualState -ne 'Standard'}).Count
    AutoProvisioning_Off  = ($new | Where-Object {$_.SettingType -eq 'AutoProvisioning' -and $_.ActualState -eq 'Off'}).Count
    Regulatory_Failed     = ($new | Where-Object {$_.SettingType -like 'Regulatory*' -and -not (Is-Pass $_.Compliance)}).Count
    GeneratedOn           = (Get-Date -Format 'yyyy-MM-dd HH:mm')
}

$stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$outFile = Join-Path $OutputFolder "DefenderForCloud_Comparison_$stamp.xlsx"

$new | Export-Excel -Path $outFile -WorksheetName 'Assessment' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
$summary | Export-Excel -Path $outFile -WorksheetName 'Summary' -AutoSize -BoldTopRow

$order = @{ 'Regressed'=0;'Improved'=1;'New'=2;'Removed'=3;'Unchanged'=4 }
$comparison |
    Sort-Object @{e={$order[$_.Change]}}, SubscriptionName, SettingType |
    Export-Excel -Path $outFile -WorksheetName 'Comparison' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
$compSummary | Export-Excel -Path $outFile -WorksheetName 'Comparison_Summary' -AutoSize -BoldTopRow

Write-Host "Comparison written to: $outFile" -ForegroundColor Green
$compSummary | Format-List