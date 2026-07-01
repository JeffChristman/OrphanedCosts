<#
.SYNOPSIS
    Generates an Excel (.xlsx) workbook of the Sentinel current-state findings:
    a Summary sheet and a sortable/filterable Findings sheet with color-coded
    status, frozen header, and autofilter.

.REQUIREMENTS
    ImportExcel module (no Excel install needed):
        Install-Module ImportExcel -Scope CurrentUser

.USAGE
    .\Make-FindingsWorkbook.ps1
    .\Make-FindingsWorkbook.ps1 -OutputPath C:\repo\Sentinel-Assessment-Findings.xlsx
#>

[CmdletBinding()]
param(
    [string] $OutputPath = ".\Sentinel-Assessment-Findings.xlsx"
)

$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    throw "ImportExcel module not found. Run: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

# ---- Findings (matches the report + dashboard) ----
$findings = @(
    [pscustomobject]@{ 'Finding Area'='Automation Rules'; 'Then (Dec 2025)'='"Not yet widely implemented"'; 'Now (Live)'='0 configured'; Status='GAP'; Category='Action Required'; Baseline='CIS Azure 5 / NIST IR-4(1),SI-4 / MS SOAR'; Recommendation='Implement automation rules to trigger the existing playbook library; standardize routing, enrichment, tagging, escalation.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='UEBA'; 'Then (Dec 2025)'='"Fully enabled" in Government'; 'Now (Live)'='Sources attached; enabled flag unconfirmed'; Status='VERIFY'; Category='Action Required'; Baseline='CIS Azure 1.x / NIST AU-6(3),SI-4'; Recommendation='Verify UEBA enablement in portal; confirm AzureAD/Signin/Audit/endpoint sources feed entity profiles.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='Enterprise Log Forwarding'; 'Then (Dec 2025)'='Not assessed'; 'Now (Live)'='106 of 106 subscriptions forwarding'; Status='HEALTHY'; Category='New - Strength'; Baseline='CIS Azure 5.1.x / NIST AU-2,AU-12'; Recommendation='Enforce via Azure Policy (DeployIfNotExists) so new subscriptions inherit forwarding automatically.'; 'New Area'='Yes' }
    [pscustomobject]@{ 'Finding Area'='Operational Maturity Inventory'; 'Then (Dec 2025)'='Not assessed'; 'Now (Live)'='112 playbooks; 166 RBAC; 80 workbooks; 19 watchlists; 19 hunting queries'; Status='HEALTHY'; Category='New - Strength'; Baseline='MS Sentinel best practice'; Recommendation='Reframes maturity: focus next phase on orchestration/verification, not new capability.'; 'New Area'='Yes' }
    [pscustomobject]@{ 'Finding Area'='Threat Intelligence'; 'Then (Dec 2025)'='Not assessed'; 'Now (Live)'='0 indicators returned (30d)'; Status='VERIFY'; Category='New - Verify'; Baseline='MS TI / NIST PM-16,SI-5'; Recommendation='Confirm a TI feed is active; indicators may land in a different table than queried.'; 'New Area'='Yes' }
    [pscustomobject]@{ 'Finding Area'='Playbook Enablement'; 'Then (Dec 2025)'='Not assessed'; 'Now (Live)'='0 of 112 enabled returned'; Status='VERIFY'; Category='New - Verify'; Baseline='MS SOAR / NIST IR-4(1)'; Recommendation='Confirm how many playbooks are active; 112 disabled is operationally implausible (likely artifact).'; 'New Area'='Yes' }
    [pscustomobject]@{ 'Finding Area'='Content Hub Currency'; 'Then (Dec 2025)'='"Some updates available"'; 'Now (Live)'='0 packages returned'; Status='VERIFY'; Category='Verify'; Baseline='MS content / NIST SI-3,SI-4'; Recommendation='Verify installed content; establish monthly update cadence, test in non-prod before Gov rollout.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='Data Connectors & Ingestion'; 'Then (Dec 2025)'='"A few to refine"'; 'Now (Live)'='105 connectors; core tables fresh'; Status='HEALTHY'; Category='Healthy'; Baseline='CIS Azure 5.1.x / NIST AU-2,AU-12'; Recommendation='Maintain routine connector-health monitoring.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='Analytics & Alerting'; 'Then (Dec 2025)'='"Functional"'; 'Now (Live)'='54 enabled rules; ~19 identity/access'; Status='HEALTHY'; Category='Healthy'; Baseline='CIS Azure 5.2.x+1.x / NIST SI-4,AU-6'; Recommendation='Expand identity/privilege/access detections; tune thresholds on CSOC feedback.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='Summary Rules'; 'Then (Dec 2025)'='"Not broadly deployed"'; 'Now (Live)'='0 configured'; Status='HEALTHY'; Category='Acceptable'; Baseline='MS Sentinel perf/cost / NIST AU-6'; Recommendation='Optional; deploy on highest-volume tables for hunt speed where Gov parity allows.'; 'New Area'='No' }
    [pscustomobject]@{ 'Finding Area'='Workspace Manager'; 'Then (Dec 2025)'='Absence appropriate'; 'Now (Live)'='Not configured'; Status='HEALTHY'; Category='Healthy (as expected)'; Baseline='Governance / NIST CM-2'; Recommendation='Agree with prior report; track Preview->GA and reassess.'; 'New Area'='No' }
)

# ---- Summary sheet ----
$summary = @(
    [pscustomobject]@{ Metric='Confirmed gaps (action required)'; Value='1' }
    [pscustomobject]@{ Metric='Items needing verification';        Value='4' }
    [pscustomobject]@{ Metric='Newly discovered areas';            Value='4' }
    [pscustomobject]@{ Metric='Healthy / acceptable areas';        Value='5' }
    [pscustomobject]@{ Metric='Subscriptions forwarding to central WS'; Value='106 of 106' }
    [pscustomobject]@{ Metric=''; Value='' }
    [pscustomobject]@{ Metric='Priority action'; Value='Implement automation rules to activate the 112-playbook library (currently cannot be triggered automatically).' }
    [pscustomobject]@{ Metric=''; Value='' }
    [pscustomobject]@{ Metric='Legend: GAP';     Value='Confirmed deficiency - action required' }
    [pscustomobject]@{ Metric='Legend: VERIFY';  Value='Confirm in portal before concluding' }
    [pscustomobject]@{ Metric='Legend: HEALTHY'; Value='No action or routine maintenance' }
)

# Write Summary sheet
$summary | Export-Excel -Path $OutputPath -WorksheetName 'Summary' -AutoSize -Title 'Microsoft Sentinel - Current-State Assessment (FOUO / PRE-DECISIONAL)' -TitleBold

# Write Findings sheet with table, autofilter, frozen header
$xl = $findings | Export-Excel -Path $OutputPath -WorksheetName 'Findings' -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow -PassThru

$sheet = $xl.Workbook.Worksheets['Findings']
$rows  = $findings.Count + 1   # +1 for header

# Color-code the Status column (D). Find its index dynamically.
$statusCol = 4
for ($r = 2; $r -le $rows; $r++) {
    $val = $sheet.Cells[$r, $statusCol].Value
    switch ($val) {
        'GAP'     { $bg='#F7C1C1'; $fg='#7d1414' }
        'VERIFY'  { $bg='#FAC775'; $fg='#633806' }
        'HEALTHY' { $bg='#C0DD97'; $fg='#173404' }
        default   { $bg=$null }
    }
    if ($bg) {
        $cell = $sheet.Cells[$r, $statusCol]
        $cell.Style.Fill.PatternType = 'Solid'
        $cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($bg))
        $cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml($fg))
        $cell.Style.Font.Bold = $true
        $cell.Style.HorizontalAlignment = 'Center'
    }
}

# Wrap text on the longer columns (Then, Now, Recommendation)
foreach ($c in 2,3,7) {
    for ($r = 1; $r -le $rows; $r++) { $sheet.Cells[$r,$c].Style.WrapText = $true }
    $sheet.Column($c).Width = 40
}

Close-ExcelPackage $xl

Write-Host "Workbook written: $OutputPath" -ForegroundColor Green
Write-Host "Sheets: Summary, Findings (sortable + filterable, color-coded status)"