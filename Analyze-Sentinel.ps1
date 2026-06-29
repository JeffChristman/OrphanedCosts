<#
.SYNOPSIS
    Analyzes the raw JSON produced by the Sentinel collector (central-workspace
    edition) and produces a shareable HTML report. Pure analysis, no Azure calls.

.DESCRIPTION
    Consumes Sentinel-Raw-*.json and:
      1. Evaluates live Sentinel configuration against a BLENDED baseline:
            - CIS Microsoft Azure Foundations Benchmark
            - Microsoft Sentinel best practices
            - NIST SP 800-53 Rev 5 (AU, SI, IR families) for VA/FedRAMP
      2. Maps each live finding back to the corresponding claim in the original
         VA CSOC Toolset Configuration Report (Dec 22, 2025).
      3. Flags where the report's "optimization opportunity" framing understated
         an actual gap (report tone vs. observed reality).

.EXAMPLE
    .\Analyze-Sentinel.ps1 -InputPath .\Sentinel-Raw-vaecla-security-gov-20260625-101101.json
.EXAMPLE
    .\Analyze-Sentinel.ps1 -InputPath Sentinel-Raw-vaecla-security-gov-20260625-101101.json -OutputPath .\delta.html

.NOTES
    PowerShell 5.1+ or 7+. No modules required.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $InputPath,
    [string] $OutputPath,
    [int]    $MinAnalyticsRules = 25   # blended baseline expectation; tune to VA policy
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $InputPath)) { throw "Input file not found: $InputPath" }
if (-not $OutputPath) {
    $OutputPath = [System.IO.Path]::ChangeExtension($InputPath, $null).TrimEnd('.') + '-DELTA.html'
}

# Read JSON (handle BOM)
$raw = Get-Content -Path $InputPath -Raw -Encoding UTF8
$doc = $raw | ConvertFrom-Json
$ws  = $doc.workspace
if (-not $ws) { throw "No 'workspace' object in JSON. Was this produced by the central-workspace collector?" }

# ---------------------------------------------------------------------------
# Report claims (sections 3.1-3.7 of the VA CSOC report) + tone
# ---------------------------------------------------------------------------
$ReportClaims = @{
    automation_rules  = @{ section='3.1 Absence of Automation Rules';            said="Automation rules 'have not yet been widely implemented'; framed as expected in maturing deployments."; tone='optimization' }
    summary_rules     = @{ section='3.2 Lack of Summary Rules';                  said="Summary rules 'have not been broadly deployed'; called 'not required for normal functionality.'"; tone='optimization' }
    data_connectors   = @{ section='3.3 Misconfigured Data Connectors';          said="'A small number of connectors can be further refined'; framed as typical optimization."; tone='optimization' }
    workspace_manager = @{ section='3.4 Workspace Manager in Preview';           said="Absence is 'expected and appropriate' due to Preview status under FedRAMP; monitor only."; tone='appropriate' }
    alerting          = @{ section='3.5 Limited Alerting';                       said="'Functional analytics and alerting in place'; opportunity to expand coverage."; tone='optimization' }
    ueba              = @{ section='3.6 UEBA Not Fully Configured';              said="UEBA 'fully enabled' in Government; onboarding underway for Commercial."; tone='optimization' }
    content_hub       = @{ section='3.7 Content Hub Not Fully Updated';          said="'Some packages remain available for update'; framed as normal."; tone='optimization' }
    forwarding        = @{ section='(not covered in report)';                    said='The report does not assess whether all subscriptions forward logs to the central workspace.'; tone='absent' }
}

$SeverityOrder = @{ CRITICAL=0; GAP=1; REVIEW=2; PASS=3; INFO=4 }

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Get-Count($x) {
    if ($null -eq $x) { return 0 }
    if ($x -is [array]) { return $x.Count }
    return 1
}
function Get-SectionError($workspace, [string]$section) {
    foreach ($e in @($workspace.collectionErrors)) {
        if ($e.section -eq $section) { return $e.error }
    }
    return $null
}
function New-Finding {
    param($Area,$ClaimKey,$Baseline,$Expected,$Observed,$Status,[bool]$Understated,$Recommendation)
    $c = $ReportClaims[$ClaimKey]
    [pscustomobject]@{
        Area           = $Area
        ReportSection  = $c.section
        ReportSaid     = $c.said
        ReportTone     = $c.tone
        Baseline       = $Baseline
        Expected       = $Expected
        Observed       = $Observed
        Status         = $Status
        Understated    = $Understated
        Recommendation = $Recommendation
    }
}

# ---------------------------------------------------------------------------
# Evaluators
# ---------------------------------------------------------------------------
$findings = [System.Collections.Generic.List[object]]::new()

# 1. Automation Rules
$err = Get-SectionError $ws 'automationRules'
$n = Get-Count $ws.automationRules
if ($err) {
    $findings.Add( (New-Finding 'Automation Rules' 'automation_rules' `
        'CIS Azure 5 / NIST IR-4(1), SI-4 / MS SOAR best practice' `
        'Automation rules orchestrating playbooks' "Could not enumerate ($err)" 'REVIEW' $false `
        'Verify Security Reader access and Az.SecurityInsights version.') )
} elseif ($n -eq 0) {
    $findings.Add( (New-Finding 'Automation Rules' 'automation_rules' `
        'CIS Azure 5 / NIST IR-4(1), SI-4 / MS SOAR best practice' `
        'At least baseline automation rules to trigger playbooks' '0 automation rules configured' 'GAP' $true `
        "Report framed this as a maturity 'opportunity,' but zero automation rules means playbooks are not orchestrated at all - manual triage only. Introduce rules for routing, tagging, enrichment, escalation.") )
} else {
    $findings.Add( (New-Finding 'Automation Rules' 'automation_rules' `
        'CIS Azure 5 / NIST IR-4(1), SI-4 / MS SOAR best practice' `
        'Automation rules orchestrating playbooks' "$n automation rule(s) configured" 'PASS' $false `
        'Continue phasing in higher-impact automated response actions.') )
}

# 2. Summary Rules
$err = Get-SectionError $ws 'summaryRules'
$n = Get-Count $ws.summaryRules
if ($err) {
    $findings.Add( (New-Finding 'Summary Rules' 'summary_rules' `
        'MS Sentinel best practice (cost/perf) / NIST AU-6' `
        'Summary rules on high-volume sources' "API unavailable / not returned ($err)" 'REVIEW' $false `
        'Confirm summary-rule availability in current Gov region.') )
} else {
    $status = if ($n -gt 0) { 'PASS' } else { 'REVIEW' }
    $findings.Add( (New-Finding 'Summary Rules' 'summary_rules' `
        'MS Sentinel best practice (cost/perf) / NIST AU-6' `
        'Summary rules on Sign-in/Audit/Firewall-class high-volume logs' "$n summary rule(s) found" $status $false `
        'Deploy where Gov parity allows for investigation speed and cost control.') )
}

# 3. Data Connectors & Ingestion
$err = Get-SectionError $ws 'dataConnectors'
$conns = @($ws.dataConnectors)
$n = $conns.Count
$core = @('SigninLogs','AuditLogs','SecurityEvent','CommonSecurityLog','AzureActivity','Syslog')
$seen = @($ws.ingestionFreshness | ForEach-Object { $_._Tbl })
$coreLive = @($core | Where-Object { $seen -contains $_ })
if ($err) {
    $findings.Add( (New-Finding 'Data Connectors & Ingestion' 'data_connectors' `
        'CIS Azure 5.1.x / NIST AU-2, AU-12 / MS connector health' `
        'Connectors healthy and core security tables ingesting continuously' `
        "Connector enumeration error ($err); $($coreLive.Count) core tables fresh" 'REVIEW' $false `
        'Validate diagnostic settings, connector identity permissions, and API scopes.') )
} elseif ($coreLive.Count -ge 3) {
    $findings.Add( (New-Finding 'Data Connectors & Ingestion' 'data_connectors' `
        'CIS Azure 5.1.x / NIST AU-2, AU-12 / MS connector health' `
        'Connectors healthy and core security tables ingesting continuously' `
        "$n connector(s); core tables fresh: $($coreLive -join ', ')" 'PASS' $false `
        'Maintain connector health monitoring.') )
} else {
    $findings.Add( (New-Finding 'Data Connectors & Ingestion' 'data_connectors' `
        'CIS Azure 5.1.x / NIST AU-2, AU-12 / MS connector health' `
        'Connectors healthy and core security tables ingesting continuously' `
        "$n connector(s) onboarded but only $($coreLive.Count) core table(s) ingesting recently: $(if($coreLive){$coreLive -join ', '}else{'none'})" 'GAP' $true `
        "Report called this 'a small number of connectors to refine.' If core tables aren't ingesting, detection coverage has real holes - validate diagnostic settings, connector identity permissions, and API scopes.") )
}

# 4. Workspace Manager
$observed = if ($ws.workspaceManager) { 'Onboarding state present' } else { 'Not configured' }
$findings.Add( (New-Finding 'Workspace Manager (Preview)' 'workspace_manager' `
    'Governance / NIST CM-2 (baseline config)' `
    'Monitor only until GA & FedRAMP-validated for Gov' $observed 'INFO' $false `
    'Agree with report: absence is appropriate under FedRAMP. Track Preview->GA and reassess for multi-workspace governance at that point.') )

# 5. Analytics / Alerting
$err = Get-SectionError $ws 'analyticsRules'
$rules = @($ws.analyticsRules)
$enabled = @($rules | Where-Object { $_.Enabled })
$kw = @('sign-in','signin','privile','escalat','mfa','conditional access','owner role','key vault','nsg','firewall','impossible travel','anomal','access')
$cisCovered = @($enabled | Where-Object {
    $name = "$($_.DisplayName)".ToLower()
    ($kw | Where-Object { $name -like "*$_*" }).Count -gt 0
})
if ($err) {
    $findings.Add( (New-Finding 'Analytics / Alerting Coverage' 'alerting' `
        'CIS Azure 5.2.x + 1.x / NIST SI-4, AU-6 / MS analytics best practice' `
        ">= $MinAnalyticsRules enabled rules incl. identity/privilege/access detections" `
        "Could not enumerate ($err)" 'REVIEW' $false 'Verify Security Reader access.') )
} else {
    $nEn = $enabled.Count
    if ($nEn -ge $MinAnalyticsRules -and $cisCovered.Count -gt 0) { $status='PASS'; $under=$false }
    elseif ($nEn -eq 0) { $status='CRITICAL'; $under=$true }
    else { $status='GAP'; $under=$true }
    $findings.Add( (New-Finding 'Analytics / Alerting Coverage' 'alerting' `
        'CIS Azure 5.2.x + 1.x / NIST SI-4, AU-6 / MS analytics best practice' `
        ">= $MinAnalyticsRules enabled rules incl. identity/privilege/access detections" `
        "$nEn enabled rule(s); $($cisCovered.Count) match identity/access categories" $status $under `
        "Report described alerting as 'functional' with room to expand. If enabled rule count or identity/privilege coverage is low, that is a detection gap, not just tuning - add identity-anomaly, privilege-escalation, access rules.") )
}

# 6. UEBA
$err = Get-SectionError $ws 'ueba'
$ueba = $ws.ueba
if ($err -or $null -eq $ueba) {
    $findings.Add( (New-Finding 'UEBA' 'ueba' `
        'CIS Azure 1.x (identity) / NIST AU-6(3), SI-4 / MS UEBA best practice' `
        'UEBA enabled with AzureAD/Signin/Audit + endpoint sources' `
        "Not enabled or not detected ($(if($err){$err}else{'no settings returned'}))" 'GAP' $true `
        "Report states UEBA is 'fully enabled' in Government. If settings can't be confirmed enabled on this workspace, that directly contradicts the report claim - verify and reconcile.") )
} else {
    $enabledVal = $ueba.enabled
    $sources = $ueba.dataSources
    $srcTxt = if ($sources -is [array]) { $sources -join ', ' } else { "$sources" }
    $status = if ($enabledVal) { 'PASS' } else { 'GAP' }
    $under = -not [bool]$enabledVal
    $findings.Add( (New-Finding 'UEBA' 'ueba' `
        'CIS Azure 1.x (identity) / NIST AU-6(3), SI-4 / MS UEBA best practice' `
        'UEBA enabled with identity + endpoint sources feeding entity profiles' `
        "enabled=$enabledVal; sources=$srcTxt" $status $under `
        'Confirm AzureAD, Signin, Audit, and endpoint logs all contribute; exclude high-volume service/non-human accounts to reduce noise.') )
}

# 7. Content Hub
$err = Get-SectionError $ws 'contentPackages'
$pkgs = @($ws.contentPackages)
$installed = @($pkgs | Where-Object { $_.installedVersion })
$stale = @($installed | Where-Object { $_.installedVersion -ne $_.latestVersion -and $_.latestVersion })
if ($err) {
    $findings.Add( (New-Finding 'Content Hub Currency' 'content_hub' `
        'MS Sentinel content best practice / NIST SI-3, SI-4 (detection currency)' `
        'Installed Content Hub packages kept at latest version' "Could not enumerate ($err)" 'REVIEW' $false `
        'Set a monthly review/update cadence.') )
} elseif ($stale.Count -eq 0) {
    $findings.Add( (New-Finding 'Content Hub Currency' 'content_hub' `
        'MS Sentinel content best practice / NIST SI-3, SI-4 (detection currency)' `
        'Installed Content Hub packages kept at latest version' "$($installed.Count) installed, all current" 'PASS' $false `
        'Maintain monthly review cadence.') )
} else {
    $status = if ($stale.Count -gt 5) { 'GAP' } else { 'REVIEW' }
    $findings.Add( (New-Finding 'Content Hub Currency' 'content_hub' `
        'MS Sentinel content best practice / NIST SI-3, SI-4 (detection currency)' `
        'Installed Content Hub packages kept at latest version' `
        "$($installed.Count) installed; $($stale.Count) with updates available" $status ($stale.Count -gt 10) `
        "Report framed updates as 'normal.' A large backlog of stale detection content means missed Microsoft-authored detections - set a monthly review/update cadence and test high-impact content in non-prod first.") )
}

# 8. Subscription Forwarding
$fwd = @($doc.subscriptionForwarding)
if ($fwd.Count -eq 0) {
    $findings.Add( (New-Finding 'Subscription Log Forwarding' 'forwarding' `
        'CIS Azure 5.1.x / NIST AU-2, AU-12' `
        'Every subscription forwards Activity Log to the central workspace' `
        'Forwarding not checked (collector run with -SkipForwardingCheck?)' 'REVIEW' $false `
        'Re-run collector without -SkipForwardingCheck to verify central ingestion.') )
} else {
    $total = $fwd.Count
    $ok = @($fwd | Where-Object { $_.activityLogToCentral }).Count
    $missing = @($fwd | Where-Object { -not $_.activityLogToCentral } | ForEach-Object { if($_.subscriptionName){$_.subscriptionName}else{$_.subscriptionId} })
    if ($ok -eq $total) { $status='PASS'; $under=$false }
    elseif ($ok -eq 0) { $status='CRITICAL'; $under=$true }
    else { $status='GAP'; $under=$true }
    $observed = "$ok of $total subscription(s) forward Activity Log to central workspace"
    if ($missing.Count -gt 0) {
        $show = $missing | Select-Object -First 15
        $observed += ". Missing: $($show -join ', ')"
        if ($missing.Count -gt 15) { $observed += " (+$($missing.Count - 15) more)" }
    }
    $findings.Add( (New-Finding 'Subscription Log Forwarding' 'forwarding' `
        'CIS Azure 5.1.x / NIST AU-2, AU-12' `
        'Every subscription forwards Activity Log to the central workspace' $observed $status $under `
        'The original report does NOT assess central log forwarding at all. Any subscription not forwarding is a blind spot the report never surfaced - enforce via Azure Policy (DeployIfNotExists) on diagnostic settings.') )
}

# ---------------------------------------------------------------------------
# Render HTML
# ---------------------------------------------------------------------------
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
function E($x) { [System.Web.HttpUtility]::HtmlEncode([string]$x) }

$statusColor = @{ CRITICAL='#7d1414'; GAP='#c62828'; REVIEW='#ef6c00'; PASS='#2e7d32'; INFO='#1565c0' }
$sorted = $findings | Sort-Object @{ Expression = { $SeverityOrder[$_.Status] } }

# Summary pills
$counts = $sorted | Group-Object Status
$pills = ($counts | Sort-Object @{Expression={$SeverityOrder[$_.Name]}} | ForEach-Object {
    $c = $statusColor[$_.Name]; if (-not $c) { $c='#555' }
    "<div class='pill' style='border-left:5px solid $c'><b>$(E $_.Name)</b><span>$($_.Count)</span></div>"
}) -join "`n"

# Understated callout
$understated = @($sorted | Where-Object { $_.Understated })
$understatedBlock = ''
if ($understated.Count -gt 0) {
    $items = ($understated | ForEach-Object {
        "<li><b>$(E $_.Area)</b> &mdash; report tone: <i>$(E $_.ReportTone)</i>; observed: $(E $_.Observed)</li>"
    }) -join "`n"
    $understatedBlock = @"
<div class="callout">
  <h2>&#9888; Where the original report understated a gap</h2>
  <p>The VA report framed the items below as routine "optimization opportunities" or stated capabilities the live configuration does not fully support. These warrant attention beyond the report's tone:</p>
  <ul>$items</ul>
</div>
"@
}

# Rows
$rows = ($sorted | ForEach-Object {
    $c = $statusColor[$_.Status]; if (-not $c) { $c='#555' }
    $flag = if ($_.Understated) { "<span class='flag'>&#9888; Report understated</span>" } else { '' }
    @"
<tr>
  <td><b>$(E $_.Area)</b><div class="sub">$(E $_.ReportSection)</div></td>
  <td><span class="badge" style="background:$c">$(E $_.Status)</span>$flag</td>
  <td>$(E $_.Baseline)</td>
  <td>$(E $_.Expected)</td>
  <td>$(E $_.Observed)</td>
  <td class="claim">$(E $_.ReportSaid)</td>
  <td>$(E $_.Recommendation)</td>
</tr>
"@
}) -join "`n"

$meta = $doc.metadata
$subCount = @($meta.subscriptionsChecked).Count
$generated = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + ' UTC'

$htmlOut = @"
<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Sentinel Config vs. VA Report - Delta Analysis</title>
<style>
 body{font-family:Segoe UI,Arial,sans-serif;margin:0;color:#1a1a1a;background:#f4f6f8;}
 .banner{background:#fff3cd;border:1px solid #ffe69c;color:#664d03;padding:10px 24px;font-size:12px;font-weight:600;}
 .hdr{background:#0b3d6b;color:#fff;padding:24px;}
 .hdr h1{margin:0;font-size:22px;} .hdr p{margin:4px 0 0;font-size:13px;opacity:.85;}
 .wrap{max-width:1320px;margin:20px auto;padding:0 24px;}
 .meta{background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.08);font-size:13px;}
 .meta b{display:inline-block;width:180px;color:#0b3d6b;}
 .sum{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;}
 .pill{background:#fff;border-radius:8px;padding:10px 16px;box-shadow:0 1px 3px rgba(0,0,0,.08);font-size:13px;display:flex;flex-direction:column;}
 .pill b{font-size:11px;color:#555;text-transform:uppercase;letter-spacing:.5px;}
 .pill span{font-size:22px;font-weight:700;}
 .callout{background:#fff4f4;border:1px solid #f3c2c2;border-radius:8px;padding:16px 20px;margin-bottom:20px;}
 .callout h2{margin:0 0 8px;font-size:16px;color:#7d1414;}
 .callout ul{margin:8px 0 0;padding-left:20px;font-size:13px;line-height:1.6;}
 table{width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08);}
 th{background:#0b3d6b;color:#fff;text-align:left;padding:10px 12px;font-size:11px;text-transform:uppercase;letter-spacing:.5px;}
 td{padding:10px 12px;border-top:1px solid #eee;font-size:12.5px;vertical-align:top;}
 tr:nth-child(even) td{background:#fafbfc;}
 .badge{color:#fff;padding:2px 10px;border-radius:12px;font-weight:600;font-size:11px;display:inline-block;}
 .sub{color:#888;font-size:11px;margin-top:3px;}
 .claim{color:#555;font-style:italic;max-width:260px;}
 .flag{display:block;color:#c62828;font-size:10px;font-weight:600;margin-top:4px;}
 .foot{text-align:center;color:#888;font-size:11px;margin:24px 0;}
</style></head><body>
<div class="banner">FOR OFFICIAL USE ONLY (FOUO) &mdash; PRE-DECISIONAL. Handle per DHS/VA policy. Not for public release.</div>
<div class="hdr">
  <h1>Microsoft Sentinel - Live Configuration vs. VA CSOC Report Delta</h1>
  <p>Blended baseline: CIS Azure Foundations &bull; Microsoft Sentinel best practices &bull; NIST SP 800-53 Rev 5</p>
</div>
<div class="wrap">
  <div class="meta">
    <div><b>Central Workspace:</b> $(E $meta.centralWorkspace)</div>
    <div><b>Cloud:</b> $(E $meta.cloud)</div>
    <div><b>Subscriptions Checked:</b> $subCount</div>
    <div><b>Data Collected:</b> $(E $meta.collectedUtc)</div>
    <div><b>Analysis Generated:</b> $generated</div>
    <div><b>Compared Against:</b> VA CSOC Toolset Configuration Report, Dec 22, 2025</div>
  </div>
  <div class="sum">$pills</div>
  $understatedBlock
  <table>
    <thead><tr>
      <th>Area / Report Section</th><th>Status</th><th>Baseline (CIS / MS / NIST)</th>
      <th>Expected</th><th>Observed (Live)</th><th>What the Report Said</th><th>Recommendation</th>
    </tr></thead>
    <tbody>$rows</tbody>
  </table>
  <div class="foot">FOR OFFICIAL USE ONLY &bull; PRE-DECISIONAL &bull; Delta analysis generated from collector JSON. Baselines are closest-fit control families, not literal one-to-one citations; validate against the authoritative benchmark version in force.</div>
</div></body></html>
"@

$htmlOut | Out-File -FilePath $OutputPath -Encoding UTF8

# Console summary
Write-Host "`nAnalyzed: $InputPath"
Write-Host "Workspace: $($ws.workspaceName)"
Write-Host ("-" * 60)
foreach ($f in $sorted) {
    $flag = if ($f.Understated) { '  <-- report understated' } else { '' }
    Write-Host ("  [{0,-8}] {1}{2}" -f $f.Status, $f.Area, $flag)
}
Write-Host ("-" * 60)
Write-Host "HTML report written to: $OutputPath" -ForegroundColor Green