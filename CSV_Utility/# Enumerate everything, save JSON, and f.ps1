# Enumerate everything, save JSON, and flatten
$apiVersions = @("2023-01-01-preview","2022-05-01-preview","2022-01-01-preview")
$subs = Get-AzSubscription | Where-Object State -eq 'Enabled'
$ts = (Get-Date).ToString("yyyyMMdd_HHmmss")
$csv = Join-Path $env:TEMP "DfC_ExportSettings_Full_$ts.csv"
$dir = Join-Path $env:TEMP "DfC_ExportSettings_JSON_$ts"; New-Item -ItemType Directory -Path $dir -Force | Out-Null
$rows = @()

foreach ($s in $subs) {
  Select-AzSubscription -SubscriptionId $s.Id | Out-Null
  $resp = $null; $apiUsed = $null
  foreach ($v in $apiVersions) {
    try { $resp = Invoke-AzRestMethod -Method GET -Path "/subscriptions/$($s.Id)/providers/Microsoft.Security/exportSettings?api-version=$v"; $apiUsed=$v; break } catch {}
  }
  if (-not $resp) { continue }
  $items = ($resp.Content | ConvertFrom-Json).value
  foreach ($it in ($items | Where-Object { $_ })) {
    $file = Join-Path $dir ("$($s.Id)_$($it.name).json" -replace '[^a-zA-Z0-9_\.-]','_')
    $it | ConvertTo-Json -Depth 64 | Out-File -Encoding UTF8 $file

    $p = $it.properties
    # normalize lists
    $dt = @(); if ($p.exportDataTypes) { $dt += $p.exportDataTypes } ; if ($p.dataTypes) { $dt += $p.dataTypes }
    $dtLower = $dt | ForEach-Object { $_.ToString().ToLower() }

    $rows += [pscustomobject]@{
      SubscriptionName              = $s.Name
      SubscriptionId                = $s.Id
      ExportSettingName             = $it.name
      ApiVersion                    = $apiUsed
      ExportToLogAnalytics          = [bool]$p.exportToLogAnalytics
      WorkspaceResourceId           = $p.workspaceResourceId
      ExportToEventHub              = [bool]$p.exportToEventHub
      EventHubResourceId            = $p.eventHubResourceId
      DataTypesRaw                  = ($dt -join '; ')
      SecurityAlertsChecked         = ($dtLower -match 'securityalerts')
      SecurityRecommendationsChecked= ($dtLower -match 'securityrecommendations')
      SecureScoreControlsChecked    = ($dtLower -match 'securescorecontrols')
      RegulatoryComplianceChecked   = ($dtLower -match 'regulatory')
      RawPropertiesPreview          = ($p | ConvertTo-Json -Depth 6)
    }
  }
}
$rows | Sort-Object SubscriptionName, ExportSettingName | Export-Csv -NoTypeInformation -Encoding UTF8 $csv
Write-Host "CSV: $csv"
Write-Host "Raw JSON dir: $dir"


