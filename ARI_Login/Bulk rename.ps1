# ============================================================
#  Rename-SystemFiles.ps1
#  Copies files in CY26 System Reports to: ACRONYM_YYYY-MM-DD.ext
# ============================================================

# --- CONFIGURATION ---
$SourceFolder = "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\CY26 System Reports"# <-- Change this to your actual folder path
$DestFolder   = "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\CY26 System Reports\Renamed"  # <-- Where renamed copies will go
$WhatIf       = $false   # Set to $true for a dry run (no files copied, just previewed)

# --- ACRONYM EXTRACTION RULES ---
# Maps known system name patterns (regex) to their acronym.
# Add more rows here if you have other systems.
$AcronymMap = [ordered]@{
    'Memorial Benefits Management System \(MBMS\)' = 'MBMS'
    'Controls[_ ]by[_ ]state'                       = 'CBS'   # example: rename as needed
    # 'Some Other System \(SOS\)'                   = 'SOS'   # add more here
}

# --- SUPPORTED FILE TYPES ---
$Extensions = @('*.csv', '*.xlsx', '*.xls', '*.docx', '*.pdf')

# ============================================================
# Script logic below — no changes needed below this line
# ============================================================

if (-not (Test-Path $DestFolder)) {
    New-Item -ItemType Directory -Path $DestFolder | Out-Null
    Write-Host "Created destination folder: $DestFolder" -ForegroundColor Cyan
}

$files = Get-ChildItem -Path $SourceFolder -Recurse -Include $Extensions |
         Where-Object { $_.DirectoryName -ne $DestFolder }  # skip already-renamed copies

$results = @()

foreach ($file in $files) {
    $baseName  = $file.BaseName
    $ext       = $file.Extension

    # --- 1. Extract acronym ---
    $acronym = $null
    foreach ($pattern in $AcronymMap.Keys) {
        if ($baseName -match $pattern) {
            $acronym = $AcronymMap[$pattern]
            break
        }
    }

    # If no acronym matched, try to extract anything inside parentheses e.g. (MBMS)
    if (-not $acronym -and $baseName -match '\(([A-Z]{2,10})\)') {
        $acronym = $Matches[1]
    }

    # Final fallback: use first word of filename
    if (-not $acronym) {
        $acronym = ($baseName -split '[_ ]')[0].ToUpper()
    }

    # --- 2. Extract date (supports YYYYMMDD or YYYY-MM-DD anywhere in name) ---
    $dateStr = $null
    if ($baseName -match '(\d{4})[_\-]?(\d{2})[_\-]?(\d{2})') {
        $y = $Matches[1]; $m = $Matches[2]; $d = $Matches[3]
        # Basic sanity check
        if ([int]$m -ge 1 -and [int]$m -le 12 -and [int]$d -ge 1 -and [int]$d -le 31) {
            $dateStr = "$y-$m-$d"
        }
    }

    # --- 3. Build new name ---
    if ($dateStr) {
        $newName = "${acronym}_${dateStr}${ext}"
    } else {
        $newName = "${acronym}_NODATE${ext}"
        Write-Warning "No date found in: $($file.Name)"
    }

    $destPath = Join-Path $DestFolder $newName

    # Handle duplicates by appending a counter
    $counter = 1
    while (Test-Path $destPath) {
        $newName  = "${acronym}_${dateStr}_${counter}${ext}"
        $destPath = Join-Path $DestFolder $newName
        $counter++
    }

    # --- 4. Copy (or preview) ---
    $results += [PSCustomObject]@{
        Original = $file.Name
        NewName  = $newName
        Status   = if ($WhatIf) { "DRY RUN" } else { "Copied" }
    }

    if (-not $WhatIf) {
        Copy-Item -Path $file.FullName -Destination $destPath
    }
}

# --- Summary table ---
Write-Host ""
Write-Host "===== RESULTS =====" -ForegroundColor Cyan
$results | Format-Table -AutoSize

Write-Host ""
if ($WhatIf) {
    Write-Host "DRY RUN complete — no files were copied. Set `$WhatIf = `$false to apply." -ForegroundColor Yellow
} else {
    Write-Host "$($results.Count) file(s) copied to: $DestFolder" -ForegroundColor Green
}