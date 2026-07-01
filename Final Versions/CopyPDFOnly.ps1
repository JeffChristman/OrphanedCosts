<#
.SYNOPSIS
    Copies delivered PDF documents from source folders into a year-organized archive.

.DESCRIPTION
    Scans source directories recursively for PDF files matching the naming
    convention "<SystemName>_<YYYYMMDD>.pdf". Matching files are copied to a
    destination root, organized into subfolders by delivery year. Files that
    already exist at the destination are skipped (not overwritten, not renamed).
    A CSV index of all copied files is written upon completion.
#>


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

$folders = @(
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8 CTS\Cloud Toolset Config. (1000AD)",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8 CTS\CTS Assessments",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\2025 September System Reports",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\CY26 System Reports"
)

$destinationRoot = "C:\Delivered"

# PDF only.
$pattern = '^(?<SystemName>.+)_(?<Date>\d{8})\.(?<Extension>pdf)$'


# ---------------------------------------------------------------------------
# Main Processing
# ---------------------------------------------------------------------------

$results = foreach ($folder in $folders) {

    if (Test-Path $folder) {

        Get-ChildItem -Path $folder -Recurse -File | ForEach-Object {

            if ($_.Name -match $pattern) {

                $rawDate    = $matches['Date']
                $systemName = $matches['SystemName']

                try {
                    $parsedDate = [datetime]::ParseExact($rawDate, 'yyyyMMdd', $null)
                    $year       = $parsedDate.Year
                }
                catch {
                    return  # Skip files with unparseable dates
                }

                $targetFolder = Join-Path $destinationRoot ([string]$year)

                if (!(Test-Path $targetFolder)) {
                    New-Item -ItemType Directory -Path $targetFolder | Out-Null
                }

                $destinationPath = Join-Path $targetFolder $_.Name

                # Ignore duplicates: if the file already exists, skip it.
                if (Test-Path $destinationPath) {
                    Write-Host "Skipped (already exists): $($_.Name)"
                    return
                }

                # Copy, with OneDrive cloud-only hydration fallback.
                try {
                    Copy-Item -Path $_.FullName -Destination $destinationPath -ErrorAction Stop
                }
                catch {
                    try {
                        [System.IO.File]::ReadAllBytes($_.FullName) | Out-Null
                        Copy-Item -Path $_.FullName -Destination $destinationPath -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Could not copy (cloud-only?): $($_.FullName)"
                        return
                    }
                }

                [PSCustomObject]@{
                    SystemName   = $systemName
                    DeliveryDate = $parsedDate
                    Year         = $year
                    SourcePath   = $_.FullName
                    TargetPath   = $destinationPath
                }
            }
        }
    }
}


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

if (!(Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" | Out-Null }

$results | Export-Csv "C:\Temp\Delivered_Index.csv" -NoTypeInformation