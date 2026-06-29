<#
.SYNOPSIS
    Collects Excel files from one or more source folders into a single destination folder.

.DESCRIPTION
    Recursively searches each source directory for Excel workbooks (.xlsx, .xls, .xlsm,
    .xlsb) and copies them into a flat destination folder. Files that already exist at
    the destination are skipped rather than overwritten. A summary of copied and skipped
    files is printed to the console upon completion.

.NOTES
    Name:       Copy-ExcelFiles.ps1
    Author:     <Author>
    Created:    <Date>
    Version:    1.0

    Supported File Types:
        .xlsx   Excel Workbook
        .xls    Excel 97-2003 Workbook
        .xlsm   Excel Macro-Enabled Workbook
        .xlsb   Excel Binary Workbook

    Destination Structure:
        All matched files are copied into a single flat folder (no subfolders).
        Example: C:\Delivered\Excel\ReportA.xlsx

    Duplicate Handling:
        Files are skipped — not renamed — if a file with the same name already
        exists at the destination. Review skipped files in the console output.

    Behavior:
        - Source files are COPIED, not moved (safe/non-destructive).
        - Missing source folders log a warning and are skipped gracefully.
        - The destination folder is created automatically if it does not exist.

    Known Issue:
        The $sourceFolders array contains a duplicate entry for the CY26 System Reports
        folder. This causes that folder to be scanned twice; files found on the second
        pass will be reported as skipped (already exists). Remove the duplicate path
        to suppress the redundant warnings.
#>


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Source folders to scan. All subfolders are searched recursively.
# NOTE: The third entry below is a duplicate of the second — see .NOTES above.
$sourceFolders = @(
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\2025 September System Reports",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\CY26 System Reports"
    
    
)

# Flat destination folder where all Excel files will be collected.
$destination = "C:\Delivered\Excel"


# ---------------------------------------------------------------------------
# Initialization
# ---------------------------------------------------------------------------

# Create the destination folder if it does not already exist.
if (-not (Test-Path -Path $destination)) {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Write-Host "Created destination folder: $destination" -ForegroundColor Green
}

# Counters for the end-of-run summary.
$totalCopied  = 0
$totalSkipped = 0


# ---------------------------------------------------------------------------
# Main Processing
# ---------------------------------------------------------------------------

foreach ($folder in $sourceFolders) {

    # Warn and skip any source folder that does not exist rather than throwing an error.
    if (-not (Test-Path -Path $folder)) {
        Write-Warning "Source folder not found, skipping: $folder"
        continue
    }

    Write-Host "`nSearching in: $folder" -ForegroundColor Cyan

    # Collect all Excel file types from the current source folder, recursively.
    $excelFiles = Get-ChildItem -Path $folder -Recurse -Include "*.xlsx","*.xls","*.xlsm","*.xlsb"

    foreach ($file in $excelFiles) {

        $destFile = Join-Path $destination $file.Name

        # Skip files that already exist at the destination to prevent overwrites.
        # The source file is left untouched; the operator is notified via console.
        if (Test-Path -Path $destFile) {
            Write-Host "  Skipped (already exists): $($file.Name)" -ForegroundColor Yellow
            $totalSkipped++
            continue
        }

        # Copy the file. Source is preserved (non-destructive operation).
        Copy-Item -Path $file.FullName -Destination $destFile
        Write-Host "  Copied: $($file.FullName)" -ForegroundColor Gray
        $totalCopied++
    }
}


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host "`n--- Summary ---"         -ForegroundColor Green
Write-Host "Total files copied : $totalCopied"
Write-Host "Total files skipped: $totalSkipped"
Write-Host "Destination        : $destination"