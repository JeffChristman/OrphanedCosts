# Define source folders and destination
$sourceFolders = @(
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\2025 September System Reports",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\CY26 System Reports"
)
$destination = "C:\Delivered\Excel"

# Create destination folder if it doesn't exist
if (-not (Test-Path -Path $destination)) {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Write-Host "Created destination folder: $destination" -ForegroundColor Green
}

$totalCopied = 0
$totalSkipped = 0

foreach ($folder in $sourceFolders) {
    if (-not (Test-Path -Path $folder)) {
        Write-Warning "Source folder not found, skipping: $folder"
        continue
    }

    Write-Host "`nSearching in: $folder" -ForegroundColor Cyan

    # Get all Excel files recursively
    $excelFiles = Get-ChildItem -Path $folder -Recurse -Include "*.xlsx","*.xls","*.xlsm","*.xlsb"

    foreach ($file in $excelFiles) {
        $destFile = Join-Path $destination $file.Name

        # Handle duplicate filenames by appending a counter
        if (Test-Path -Path $destFile) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $ext = $file.Extension
            $counter = 1
            do {
                $destFile = Join-Path $destination "$baseName`_$counter$ext"
                $counter++
            } while (Test-Path -Path $destFile)

            Write-Host "  Duplicate renamed: $($file.Name) -> $(Split-Path $destFile -Leaf)" -ForegroundColor Yellow
            $totalSkipped++
        }

        Copy-Item -Path $file.FullName -Destination $destFile
        Write-Host "  Copied: $($file.FullName)" -ForegroundColor Gray
        $totalCopied++
    }
}

Write-Host "`n--- Summary ---" -ForegroundColor Green
Write-Host "Total files copied : $totalCopied"
Write-Host "Duplicates renamed : $totalSkipped"
Write-Host "Destination        : $destination"