<#
.SYNOPSIS
    Copies delivered documents from source folders into a year-organized archive.

.DESCRIPTION
    Scans one or more source directories recursively for files matching the naming
    convention "<SystemName>_<YYYYMMDD>.<ext>" (doc, docx, or pdf). Matching files
    are copied to a destination root folder, organized into subfolders by delivery
    year. Duplicate filenames are renamed with a sequential "_dup<N>" suffix to
    prevent overwrites. A CSV index of all copied files is written upon completion.

.NOTES
    Name:       Copy-DeliveredDocuments.ps1
    Author:     <Author>
    Created:    <Date>
    Version:    1.0

    File Naming Convention:
        <SystemName>_<YYYYMMDD>.<ext>
        Example: CloudToolset_20240315.docx

    Destination Structure:
        C:\Delivered\
            2023\
                SystemA_20230101.pdf
            2024\
                SystemB_20240315.docx
                SystemB_20240315_dup1.docx   # duplicate renamed automatically

    Output:
        C:\Temp\Delivered_Index.csv — tab of all copied files with metadata.

    Behavior:
        - Non-matching filenames are silently skipped.
        - Files with unparseable dates are silently skipped.
        - Source files are COPIED, not moved (safe/non-destructive).
        - Year subfolders are created automatically if they do not exist.
#>


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Source folders to scan. Add or remove paths as needed.
$folders = @(
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents - Cloud Surge (5.20)\4.4-5.8​ CTS\CTS Assessments\System Reports\2025 September System Reports",
    "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\_Shared Documents

)

# Root destination folder. Year subfolders (e.g. "2024") are created beneath this.
$destinationRoot = "C:\Delivered"
"C:\Delivered"

# Regex pattern to validate and parse the expected filename format.
# Captures: SystemName (anything before the last underscore), Date (8 digits), Extension.
# $pattern = '^(?<SystemName>.+)_(?<Date>\d{8})\.(?<Extension>doc|docx|pdf)$'

# Copy just the PDF File
$pattern = '^(?<SystemName>.+)_(?<Date>\d{8})\.(?<Extension>doc|pdf)$'

# ---------------------------------------------------------------------------
# Main Processing
# ---------------------------------------------------------------------------

$results = foreach ($folder in $folders) {

    # Skip any source folder that does not exist rather than throwing an error.
    if (Test-Path $folder) {

        Get-ChildItem -Path $folder -Recurse -File | ForEach-Object {

            # Only process files that match the expected naming convention.
            if ($_.Name -match $pattern) {

                $rawDate    = $matches['Date']
                $systemName = $matches['SystemName']

                # Parse the 8-digit date string. Skip the file if the date is invalid.
                try {
                    $parsedDate = [datetime]::ParseExact($rawDate, 'yyyyMMdd', $null)
                    $year       = $parsedDate.Year
                }
                catch {
                    return  # Skip files with unparseable dates
                }

                # Build the year-based destination subfolder path.
                $targetFolder = Join-Path $destinationRoot $year

                # Create the year subfolder if it does not already exist.
                if (!(Test-Path $targetFolder)) {
                    New-Item -ItemType Directory -Path $targetFolder | Out-Null
                }

                $destinationPath = Join-Path $targetFolder $_.Name

                # Handle duplicate filenames by appending "_dup<N>" before the extension.
                # Increments N until a unique filename is found.
                if (Test-Path $destinationPath) {
                    $base    = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
                    $ext     = $_.Extension
                    $counter = 1

                    do {
                        $newName         = "{0}_dup{1}{2}" -f $base, $counter, $ext
                        $destinationPath = Join-Path $targetFolder $newName
                        $counter++
                    } while (Test-Path $destinationPath)
                }

                # Copy the file. Source is preserved (non-destructive operation).
                Copy-Item -Path $_.FullName -Destination $destinationPath

                # Emit a result object for each successfully copied file.
                # These are collected into $results for CSV export.
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

# Export the index of all copied files to CSV for audit/tracking purposes.
$results | Export-Csv "C:\Temp\Delivered_Index.csv" -NoTypeInformation