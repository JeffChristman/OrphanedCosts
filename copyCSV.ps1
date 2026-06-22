# Define source and destination paths
$sourcePath = "C:\Azure\Scripts\VAEC1\5.0"    # Modify this to the directory where your .csv files are located
$destinationPath = "C:\Azure\Scripts\Data"  # Modify this to the directory where you want to copy the .csv files

# Check if source and destination paths exist
if (-Not (Test-Path -Path $sourcePath)) {
    Write-Host "Source path does not exist: $sourcePath"
    exit
}
if (-Not (Test-Path -Path $destinationPath)) {
    Write-Host "Destination path does not exist: $destinationPath"
    exit
}

# Get all .csv files from the source path
$csvFiles = Get-ChildItem -Path $sourcePath -Filter *.csv

# Copy each .csv file to the destination path
foreach ($file in $csvFiles) {
    $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.Name
    Copy-Item -Path $file.FullName -Destination $destinationFile -Force
    Write-Host "Copied $($file.Name) to $destinationPath"
}

Write-Host "All .csv files have been copied."