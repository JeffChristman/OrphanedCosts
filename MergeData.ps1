# Define the paths to the CSV files
$csvFile1 = "C:\Azure\Scripts\VAEC1\2.0\20101.csv"
$csvFile2 = "C:\Azure\Scripts\VAEC2\2.0\20101.csv"
$outputFile = "C:\Azure\Scripts\VAEC2\2.0\vmergedFile.csv"

# Import the first CSV file
$data1 = Import-Csv -Path $csvFile1

# Import the second CSV file
$data2 = Import-Csv -Path $csvFile2

# Combine the data
$mergedData = $data1 + $data2

# Export the combined data to a new CSV file
$mergedData | Export-Csv -Path $outputFile -NoTypeInformation

Write-Host "CSV files have been successfully merged and saved to $outputFile"