# Function to flatten nested JSON
Function Flatten-Object {
    Param (
        [Parameter(ValueFromPipeline = $true)]
        [PSObject] $InputObject,
        [string] $Prefix = ""
    )

    process {
        $properties = @{}
        foreach ($property in $InputObject.PSObject.Properties) {
            $name = if ($Prefix) { "$Prefix.$($property.Name)" } else { $property.Name }

            if ($property.Value -is [PSCustomObject]) {
                $nestedProperties = Flatten-Object -InputObject $property.Value -Prefix $name
                foreach ($nestedProperty in $nestedProperties.PSObject.Properties) {
                    $properties[$nestedProperty.Name] = $nestedProperty.Value
                }
            } elseif ($property.Value -is [System.Collections.IEnumerable] -and
                !($property.Value -is [string]) -and
                !($property.Value -is [IDictionary])) {
                $counter = 0
                foreach ($item in $property.Value) {
                    $nestedProperties = Flatten-Object -InputObject $item -Prefix "$name.$counter"
                    foreach ($nestedProperty in $nestedProperties.PSObject.Properties) {
                        $properties["$($nestedProperty.Name)"] = $nestedProperty.Value
                    }
                    $counter++
                }
            } else {
                $properties[$name] = $property.Value
            }
        }

        [PSCustomObject]$properties
    }
}

# Path to the input folder containing JSON files
$jsonFolderPath = "C:\Azure\Scripts\AWS\VAEC2\JSON\Logging"

# Retrieve all JSON files from the specified folder
$jsonFiles = Get-ChildItem -Path $jsonFolderPath -Filter *.json

foreach ($jsonFile in $jsonFiles) {

    # Path to the output CSV file with the same structure as the original JSON file
    $csvFilePath = [System.IO.Path]::ChangeExtension($jsonFile.FullName, ".csv")

    # Read the JSON content from the file
    $jsonContent = Get-Content -Path $jsonFile.FullName -Raw | ConvertFrom-Json

    # Flatten the JSON content
    $flattenedJson = $jsonContent.data.policyValues.items | ForEach-Object { Flatten-Object -InputObject $_ }

    # Export the flattened JSON content to a CSV file
    $flattenedJson | Export-Csv -Path $csvFilePath -NoTypeInformation

    Write-Output "JSON data from $($jsonFile.FullName) has been successfully converted to CSV and saved to $csvFilePath"
}
This script processes each JSON file in the given fol