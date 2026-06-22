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

# Path to the input JSON file
$jsonFilePath = "C:\Users\OITNSMCHRISJ\OneDrive - Department of Veterans Affairs\01_Projects\Azure\Scripts\example.json"

# Path to the output CSV file
$csvFilePath = "AWS\VAEC2\JSON\example.csv"

# Read the JSON content from the file
$jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json

# Flatten the JSON content
$flattenedJson = $jsonContent.data.policyValues.items | ForEach-Object { Flatten-Object -InputObject $_ }

# Export the flattened JSON content to a CSV file
$flattenedJson | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Output "JSON data has been successfully converted to CSV and saved to $csvFilePath"