<#
.SYNOPSIS
This script flattens nested JSON objects and exports the flattened data to a CSV file.

.DESCRIPTION
This script prompts the user to enter the file paths for an input JSON file and an output CSV file.
It then reads the JSON content, flattens any nested objects, and exports the result to a CSV file.

.PARAMETER None
This script has no specific parameters; it interacts with the user through prompts.

.EXAMPLE
Run the script and follow the prompts to provide the required file paths:
    .\FlattenJsonToCsv.ps1
#>

# Function to flatten nested JSON
Function Flatten-Object {
    <#
    .SYNOPSIS
    Flattens a nested JSON object.

    .DESCRIPTION
    This function flattens a nested JSON object by recursively processing each property.
    Nested objects and arrays are expanded and included in the resulting flat structure.

    .PARAMETER InputObject
    The JSON object to flatten.

    .PARAMETER Prefix
    An optional prefix to prepend to the property names in the flattened object.
    
    .INPUTS
    PSObject. The function expects a PSObject as input.

    .OUTPUTS
    PSObject. The function outputs a flattened PSObject.

    .EXAMPLE
    $flattened = Flatten-Object -InputObject $jsonObject
    #>
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
                # If the property is a nested object, recursively flatten it
                $nestedProperties = Flatten-Object -InputObject $property.Value -Prefix $name
                foreach ($nestedProperty in $nestedProperties.PSObject.Properties) {
                    $properties[$nestedProperty.Name] = $nestedProperty.Value
                }
            } elseif ($property.Value -is [System.Collections.IEnumerable] -and
                !($property.Value -is [string]) -and
                !($property.Value -is [Hashtable])) {
                # If the property is an array, process each element
                $counter = 0
                foreach ($item in $property.Value) {
                    $nestedProperties = Flatten-Object -InputObject $item -Prefix "$name.$counter"
                    foreach ($nestedProperty in $nestedProperties.PSObject.Properties) {
                        $properties["$($nestedProperty.Name)"] = $nestedProperty.Value
                    }
                    $counter++
                }
            } else {
                # If the property is a simple value, add it directly
                $properties[$name] = $property.Value
            }
        }

        [PSCustomObject]$properties
    }
}

# Prompt user for the input JSON file path
$jsonFilePath = Read-Host -Prompt "Enter the path to the input JSON file"

# Prompt user for the output CSV file path
$csvFilePath = Read-Host -Prompt "Enter the path to the output CSV file"

# Read the JSON content from the file
try {
    $jsonContent = Get-Content -Path $jsonFilePath -Raw | ConvertFrom-Json
} catch {
    Write-Error "Failed to read JSON file. Please check the file path and format."
    exit
}

# Check if the JSON content includes the necessary structure
if (-not $jsonContent.data -or -not $jsonContent.data.policyValues -or -not $jsonContent.data.policyValues.items) {
    Write-Error "The JSON data doesn't have the expected structure."
    exit
}

# Flatten the JSON content
$flattenedJson = $jsonContent.data.policyValues.items | ForEach-Object { Flatten-Object -InputObject $_ }

# Export the flattened JSON content to a CSV file
try {
    $flattenedJson | Export-Csv -Path $csvFilePath -NoTypeInformation
    Write-Output "JSON data has been successfully converted to CSV and saved to $csvFilePath"
} catch {
    Write-Error "Failed to export CSV file. Please check the file path and permissions."
}