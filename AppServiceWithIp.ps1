# Requires Az Module
# Connect to Azure
#Connect-AzAccount

# Optional: Set the subscription if you have more than one
Set-AzContext -Subscription "EIS-GOV-INTERNAL"

# Get all App Services across all Resource Groups
$appServices = Get-AzWebApp

# Output array
$results = @()

foreach ($app in $appServices) {
    # Attempt to retrieve the public virtual IP
    $ip = $null

    # Get publishing profile (contains hostname)
    $hostnames = (Get-AzWebApp -ResourceGroupName $app.ResourceGroup -Name $app.Name).DefaultHostName

    # Use DNS to resolve IP if needed
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($hostnames) | 
              Where-Object { $_.AddressFamily -eq 'InterNetwork' } | 
              Select-Object -First 1
    } catch {
        $ip = "Resolution failed"
    }

    $results += [PSCustomObject]@{
        AppName        = $app.Name
        ResourceGroup  = $app.ResourceGroup
        Location       = $app.Location
        DefaultHost    = $hostnames
        PublicIP       = $ip.IPAddressToString
    }
}

# Display in table
$results | Format-Table -AutoSize

# Export to CSV
$results | Export-Csv -Path "./AppServices_PublicIPs.csv" -NoTypeInformation
