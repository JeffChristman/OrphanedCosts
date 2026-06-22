
# This script performs an assessment of Azure Network Security Groups (NSGs), Route Tables (User-Defined Routes), and Public IP addresses,
# and then exports the results to a CSV file.

# Initialize an empty array to store the results
$results = @()

# Get all NSGs (Network Security Groups)
$nsgs = Get-AzNetworkSecurityGroup
foreach ($nsg in $nsgs) {
    # For each NSG, iterate through all its security rules
    foreach ($rule in $nsg.SecurityRules) {
        # Add a custom object representing each security rule to the results array
        $results += [PSCustomObject]@{
            Subscription       = $sub.Name         # Name of the Azure subscription
            ResourceType       = "NSG Rule"        # Resource type is NSG Rule
            NSGName            = $nsg.Name         # Name of the NSG
            RuleName           = $rule.Name        # Name of the security rule
            Direction          = $rule.Direction   # Direction of traffic (Inbound/Outbound)
            Access             = $rule.Access      # Access type (Allow/Deny)
            Priority           = $rule.Priority    # Priority of the rule
            Source             = $rule.SourceAddressPrefix       # Source address prefix
            Destination        = $rule.DestinationAddressPrefix  # Destination address prefix
            Port               = $rule.DestinationPortRange      # Destination port range
            Protocol           = $rule.Protocol    # Protocol (e.g., TCP, UDP)
        }
    }
}

# Get all Route Tables (User-Defined Routes)
$routes = Get-AzRouteTable
foreach ($routeTable in $routes) {
    # For each route table, iterate through all its routes
    foreach ($route in $routeTable.Routes) {
        # Add a custom object representing each route to the results array
        $results += [PSCustomObject]@{
            Subscription       = $sub.Name             # Name of the Azure subscription
            ResourceType       = "Route"               # Resource type is Route
            Name               = $routeTable.Name      # Name of the route table
            AddressPrefix      = $route.AddressPrefix  # Address prefix for the route
            NextHopType        = $route.NextHopType    # Next hop type (e.g., VirtualAppliance, Internet)
            NextHopIpAddress   = $route.NextHopIpAddress  # Next hop IP address
        }
    }
}

# Get all Public IP addresses
$publicIPs = Get-AzPublicIpAddress
foreach ($ip in $publicIPs) {
    # Add a custom object representing each public IP address to the results array
    $results += [PSCustomObject]@{
        Subscription       = $sub.Name                  # Name of the Azure subscription
        ResourceType       = "Public IP"                # Resource type is Public IP
        Name               = $ip.Name                   # Name of the public IP resource
        IPAddress          = $ip.IpAddress              # The public IP address
        AssignedTo         = $ip.IpConfiguration.Id     # ID of the IP configuration assigned to this public IP
        AllocationMethod   = $ip.PublicIpAllocationMethod  # Allocation method (e.g., Static, Dynamic)
        SKU                = $ip.Sku.Name               # SKU (Stock Keeping Unit) name of the public IP
    }
}

# Export the collected results to a CSV file
$timestamp = Get-Date -Format "yyyyMMdd-HHmm"  # Generate a timestamp for the filename
$csvPath = "$env:USERPROFILE\Desktop\Azure_NetworkSecurity_Assessment_$timestamp.csv"  # Set the output path for the CSV file
$results | Export-Csv -Path $csvPath -NoTypeInformation  # Export the results to the CSV without type information
Write-Host "`nExport complete: $csvPath" -ForegroundColor Green  # Inform the user that the export is complete
```

This documentation explains each step and the purpose of each part of the script. If you need further customization or have additional questions, feel free to ask!