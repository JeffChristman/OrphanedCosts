# Login to Azure if not already logged in
#Connect-AzAccount

# Optional: Select subscription if multiple
# Select-AzSubscription -SubscriptionName "Your Subscription Name"

# Output file
$outputCsv = "PublicIP_Associations.csv"

# Create a list to hold the output
$results = @()

# Get all Public IPs in the subscription
$publicIPs = Get-AzPublicIpAddress -ErrorAction SilentlyContinue

foreach ($ip in $publicIPs) {
    $resourceGroup = $ip.ResourceGroupName
    $ipName = $ip.Name
    $ipAddress = $ip.IpAddress
    $ipAllocation = $ip.PublicIpAllocationMethod
    $ipSku = $ip.Sku.Name
    $ipFqdn = $ip.DnsSettings.Fqdn
    $associatedResource = "None"

    # Check if it's linked to a NIC (used in VM)
    $nic = Get-AzNetworkInterface -ResourceGroupName $resourceGroup | Where-Object {
        $_.IpConfigurations.PublicIpAddress.Id -eq $ip.Id
    }
    if ($nic) {
        $vm = Get-AzVM -ResourceGroupName $resourceGroup | Where-Object {
            $_.NetworkProfile.NetworkInterfaces.Id -eq $nic.Id
        }
        $associatedResource = if ($vm) { "VM: $($vm.Name)" } else { "NIC: $($nic.Name)" }
    }

    # Check Load Balancer
    if ($associatedResource -eq "None") {
        $lbs = Get-AzLoadBalancer -ResourceGroupName $resourceGroup
        foreach ($lb in $lbs) {
            foreach ($frontend in $lb.FrontendIpConfigurations) {
                if ($frontend.PublicIpAddress.Id -eq $ip.Id) {
                    $associatedResource = "LoadBalancer: $($lb.Name)"
                }
            }
        }
    }

    # Check Application Gateway
    if ($associatedResource -eq "None") {
        $gateways = Get-AzApplicationGateway -ResourceGroupName $resourceGroup
        foreach ($gw in $gateways) {
            foreach ($frontend in $gw.FrontendIpConfigurations) {
                if ($frontend.PublicIpAddress.Id -eq $ip.Id) {
                    $associatedResource = "AppGateway: $($gw.Name)"
                }
            }
        }
    }

    # Build result object
    $results += [PSCustomObject]@{
        ResourceGroup       = $resourceGroup
        PublicIPName        = $ipName
        IPAddress           = $ipAddress
        AllocationMethod    = $ipAllocation
        Sku                 = $ipSku
        FQDN                = $ipFqdn
        AssociatedResource  = $associatedResource
    }
}

# Output to console as table
$results | Format-Table -AutoSize

# Export to CSV
$results | Export-Csv -Path $outputCsv -NoTypeInformation

Write-Host "`n✅ Export complete: $outputCsv"
