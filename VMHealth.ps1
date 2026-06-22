# Get all subscriptions in the tenant
$subscriptions = Get-AzSubscription

# Variable to store the results
$results = @()

foreach ($subscription in $subscriptions) {
    # Select the subscription
    Select-AzSubscription -SubscriptionId $subscription.Id

    # Check if the subscription is enrolled in Defender for Cloud
    $pricingDetails = Get-AzSecurityPricing -Name VirtualMachines
    $subscriptionStatus = if ($pricingDetails.PricingTier -ne 'Free') { "Enrolled" } else { "Not Enrolled" }

    # Get all the VMs in the subscription
    $vms = Get-AzVM

    foreach ($vm in $vms) {
        # Check if the MDE agent is installed on the VM
        $vmStatus = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status
        $extensions = $vmStatus.Extensions

        $mdeAgentInstalled = $false
        foreach ($extension in $extensions) {
            if ($extension.Name -eq "MicrosoftMonitoringAgent" -or $extension.Publisher -eq "Microsoft.Azure.Security.Monitoring") {
                $mdeAgentInstalled = $true
            }
        }

        $mdeAgentStatus = if ($mdeAgentInstalled) { "Installed" } else { "Not Installed" }

        # Get the health status of the VM
        $instanceView = Get-AzVMInstanceView -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name
        $vmHealthState = $instanceView.Statuses | Where-Object { $_.Code -like 'PowerState/*' } | Select-Object -ExpandProperty DisplayStatus
        if ($vmHealthState -eq 'VM running') {
            $vmHealth = "Healthy"
        }
        else {
            $vmHealth = "Unhealthy"
        }

        # Add the results to the array
        $results += [PSCustomObject]@{
            SubscriptionId     = $subscription.Id
            SubscriptionName   = $subscription.Name
            SubscriptionStatus = $subscriptionStatus
            VMName             = $vm.Name
            ResourceGroupName  = $vm.ResourceGroupName
            MDEAgentStatus     = $mdeAgentStatus
            VMHealth           = $vmHealth
        }
    }
}

# Output the results as a table on the screen
$results | Format-Table -AutoSize

# Export the results to a CSV file
$results | Export-Csv -Path "VM_Status_Report.csv" -NoTypeInformation