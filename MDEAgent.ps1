Get-AzSubscription

# Ensure you are logged in to your Azure account
#Connect-AzAccount

# Set the subscription ID
$subscriptionId = "1e5b8da5-fbcf-4f83-b932-ed8af1530065"

# Select the subscription
Select-AzSubscription -SubscriptionId $subscriptionId

# Check if the subscription is enrolled in Defender for Cloud
$defenderStatus = Get-AzSecurityContact
if ($defenderStatus) {
    Write-Host "The subscription is enrolled in Defender for Cloud."
} else {
    Write-Host "The subscription is NOT enrolled in Defender for Cloud."
}

# Get all the VMs in the subscription
$vms = Get-AzVM

foreach ($vm in $vms) {
    Write-Host "VM Name: $($vm.Name) - Resource Group: $($vm.ResourceGroupName)"

    # Check if the MDE agent is installed on the VM
    $vmStatus = Get-AzVM -ResourceGroupName $vm.ResourceGroupName -Name $vm.Name -Status
    $extensions = $vmStatus.Extensions

    $mdeAgentInstalled = $false
    foreach ($extension in $extensions) {
        if ($extension.Name -eq "MicrosoftMonitoringAgent" -or $extension.Publisher -eq "Microsoft.Azure.Security.Monitoring") {
            $mdeAgentInstalled = $true
        }
    }

    if ($mdeAgentInstalled) {
        Write-Host "MDE Agent is installed on this VM."
    } else {
        Write-Host "MDE Agent is NOT installed on this VM."
    }
}