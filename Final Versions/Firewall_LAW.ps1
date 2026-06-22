# This script performs an audit of Azure Firewalls across all subscriptions in Azure GovCloud,
# checking for diagnostic settings pointing to a specific Log Analytics Workspace and exporting the results to a table.

# Connect to Azure GovCloud (uncomment the following line if needed)
# Connect-AzAccount -Environment AzureUSGovernment

# Set the target Log Analytics Workspace Name
$targetWorkspaceName = "vaecla-security-gov"

# Initialize an empty array to store the results
$results = @()

# Get all subscriptions in the current Azure context
$subscriptions = Get-AzSubscription

foreach ($sub in $subscriptions) {
    # Set the context to the current subscription
    Set-AzContext -SubscriptionId $sub.Id

    # Get all firewalls in the current subscription
    $firewalls = Get-AzFirewall

    foreach ($fw in $firewalls) {
        # Get diagnostic settings for the current firewall
        $diagSettings = Get-AzDiagnosticSetting -ResourceId $fw.Id -ErrorAction SilentlyContinue

        if ($diagSettings) {
            # Initialize a flag to check if the target workspace is found
            $workspaceFound = $false
            
            # Check if any of the diagnostic settings point to the target workspace
            foreach ($ws in $diagSettings.WorkspaceId) {
                $workspace = Get-AzOperationalInsightsWorkspace -ResourceId $ws -ErrorAction SilentlyContinue
                if ($workspace.Name -eq $targetWorkspaceName) {
                    $workspaceFound = $true
                    break
                }
            }

            # Add a custom object to the results representing the firewall's diagnostic status
            $results += [PSCustomObject]@{
                SubscriptionName = $sub.Name                                      # Name of the subscription
                FirewallName     = $fw.Name                                       # Name of the firewall
                ResourceGroup    = $fw.ResourceGroupName                          # Name of the resource group containing the firewall
                Region           = $fw.Location                                   # Location/Region of the firewall
                DiagnosticStatus = if ($workspaceFound) {                         # Diagnostic status
                                     "✅ Sent to $targetWorkspaceName" 
                                   } else { 
                                     "❌ Not Sent" 
                                   }
            }
        }
        else {
            # If there are no diagnostic settings, add a custom object with a "No Diagnostic Settings Found" status
            $results += [PSCustomObject]@{
                SubscriptionName = $sub.Name          # Name of the subscription
                FirewallName     = $fw.Name           # Name of the firewall
                ResourceGroup    = $fw.ResourceGroupName  # Name of the resource group containing the firewall
                Region           = $fw.Location       # Location/Region of the firewall
                DiagnosticStatus = "❌ No Diagnostic Settings Found"  # Diagnostic status
            }
        }
    }
}

# Output the results to a table format
$results | Format-Table -AutoSize

# Optional: Export the results to a CSV file (uncomment the following line if needed)
# $results | Export-Csv -Path ".\FirewallDiagAudit-GovCloud.csv" -NoTypeInformation


