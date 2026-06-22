# Connect to Azure GovCloud
#Connect-AzAccount -Environment AzureUSGovernment

# Set the target Log Analytics Workspace Name
$targetWorkspaceName = "vaecla-security-gov"
$results = @()

# Get all subscriptions
$subscriptions = Get-AzSubscription

foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id

    # Get all firewalls in this subscription
    $firewalls = Get-AzFirewall

    foreach ($fw in $firewalls) {
        $diagSettings = Get-AzDiagnosticSetting -ResourceId $fw.Id -ErrorAction SilentlyContinue

        if ($diagSettings) {
            $workspaceFound = $false
            foreach ($ws in $diagSettings.WorkspaceId) {
                $workspace = Get-AzOperationalInsightsWorkspace -ResourceId $ws -ErrorAction SilentlyContinue
                if ($workspace.Name -eq $targetWorkspaceName) {
                    $workspaceFound = $true
                    break
                }
            }

            $results += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                FirewallName     = $fw.Name
                ResourceGroup    = $fw.ResourceGroupName
                Region           = $fw.Location
                DiagnosticStatus = if ($workspaceFound) { "✅ Sent to $targetWorkspaceName" } else { "❌ Not Sent" }
            }
        }
        else {
            $results += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                FirewallName     = $fw.Name
                ResourceGroup    = $fw.ResourceGroupName
                Region           = $fw.Location
                DiagnosticStatus = "❌ No Diagnostic Settings Found"
            }
        }
    }
}

# Output to table
$results | Format-Table -AutoSize

# Optional: Export to CSV
#$results | Export-Csv -Path ".\FirewallDiagAudit-GovCloud.csv" -NoTypeInformation
