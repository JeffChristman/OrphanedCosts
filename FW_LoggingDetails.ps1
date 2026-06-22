# Login if needed
#Connect-AzAccount

# Get all subscriptions
$subscriptions = Get-AzSubscription

# Output results
$diagData = @()

foreach ($sub in $subscriptions) {
    Write-Host "🔍 Subscription: $($sub.Name)" -ForegroundColor Cyan
    Set-AzContext -SubscriptionId $sub.Id | Out-Null

    # Get all firewall resources
    $firewalls = Get-AzResource -ResourceType "Microsoft.Network/azureFirewalls"

    foreach ($fw in $firewalls) {
        Write-Host "  → Checking diagnostics for: $($fw.Name)"

        # Get diagnostic settings associated with this resource
        $diagSettings = Get-AzResource -ResourceType "Microsoft.Insights/diagnosticSettings" `
                                       -ResourceName "*" `
                                       -ResourceGroupName $fw.ResourceGroupName `
                                       -ExpandProperties `
                                       | Where-Object { $_.Properties?.Scope -eq $fw.ResourceId -or $_.Properties?.targetResourceId -eq $fw.ResourceId }

        foreach ($setting in $diagSettings) {
            $diagName = $setting.Name
            $workspaceId = $setting.Properties.workspaceId

            foreach ($log in $setting.Properties.logs) {
                $diagData += [PSCustomObject]@{
                    SubscriptionName     = $sub.Name
                    FirewallName         = $fw.Name
                    ResourceGroup        = $fw.ResourceGroupName
                    DiagnosticSetting    = $diagName
                    WorkspaceResourceId  = $workspaceId
                    LogCategory          = $log.category
                    LoggingEnabled       = $log.enabled
                }
            }
        }
    }
}

# Output table and save
$diagData | Format-Table -AutoSize
$diagData | Export-Csv -Path "./FirewallDiagSettings_JSONExtract.csv" -NoTypeInformation

Write-Host "`n✅ JSON-extracted diagnostic setting report saved to FirewallDiagSettings_JSONExtract.csv"
