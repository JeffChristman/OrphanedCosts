# Requires Az module
#Connect-AzAccount

$subscriptions = Get-AzSubscription
$results = @()

foreach ($sub in $subscriptions) {
    Write-Host "Processing subscription: $($sub.Name)" -ForegroundColor Cyan
    Set-AzContext -SubscriptionId $sub.Id

    $apps = Get-AzWebApp

    foreach ($app in $apps) {
        $resourceId = "/subscriptions/$($sub.Id)/resourceGroups/$($app.ResourceGroup)/providers/Microsoft.Web/sites/$($app.Name)?api-version=2022-03-01"

        # Use REST API to fetch the raw App Service JSON
        $webAppDetails = Invoke-AzRestMethod -Method GET -Path $resourceId
        $json = $webAppDetails.Content | ConvertFrom-Json

        # Try to access the virtual IP property (only present in some configurations)
        $vip = $json.properties.virtualIp

        $results += [PSCustomObject]@{
            Subscription   = $sub.Name
            AppName        = $app.Name
            ResourceGroup  = $app.ResourceGroup
            Location       = $app.Location
            VirtualIP      = if ($vip) { $vip } else { "None" }
        }
    }
}

# Output results
$results | Format-Table -AutoSize
$results | Export-Csv -Path "./AppServices_VirtualIPs.csv" -NoTypeInformation
