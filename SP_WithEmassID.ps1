$subscriptions = Get-AzSubscription

$results = @()

foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id

    # Get all User Assigned Managed Identities
    $uamis = Get-AzUserAssignedIdentity

    foreach ($uami in $uamis) {
        $uamiResource = Get-AzResource -ResourceId $uami.Id

        $vaEmassId = if ($uamiResource.Tags.ContainsKey('va_emass_id')) { $uamiResource.Tags['va_emass_id'] } else { "N/A" }

        $entry = [PSCustomObject]@{
            SubscriptionName = $sub.Name
            SubscriptionId   = $sub.Id
            IdentityName     = $uami.Name
            ResourceGroup    = $uami.ResourceGroupName
            Region           = $uamiResource.Location
            Tags             = ($uamiResource.Tags -join "; ")
            VaEmassId        = $vaEmassId
            IdentityType     = "UserAssigned"
            LoggedIn         = "Unknown"
            ActivityLog      = "Unknown"
            Diagnostics      = "Unknown"
        }

        # Check sign-in logs via MS Graph (not PowerShell-native; requires API integration)
        # Placeholder: $entry.LoggedIn = "GraphAPI check required"

        # Check activity logs
        $activity = Get-AzActivityLog -StartTime (Get-Date).AddDays(-30) `
                    -ResourceId $uami.Id -WarningAction SilentlyContinue

        if ($activity) {
            $entry.ActivityLog = "Yes"
        } else {
            $entry.ActivityLog = "No"
        }

        # Diagnostic settings
        $diag = Get-AzDiagnosticSetting -ResourceId $uami.Id -ErrorAction SilentlyContinue
        if ($diag) {
            $entry.Diagnostics = ($diag.WorkspaceId) ? "LogAnalytics" : "Other"
        } else {
            $entry.Diagnostics = "No"
        }

        $results += $entry
    }

    # System-assigned identities: scan common services (VMs, Functions, WebApps)
    $vms = Get-AzVM
    foreach ($vm in $vms) {
        if ($vm.Identity.Type -match "SystemAssigned") {
            $vmResource = Get-AzResource -ResourceId $vm.Id

            $vaEmassId = if ($vmResource.Tags.ContainsKey('va_emass_id')) { $vmResource.Tags['va_emass_id'] } else { "N/A" }

            $entry = [PSCustomObject]@{
                SubscriptionName = $sub.Name
                SubscriptionId   = $sub.Id
                IdentityName     = $vm.Name
                ResourceGroup    = $vm.ResourceGroupName
                Region           = $vmResource.Location
                Tags             = ($vmResource.Tags -join "; ")
                VaEmassId        = $vaEmassId
                IdentityType     = "SystemAssigned (VM)"
                LoggedIn         = "Unknown"
                ActivityLog      = "Unknown"
                Diagnostics      = "Unknown"
            }

            $vmActivity = Get-AzActivityLog -StartTime (Get-Date).AddDays(-30) `
                          -ResourceId $vm.Id -WarningAction SilentlyContinue

            $entry.ActivityLog = ($vmActivity) ? "Yes" : "No"

            $diag = Get-AzDiagnosticSetting -ResourceId $vm.Id -ErrorAction SilentlyContinue
            $entry.Diagnostics = ($diag) ? "Enabled" : "No"

            $results += $entry
        }
    }

    # You can repeat similar logic for Function Apps, Web Apps, etc.
}

# Export the result
$results | Export-Csv -Path "./ManagedIdentityLoggingReport.csv" -NoTypeInformation

Write-Host "✅ Report generated: ManagedIdentityLoggingReport.csv"