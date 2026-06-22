# Login to Azure if needed
#Connect-AzAccount

# Get all subscriptions
#$subscriptions = Get-AzSubscription

# Gather all Log Analytics Workspaces across all subscriptions
$workspaceLookup = @{}

foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
    Write-Host "🔍 Collecting workspaces in: $($sub.Name)" -ForegroundColor Cyan

    try {
        $workspaces = Get-AzOperationalInsightsWorkspace
        foreach ($ws in $workspaces) {
            $workspaceLookup[$ws.Id.ToLower()] = $ws.Name
        }
    } catch {
        Write-Warning "⚠️ Could not retrieve workspaces in $($sub.Name)"
    }
}

# Initialize final results
$report = @()

foreach ($sub in $subscriptions) {
    Set-AzContext -SubscriptionId $sub.Id | Out-Null
    Write-Host "`n📂 Processing Subscription: $($sub.Name)" -ForegroundColor Cyan

    # Step 1: Get all firewalls in the subscription
    try {
        $firewalls = Get-AzFirewall
    } catch {
        Write-Warning "⚠️ Failed to retrieve firewalls in $($sub.Name)"
        continue
    }

    # Step 2: Get policy assignments (optional, can skip if not needed)
    $assignments = Get-AzPolicyAssignment
    $diagAssignments = $assignments | Where-Object {
        ($_.Properties.DisplayName -like "*diagnostic*" -or $_.Properties.DisplayName -like "*log analytics*") -and
        ($_.Properties.Scope -like "*Microsoft.Network/azureFirewalls*")
    }

    foreach ($fw in $firewalls) {
        Write-Host "🔎 Firewall: $($fw.Name)"

        $diag = $null
        $workspaceName = "None"
        $workspaceId = "None"
        $policyName = "None"
        $policyDisplayName = "None"
        $complianceState = "Unknown"

        # Step 3: Get diagnostic settings
        try {
            $diag = Get-AzDiagnosticSetting -ResourceId $fw.Id
        } catch {
            Write-Warning "  ⚠️ No diagnostic settings for $($fw.Name)"
        }

        if ($diag) {
            if ($diag.WorkspaceId) {
                $lookupKey = $diag.WorkspaceId.ToLower()
                $workspaceId = $diag.WorkspaceId
                if ($workspaceLookup.ContainsKey($lookupKey)) {
                    $workspaceName = $workspaceLookup[$lookupKey]
                } else {
                    $workspaceName = "Unknown"
                }
            }
        }

        # Step 4: Match policy assignment (if scoped to the firewall or resource group)
        $relatedPolicy = $diagAssignments | Where-Object {
            $_.Properties.Scope -like "*$($fw.Name)*" -or $_.Properties.Scope -like "*$($fw.ResourceGroupName)*"
        } | Select-Object -First 1

        if ($relatedPolicy) {
            $policyName = $relatedPolicy.Name
            $policyDisplayName = $relatedPolicy.Properties.DisplayName
        }

        # Step 5: Check compliance state
        try {
            $compliance = Get-AzPolicyState -ResourceId $fw.Id -Top 1 -ErrorAction SilentlyContinue
            if ($compliance) {
                $complianceState = $compliance.ComplianceState
            }
        } catch {
            $complianceState = "Unknown"
        }

        # Step 6: Add log data
        if ($diag) {
            foreach ($log in $diag.Logs) {
                $report += [PSCustomObject]@{
                    SubscriptionName     = $sub.Name
                    SubscriptionId       = $sub.Id
                    FirewallName         = $fw.Name
                    ResourceGroup        = $fw.ResourceGroupName
                    Location             = $fw.Location
                    WorkspaceName        = $workspaceName
                    WorkspaceResourceId  = $workspaceId
                    LogCategory          = $log.Category
                    LoggingEnabled       = $log.Enabled
                    PolicyAssignment     = $policyName
                    PolicyDisplayName    = $policyDisplayName
                    ComplianceState      = $complianceState
                }
            }
        } else {
            $report += [PSCustomObject]@{
                SubscriptionName     = $sub.Name
                SubscriptionId       = $sub.Id
                FirewallName         = $fw.Name
                ResourceGroup        = $fw.ResourceGroupName
                Location             = $fw.Location
                WorkspaceName        = "None"
                WorkspaceResourceId  = "None"
                LogCategory          = "None"
                LoggingEnabled       = $false
                PolicyAssignment     = $policyName
                PolicyDisplayName    = $policyDisplayName
                ComplianceState      = $complianceState
            }
        }
    }
}

# Output results and export
$report | Format-Table -AutoSize
$report | Export-Csv -Path "./AzureFirewall_Diagnostics_Report.csv" -NoTypeInformation

Write-Host "`n✅ Report complete: AzureFirewall_Diagnostics_Report.csv"

