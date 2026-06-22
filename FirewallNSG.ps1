
# Enhanced Azure Firewall and NSG Compliance Assessment Script
# This script includes subscription summaries, deduplication of NSG rules, and opportunities for NSG rule consolidation.

# Connect to Azure
#Connect-AzAccount

# Get all subscriptions
$subscriptions = Get-AzSubscription

# Initialize arrays to store results, NSG results, summary, and rule hashes for deduplication
$results         = @()
$nsgResults      = @()
$summary         = @()
$ruleHashes      = @{}

foreach ($sub in $subscriptions) {
    # Set the context to the current subscription
    Set-AzContext -SubscriptionId $sub.Id

    # Reset per-subscription counts
    $subNSGCount      = 0
    $subFirewallCount = 0

    # Azure Firewall Rules Section
    $firewalls        = Get-AzFirewall
    $subFirewallCount = $firewalls.Count

    foreach ($fw in $firewalls) {
        # Determine if the firewall has a policy and rules
        $hasPolicy   = if ($fw.FirewallPolicy) { "Yes" } else { "No" }
        $hasRules    = if (($fw.NetworkRuleCollections.Count -eq 0) -and ($fw.ApplicationRuleCollections.Count -eq 0)) { "No" } else { "Yes" }

        try {
            # Get diagnostic settings for the firewall
            $diag       = Get-AzDiagnosticSetting -ResourceId $fw.Id
            $diagStatus = if ($diag.Count -eq 0) { "Disabled" } else { "Enabled" }
        } catch {
            $diagStatus = "Error"
        }

        # Iterate through network rule collections and rules within each collection
        foreach ($collection in $fw.NetworkRuleCollections) {
            foreach ($rule in $collection.Rules) {
                # Add a custom object representing each rule to the results array
                $results += [PSCustomObject]@{
                    Subscription         = $sub.Name
                    FirewallName         = $fw.Name
                    ResourceGroup        = $fw.ResourceGroupName
                    Location             = $fw.Location
                    SKU                  = $fw.Sku.Name
                    ThreatIntelMode      = $fw.ThreatIntelMode
                    TLSInspection        = $fw.EnableTLSInspection
                    FirewallPolicy       = $hasPolicy
                    RuleCollectionsExist = $hasRules
                    Zones                = ($fw.Zones            -join ", ")
                    DiagnosticsEnabled   = $diagStatus
                    Protocols            = ($rule.Protocols      -join ", ")
                    Source               = ($rule.SourceAddresses      -join ", ")
                    Destination          = ($rule.DestinationAddresses -join ", ")
                    Port                 = ($rule.DestinationPorts      -join ", ")
                }
            }
        }
    }

    # NSG Rules & Consolidation Section
    $nsgs        = Get-AzNetworkSecurityGroup
    $subNSGCount = $nsgs.Count

    foreach ($nsg in $nsgs) {
        foreach ($rule in $nsg.SecurityRules) {
            # Determine compliance issues based on rule properties
            $issue      = ""
            $compliance = "Compliant"
            $hash       = ($rule.Direction + $rule.Access + $rule.Priority +
                           $rule.SourceAddressPrefix + $rule.DestinationAddressPrefix +
                           $rule.DestinationPortRange + $rule.Protocol)

            if ($rule.Access -eq "Allow" -and $rule.Direction -eq "Inbound" -and
               ($rule.SourceAddressPrefix -eq "*" -or $rule.SourceAddressPrefix -eq "0.0.0.0/0")) {

                if ($rule.DestinationPortRange -eq "*" -or $rule.DestinationPortRange -eq "0-65535") {
                    $issue = "Overly permissive: Any source to any port"
                }
                elseif ($rule.DestinationPortRange -in @("22", "3389", "80", "443")) {
                    $issue = "Common exposed port to the internet"
                }
                else {
                    $issue = "General overly permissive rule"
                }
                $compliance = "Non-Compliant"
            }

            # Suggest consolidation if the rule is duplicated
            $suggestConsolidation = if ($ruleHashes.ContainsKey($hash)) {
                "Yes"
            } else {
                $ruleHashes[$hash] = 1
                "No"
            }

            # Add a custom object representing each NSG rule to the nsgResults array
            $nsgResults += [PSCustomObject]@{
                Subscription         = $sub.Name
                NSGName              = $nsg.Name
                RuleName             = $rule.Name
                Direction            = $rule.Direction
                Access               = $rule.Access
                Priority             = $rule.Priority
                Source               = (@($rule.SourceAddressPrefix)      -join ", ")
                Destination          = (@($rule.DestinationAddressPrefix) -join ", ")
                Port                 = (@($rule.DestinationPortRange)     -join ", ")
                Protocol             = $rule.Protocol
                IssueDetected        = $issue
                ComplianceStatus     = $compliance
                SuggestConsolidation = $suggestConsolidation
            }

            # Output to console if a non-compliant rule is detected
            if ($compliance -eq "Non-Compliant") {
                Write-Host "Non-Compliant Rule Detected in $($nsg.Name) - $($rule.Name): $issue" -ForegroundColor Red
            }
        }
    }

    # Subscription Summary Section
    $summary += [PSCustomObject]@{
        Subscription   = $sub.Name
        NSG_Count      = $subNSGCount
        Firewall_Count = $subFirewallCount
    }
}

# Output results to the console in table format
$summary    | Format-Table -AutoSize
$results    | Format-Table -AutoSize
$nsgResults | Format-Table -AutoSize

# Export results to CSV files
$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$results    | Export-Csv -Path "./AzureFirewallAssessment_$timestamp.csv"   -NoTypeInformation
$nsgResults | Export-Csv -Path "./AzureNSGAssessment_$timestamp.csv"        -NoTypeInformation
$summary    | Export-Csv -Path "./AzureAssessment_SubscriptionSummary_$timestamp.csv" -NoTypeInformation
```

