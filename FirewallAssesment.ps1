# Azure Firewall Best Practices Audit Script with Scoring
# Requires: Connect-AzAccount and Reader access across all subscriptions

$subscriptions = Get-AzSubscription
$auditResults = @()

foreach ($sub in $subscriptions) {
    Write-Host "`nAuditing subscription: $($sub.Name)"
    Set-AzContext -SubscriptionId $sub.Id | Out-Null

    $firewalls = Get-AzFirewall
    foreach ($fw in $firewalls) {
        $fwName = $fw.Name
        $fwRg = $fw.ResourceGroupName
        $fwId = $fw.Id

        # Diagnostic Settings
        $diag = Get-AzDiagnosticSetting -ResourceId $fwId -ErrorAction SilentlyContinue
        $hasDiag = if ($diag -and $diag.Enabled -eq $true) { $true } else { $false }

        # Threat Intelligence Mode
        $tiMode = $fw.ThreatIntelMode
        $tiConfigured = $tiMode -in @("Alert", "Deny")

        # DNS settings
        $dnsProxy = $fw.DnsSettings.EnableProxy
        $customDns = $fw.DnsSettings.Servers.Count -gt 0

        # Idle timeout (default 4 mins)
        $idleTimeout = $fw.IdleTimeoutInMinutes

        # Rule Analysis
        $highRisk = 0; $mediumRisk = 0; $totalRules = 0
        foreach ($rc in $fw.NetworkRuleCollections + $fw.ApplicationRuleCollections + $fw.NatRuleCollections) {
            foreach ($rule in $rc.Rules) {
                $totalRules++
                $isAnySource = ($rule.SourceAddresses -contains "*")
                $isAnyDest = ($rule.DestinationAddresses -contains "*")
                $isAnyPort = ($rule.DestinationPorts -contains "*")
                $isAllow = $rule.Action -eq "Allow"

                if ($isAllow -and ($isAnySource -or $isAnyDest -or $isAnyPort)) {
                    $highRisk++
                } elseif ($isAnySource -or $isAnyDest) {
                    $mediumRisk++
                }
            }
        }

        # Final risk score and recommendation
        $status = "Pass"
        if (!$hasDiag -or !$tiConfigured -or $highRisk -gt 0) {
            $status = "Attention Required"
        }

        $auditResults += [PSCustomObject]@{
            Subscription        = $sub.Name
            ResourceGroup       = $fwRg
            FirewallName        = $fwName
            DiagnosticEnabled   = $hasDiag
            ThreatIntelMode     = $tiMode
            DnsProxyEnabled     = $dnsProxy
            CustomDnsUsed       = $customDns
            IdleTimeoutMinutes  = $idleTimeout
            TotalRules          = $totalRules
            HighRiskRules       = $highRisk
            MediumRiskRules     = $mediumRisk
            AuditStatus         = $status
        }
    }
}

# Output to CSV
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$auditResults | Export-Csv -Path "$env:USERPROFILE\Desktop\AzureFirewall-Audit-$timestamp.csv" -NoTypeInformation
Write-Host "`nAudit complete. Results saved to Desktop."

