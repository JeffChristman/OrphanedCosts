

<#
.SYNOPSIS
  Generate a best-practice / security configuration report for AKS clusters.

.DESCRIPTION
  For each AKS managed cluster (Microsoft.ContainerService/managedClusters), the script
  evaluates key controls:

    - Kubernetes RBAC enabled
    - AKS-managed Entra ID integration (managed AAD)
    - Entra ID Azure RBAC enabled
    - Local accounts (--admin) disabled
    - Private cluster or API server authorized IP ranges configured
    - Network policy enabled (azure or calico)
    - Azure Policy add-on enabled
    - OMS/Log Analytics monitoring enabled
    - Defender add-on enabled (if present)
    - Any node pools with public IPs
    - Any node pools with encryptionAtHost and/or FIPS enabled
    - Diagnostic settings configured

.REQUIREMENTS
  Az.Accounts, Az.Aks, Az.Monitor
  Connect-AzAccount first, or run in an authenticated context.
#>

[CmdletBinding()]
param(
    [switch]$AllSubscriptions,
    [string[]]$SubscriptionIds,
    [string]$OutputPath
)

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Aks      -ErrorAction Stop
Import-Module Az.Monitor  -ErrorAction Stop

# Resolve subscriptions
$subs = @()

if ($AllSubscriptions) {
    $subs = Get-AzSubscription
}
elseif ($SubscriptionIds) {
    $subs = Get-AzSubscription | Where-Object { $SubscriptionIds -contains $_.SubscriptionId }
}
else {
    $ctx = Get-AzContext
    if (-not $ctx) {
        throw "No current Az context. Run Connect-AzAccount first."
    }
    $subs = ,$ctx.Subscription
}

$results = @()

foreach ($sub in $subs) {
    Write-Verbose "Processing subscription $($sub.Name) [$($sub.Id)]"
    Select-AzSubscription -SubscriptionId $sub.Id | Out-Null

    $clusters = Get-AzAksCluster -SubscriptionId $sub.Id

    foreach ($c in $clusters) {

        # Core properties
        $aadProfile  = $c.AadProfile
        $netProfile  = $c.NetworkProfile
        $apiProfile  = $c.ApiServerAccessProfile
        $addonProfiles = $c.AddonProfiles
        $agentPools = $c.AgentPoolProfiles

        # AAD / RBAC
        $rbacEnabled    = $c.EnableRBAC
        $aadManaged     = $false
        $aadAzureRBAC   = $false
        if ($aadProfile) {
            $aadManaged   = $aadProfile.Managed
            $aadAzureRBAC = $aadProfile.EnableAzureRBAC
        }

        # Local admin accounts (az aks get-credentials --admin)
        # PS object property is typically DisableLocalAccount (singular)
        $localAccountsDisabled = $false
        if ($null -ne $c.DisableLocalAccount) {
            $localAccountsDisabled = [bool]$c.DisableLocalAccount
        }

        # API server access / private cluster
        $isPrivateCluster = $false
        $authorizedIpConfigured = $false
        if ($apiProfile) {
            if ($apiProfile.EnablePrivateCluster) {
                $isPrivateCluster = $true
            }
            if ($apiProfile.AuthorizedIPRanges -and $apiProfile.AuthorizedIPRanges.Count -gt 0) {
                $authorizedIpConfigured = $true
            }
        }

        # Network policy
        $networkPolicy            = $null
        $networkPolicyCompliant   = $false
        if ($netProfile) {
            $networkPolicy = $netProfile.NetworkPolicy
            if ($networkPolicy -in @('azure', 'calico')) {
                $networkPolicyCompliant = $true
            }
        }

        # Add-ons
        $azurePolicyEnabled  = $false
        $omsEnabled          = $false
        $defenderEnabled     = $false

        if ($addonProfiles) {
            if ($addonProfiles.ContainsKey('azurepolicy')) {
                $azurePolicyEnabled = [bool]$addonProfiles['azurepolicy'].Enabled
            }
            if ($addonProfiles.ContainsKey('omsagent')) {
                $omsEnabled = [bool]$addonProfiles['omsagent'].Enabled
            }
            if ($addonProfiles.ContainsKey('microsoftDefender')) {
                $defenderEnabled = [bool]$addonProfiles['microsoftDefender'].Enabled
            }
        }

        # Agent pool flags
        $anyNodePublicIP        = $false
        $anyNodeEncryptionAtHost = $false
        $anyNodeFips            = $false

        if ($agentPools) {
            foreach ($pool in $agentPools) {
                if ($pool.EnableNodePublicIP) {
                    $anyNodePublicIP = $true
                }
                if ($pool.EnableEncryptionAtHost) {
                    $anyNodeEncryptionAtHost = $true
                }
                if ($pool.EnableFIPS) {
                    $anyNodeFips = $true
                }
            }
        }

        # Diagnostic settings on the managed cluster resource
        $diag = $null
        try {
            $diag = Get-AzDiagnosticSetting -ResourceId $c.Id -ErrorAction Stop
        }
        catch {
            # No diag or no permission
        }
        $hasDiag = $ -and ($diag | Measure-Object).Count -gt 0

        $results += [PSCustomObject]@{
            SubscriptionId            = $sub.Id
            SubscriptionName          = $sub.Name
            ResourceGroup             = $c.ResourceGroupName
            ClusterName              = $c.Name
            Location                  = $c.Location
            KubernetesVersion         = $c.KubernetesVersion

            # Identity / RBAC
            EnableRBAC                = $rbacEnabled
            AadManaged                = $aadManaged
            AadAzureRBAC              = $aadAzureRBAC
            LocalAccountsDisabled     = $localAccountsDisabled

            # API server & network
            PrivateCluster            = $isPrivateCluster
            ApiAuthorizedIPRangesSet  = $authorizedIpConfigured
            NetworkPolicy             = $networkPolicy
            NetworkPolicyCompliant    = $networkPolicyCompliant

            # Add-ons
            AzurePolicyAddonEnabled   = $azurePolicyEnabled
            OMSMonitoringAddonEnabled = $omsEnabled
            DefenderAddonEnabled      = $defenderEnabled

            # Node pools
            AnyNodePublicIP           = $anyNodePublicIP
            AnyNodeEncryptionAtHost   = $anyNodeEncryptionAtHost
            AnyNodeFIPS               = $anyNodeFips

            # Monitoring
            HasDiagnosticSettings     = $hasDiag

            # Raw objects if you want to dig further
            RawAadProfile             = $aadProfile
            RawNetworkProfile         = $netProfile
            RawApiServerAccessProfile = $apiProfile
            RawAddonProfiles          = $addonProfiles
            RawAgentPools             = $agentPools
        }
    }
}

if ($OutputPath) {
    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported report to $OutputPath"
}

$results

