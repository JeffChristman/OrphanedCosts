<#
Checks Key Vaults against CIS-style policies:
- Defender for Key Vault enabled (subscription-level)
- Logging (diagnostic settings) configured
- All keys have expiration date
- All secrets have expiration date
- Vault recoverable (soft delete + purge protection)
- RBAC enabled
- Private endpoints used
- Automatic key rotation enabled (per key, where policy exists)
- Vault actually used to store secrets (basic heuristic)

Outputs: one row per vault with boolean flags.
#>

Import-Module Az.Accounts
Import-Module Az.KeyVault
Import-Module Az.Monitor
Import-Module Az.Security

#Connect-AzAccount
Connect-AzAccount -EnvironmentName AzureUSGovernment -Tenant 'vaazuregov.onmicrosoft.com'

$results = @()

$subs = Get-AzSubscription
foreach ($sub in $subs) {
    Write-Host "Processing subscription $($sub.Name) ($($sub.Id))..." -ForegroundColor Cyan
    Set-AzContext -SubscriptionId $sub.Id | Out-Null

    # Defender for Key Vault is subscription-wide
    $defenderKv = $null
    try {
        $defenderKv = Get-AzSecurityPricing -Name "KeyVaults" -ErrorAction Stop
    } catch {
        Write-Warning "Could not query Defender pricing for KeyVaults in $($sub.Name): $_"
    }
    $defenderKvOn = $false
    if ($defenderKv -and $defenderKv.PricingTier -match "Standard|P1|P2") {
        $defenderKvOn = $true
    }

    $vaults = Get-AzKeyVault
    foreach ($vault in $vaults) {

        Write-Host "  Checking vault $($vault.VaultName) in RG $($vault.ResourceGroupName)" -ForegroundColor Yellow

        # --- Logging / diagnostic settings ---
        $diagSettings = $null
        $loggingEnabled = $false
        try {
            $diagSettings = Get-AzDiagnosticSetting -ResourceId $vault.ResourceId -ErrorAction Stop
            if ($diagSettings) {
                # At least one destination configured
                $loggingEnabled = $diagSettings.Logs.Enabled -contains $true -or
                                  $diagSettings.Metrics.Enabled -contains $true
            }
        } catch {
            # No diagnostic settings or error
            $loggingEnabled = $false
        }

        # --- Recoverable: soft delete + purge protection ---
        $softDeleteEnabled      = $false
        $purgeProtectionEnabled = $false
        # Property names differ slightly between Az versions, so be defensive
        if ($vault.PSObject.Properties.Name -contains "EnableSoftDelete") {
            $softDeleteEnabled = [bool]$vault.EnableSoftDelete
        } elseif ($vault.PSObject.Properties.Name -contains "SoftDeleteEnabled") {
            $softDeleteEnabled = [bool]$vault.SoftDeleteEnabled
        }
        if ($vault.PSObject.Properties.Name -contains "EnablePurgeProtection") {
            $purgeProtectionEnabled = [bool]$vault.EnablePurgeProtection
        }

        $vaultRecoverable = $softDeleteEnabled -and $purgeProtectionEnabled

        # --- RBAC enabled ---
        $rbacEnabled = $false
        if ($vault.PSObject.Properties.Name -contains "EnableRbacAuthorization") {
            $rbacEnabled = [bool]$vault.EnableRbacAuthorization
        }

        # --- Private endpoints ---
        $peCount = 0
        if ($vault.PSObject.Properties.Name -contains "PrivateEndpointConnections" -and
            $vault.PrivateEndpointConnections) {
            $peCount = $vault.PrivateEndpointConnections.Count
        }
        $usesPrivateEndpoints = $peCount -gt 0

        # --- Keys: expiration + rotation policy ---
        $allKeysHaveExpiry   = $true
        $keysMissingExpiry   = @()
        $anyKeyHasRotation   = $false
        $keysWithoutRotation = @()

        try {
            $keys = Get-AzKeyVaultKey -VaultName $vault.VaultName -ErrorAction Stop
            foreach ($key in $keys) {
                $exp = $key.Attributes.Expires
                if (-not $exp) {
                    $allKeysHaveExpiry = $false
                    $keysMissingExpiry += $key.Name
                }

                # Rotation policy (if cmdlet/version supports it)
                try {
                    $rotation = Get-AzKeyVaultKeyRotationPolicy -VaultName $vault.VaultName -Name $key.Name -ErrorAction Stop
                    if ($rotation -and $rotation.LifetimeActions.Count -gt 0) {
                        $anyKeyHasRotation = $true
                    } else {
                        $keysWithoutRotation += $key.Name
                    }
                } catch {
                    # Rotation policy not defined or cmdlet not available - treat as no rotation
                    $keysWithoutRotation += $key.Name
                }
            }
        } catch {
            # No keys or no access
            $allKeysHaveExpiry = $false
        }

        # --- Secrets: expiration dates ---
        $allSecretsHaveExpiry = $true
        $secretsMissingExpiry = @()
        $hasSecrets           = $false

        try {
            $secrets = Get-AzKeyVaultSecret -VaultName $vault.VaultName -ErrorAction Stop
            if ($secrets) { $hasSecrets = $true }
            foreach ($secret in $secrets) {
                $sexp = $secret.Attributes.Expires
                if (-not $sexp) {
                    $allSecretsHaveExpiry = $false
                    $secretsMissingExpiry += $secret.Name
                }
            }
        } catch {
            # No secrets or insufficient permissions
            $allSecretsHaveExpiry = $false
        }

        # --- Assemble result row ---
        $results += [PSCustomObject]@{
            SubscriptionName                 = $sub.Name
            SubscriptionId                   = $sub.Id
            ResourceGroup                    = $vault.ResourceGroupName
            VaultName                        = $vault.VaultName
            Location                         = $vault.Location

            DefenderForKeyVaultOn           = $defenderKvOn
            LoggingEnabled                   = $loggingEnabled

            VaultRecoverable                 = $vaultRecoverable
            SoftDeleteEnabled                = $softDeleteEnabled
            PurgeProtectionEnabled           = $purgeProtectionEnabled

            RbacEnabled                      = $rbacEnabled
            UsesPrivateEndpoints             = $usesPrivateEndpoints
            PrivateEndpointConnectionCount   = $peCount

            AllKeysHaveExpiry                = $allKeysHaveExpiry
            KeysMissingExpiry                = ($keysMissingExpiry -join "; ")

            AnyKeyHasRotationPolicy          = $anyKeyHasRotation
            KeysWithoutRotationPolicy        = ($keysWithoutRotation -join "; ")

            AllSecretsHaveExpiry             = $allSecretsHaveExpiry
            SecretsMissingExpiry             = ($secretsMissingExpiry -join "; ")

            HasSecretsStored                 = $hasSecrets
        }
    }
}

# Output to screen
$results | Format-Table -AutoSize

# Export for your report / Excel workflows
$results | Export-Csv -Path ".\KeyVault_CIS_Assessment.csv" -NoTypeInformation -Encoding UTF8

