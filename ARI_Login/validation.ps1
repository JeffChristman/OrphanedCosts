<#
.SYNOPSIS
    CIS v2.0 Azure Government Validation Scripts
    Generated from Turbot ALARM export – Controls_by_state_20260428105827.xlsx
    Target environment: Azure US Government (AzureUSGovernment)

.NOTES
    Prerequisites:
        Install-Module Az -Scope CurrentUser
        Connect-AzAccount -Environment AzureUSGovernment

    Usage:
        Invoke-AllCISChecks                          # interactive subscription picker
        Invoke-AllCISChecks -SubscriptionId '<guid>' # target a specific subscription
        Invoke-AllCISChecks                          # choose 'A' at the prompt for all subs

    All check functions also accept -Subscriptions directly if called individually:
        $subs = Select-AzGovSubscription
        Test-CIS_7_04 -Subscriptions $subs
#>

#region ── HELPERS ─────────────────────────────────────────────────────────────

function Connect-AzGov {
    $ctx = Get-AzContext
    if (-not $ctx -or $ctx.Environment.Name -ne 'AzureUSGovernment') {
        Connect-AzAccount -Environment AzureUSGovernment
    }
}

function Select-AzGovSubscription {
    <#
    .SYNOPSIS
        Interactive subscription picker. Returns one or more subscription objects.
    .PARAMETER SubscriptionId
        Skip the menu by supplying a known subscription ID.
    #>
    param([string]$SubscriptionId)

    $subs = Get-AzSubscription | Sort-Object Name
    if (-not $subs) { Write-Error "No subscriptions found."; return $null }

    if ($SubscriptionId) {
        $chosen = $subs | Where-Object { $_.Id -eq $SubscriptionId }
        if (-not $chosen) { Write-Error "Subscription '$SubscriptionId' not found."; return $null }
        Set-AzContext -SubscriptionId $chosen.Id | Out-Null
        Write-Host "Context set to: $($chosen.Name)" -ForegroundColor Green
        return @($chosen)
    }

    Write-Host "`nAvailable Subscriptions:" -ForegroundColor Cyan
    Write-Host ("-" * 70)
    for ($i = 0; $i -lt $subs.Count; $i++) {
        Write-Host ("  [{0,2}] {1}" -f ($i + 1), $subs[$i].Name) -NoNewline
        Write-Host "  ($($subs[$i].Id))" -ForegroundColor DarkGray
    }
    Write-Host ("-" * 70)

    do {
        $sel = Read-Host "Select subscription [1-$($subs.Count)], comma-separated list, or 'A' for all"
        if ($sel -match '^[Aa]$') {
            Write-Host "Running against ALL $($subs.Count) subscriptions." -ForegroundColor Yellow
            return $subs
        }
        # Support comma-separated multi-select e.g. "1,3,5"
        $indices = $sel -split ',' | ForEach-Object { [int]$_.Trim() - 1 }
        $valid   = $indices | Where-Object { $_ -ge 0 -and $_ -lt $subs.Count }
    } until ($valid.Count -gt 0 -and $valid.Count -eq $indices.Count)

    $chosen = $subs[$valid]
    if ($chosen.Count -eq 1) {
        Set-AzContext -SubscriptionId $chosen.Id | Out-Null
        Write-Host "Context set to: $($chosen.Name)" -ForegroundColor Green
    } else {
        Write-Host "Selected $($chosen.Count) subscriptions." -ForegroundColor Green
    }
    return @($chosen)
}

function Out-CISResult {
    param($ControlID, $SubscriptionName, $ResourceName, $Status, $Detail)
    [PSCustomObject]@{
        ControlID        = $ControlID
        SubscriptionName = $SubscriptionName
        ResourceName     = $ResourceName
        Status           = $Status   # PASS / FAIL / UNKNOWN
        Detail           = $Detail
        CheckedAt        = (Get-Date -Format 'o')
    }
}

#endregion

#region ── 02 – MICROSOFT DEFENDER ────────────────────────────────────────────

function Test-CIS_2_01_03 {
    <# Ensure Microsoft Defender for Databases is Set to 'On' #>
    param([object[]]$Subscriptions)
    $id = '2.01.03'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        foreach ($plan in @('SqlServers','OpenSourceRelationalDatabases')) {
            $p      = Get-AzSecurityPricing -Name $plan -ErrorAction SilentlyContinue
            $status = if ($p.PricingTier -eq 'Standard') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $plan $status "Defender tier: $($p.PricingTier)"
        }
    }
}

function Test-CIS_2_01_14 {
    <# Ensure ASC Default Policy Settings are Not Set to 'Disabled' #>
    param([object[]]$Subscriptions)
    $id = '2.01.14'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $policies = Get-AzPolicyAssignment | Where-Object { $_.Properties.DisplayName -like '*ASC Default*' -or $_.Name -like '*SecurityCenter*' }
        if ($policies) {
            foreach ($p in $policies) {
                $disabled = $p.Properties.Parameters.PSObject.Properties | Where-Object { $_.Value.value -eq 'Disabled' }
                $status   = if ($disabled) { 'FAIL' } else { 'PASS' }
                Out-CISResult $id $sub.Name $p.Name $status "Disabled params: $($disabled.Name -join ', ')"
            }
        } else {
            Out-CISResult $id $sub.Name 'N/A' 'UNKNOWN' 'No ASC default policy assignment found'
        }
    }
}

function Test-CIS_2_01_15 {
    <# Ensure Auto provisioning of Log Analytics agent for Azure VMs is 'On' #>
    param([object[]]$Subscriptions)
    $id = '2.01.15'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $ap     = Get-AzSecurityAutoProvisioningSetting -Name 'mma' -ErrorAction SilentlyContinue
        $status = if ($ap.AutoProvision -eq 'On') { 'PASS' } else { 'FAIL' }
        Out-CISResult $id $sub.Name $sub.Name $status "MMA auto-provision: $($ap.AutoProvision)"
    }
}

function Test-CIS_2_01_19 {
    <# Ensure 'Additional email addresses' is Configured with a Security Contact Email #>
    param([object[]]$Subscriptions)
    $id = '2.01.19'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $contacts = Get-AzSecurityContact -ErrorAction SilentlyContinue
        $hasEmail = $contacts | Where-Object { $_.Email -match '@' }
        $status   = if ($hasEmail) { 'PASS' } else { 'FAIL' }
        Out-CISResult $id $sub.Name $sub.Name $status "Security contacts: $($contacts.Email -join '; ')"
    }
}

#endregion

#region ── 03 – STORAGE ACCOUNTS ──────────────────────────────────────────────

function Test-CIS_3_01 {
    <# Ensure 'Secure transfer required' is Enabled #>
    param([object[]]$Subscriptions)
    $id = '3.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $status = if ($_.EnableHttpsTrafficOnly) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "EnableHttpsTrafficOnly=$($_.EnableHttpsTrafficOnly)"
        }
    }
}

function Test-CIS_3_02 {
    <# Ensure Infrastructure Encryption is Enabled for each Storage Account #>
    param([object[]]$Subscriptions)
    $id = '3.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $enc    = $_.Encryption.RequireInfrastructureEncryption
            $status = if ($enc) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "InfraEncryption=$enc"
        }
    }
}

function Test-CIS_3_05 {
    <# Ensure Storage Logging is Enabled for Queue Service (Read/Write/Delete) #>
    param([object[]]$Subscriptions)
    $id = '3.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $sa = $_
            try {
                $log    = Get-AzStorageServiceLoggingProperty -ServiceType Queue -Context $sa.Context -ErrorAction Stop
                $ok     = $log.LoggingOperations -band [Microsoft.Azure.Storage.Shared.Protocol.LoggingOperations]::All
                $status = if ($ok) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $sa.StorageAccountName $status "Queue logging: $($log.LoggingOperations)"
            } catch {
                Out-CISResult $id $sub.Name $sa.StorageAccountName 'UNKNOWN' $_.Exception.Message
            }
        }
    }
}

function Test-CIS_3_08 {
    <# Ensure Default Network Access Rule for Storage Accounts is Set to Deny #>
    param([object[]]$Subscriptions)
    $id = '3.08'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $rule   = $_.NetworkRuleSet.DefaultAction
            $status = if ($rule -eq 'Deny') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "DefaultNetworkAction=$rule"
        }
    }
}

function Test-CIS_3_09 {
    <# Ensure 'Allow Azure services on trusted list' is Enabled #>
    param([object[]]$Subscriptions)
    $id = '3.09'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $bypass = $_.NetworkRuleSet.Bypass
            $status = if ($bypass -match 'AzureServices') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "Bypass=$bypass"
        }
    }
}

function Test-CIS_3_10 {
    <# Ensure Private Endpoints are used to access Storage Accounts #>
    param([object[]]$Subscriptions)
    $id = '3.10'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $pe     = $_.PrivateEndpointConnections | Where-Object { $_.PrivateLinkServiceConnectionState.Status -eq 'Approved' }
            $status = if ($pe) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "Approved private endpoints: $($pe.Count)"
        }
    }
}

function Test-CIS_3_11 {
    <# Ensure Soft Delete is Enabled for Azure Containers and Blob Storage #>
    param([object[]]$Subscriptions)
    $id = '3.11'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $sa = $_
            try {
                $svc    = Get-AzStorageBlobServiceProperty -StorageAccountName $sa.StorageAccountName -ResourceGroupName $sa.ResourceGroupName -ErrorAction Stop
                $blobOK = $svc.DeleteRetentionPolicy.Enabled
                $contOK = $svc.ContainerDeleteRetentionPolicy.Enabled
                $status = if ($blobOK -and $contOK) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $sa.StorageAccountName $status "BlobSoftDelete=$blobOK ContainerSoftDelete=$contOK"
            } catch {
                Out-CISResult $id $sub.Name $sa.StorageAccountName 'UNKNOWN' $_.Exception.Message
            }
        }
    }
}

function Test-CIS_3_12 {
    <# Ensure Storage for Critical Data is Encrypted with Customer Managed Keys #>
    param([object[]]$Subscriptions)
    $id = '3.12'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $keySource = $_.Encryption.KeySource
            $status    = if ($keySource -eq 'Microsoft.Keyvault') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "KeySource=$keySource"
        }
    }
}

function Test-CIS_3_13 {
    <# Ensure Storage Logging is Enabled for Blob Service (Read/Write/Delete) #>
    param([object[]]$Subscriptions)
    $id = '3.13'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $sa = $_
            try {
                $log    = Get-AzStorageServiceLoggingProperty -ServiceType Blob -Context $sa.Context -ErrorAction Stop
                $ok     = $log.LoggingOperations -band [Microsoft.Azure.Storage.Shared.Protocol.LoggingOperations]::All
                $status = if ($ok) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $sa.StorageAccountName $status "Blob logging: $($log.LoggingOperations)"
            } catch {
                Out-CISResult $id $sub.Name $sa.StorageAccountName 'UNKNOWN' $_.Exception.Message
            }
        }
    }
}

function Test-CIS_3_15 {
    <# Ensure Minimum TLS version is set to 1.2 #>
    param([object[]]$Subscriptions)
    $id = '3.15'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzStorageAccount | ForEach-Object {
            $tls    = $_.MinimumTlsVersion
            $status = if ($tls -eq 'TLS1_2') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.StorageAccountName $status "MinTLS=$tls"
        }
    }
}

#endregion

#region ── 04 – DATABASE SERVICES ─────────────────────────────────────────────

function Test-CIS_4_01_01 {
    <# Ensure 'Auditing' is set to 'On' for SQL Servers #>
    param([object[]]$Subscriptions)
    $id = '4.01.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $audit  = Get-AzSqlServerAudit -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName
            $status = if ($audit.BlobStorageTargetState -eq 'Enabled' -or $audit.LogAnalyticsTargetState -eq 'Enabled') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "BlobAudit=$($audit.BlobStorageTargetState) LAW=$($audit.LogAnalyticsTargetState)"
        }
    }
}

function Test-CIS_4_01_02 {
    <# Ensure no SQL Databases allow ingress from 0.0.0.0/0 #>
    param([object[]]$Subscriptions)
    $id = '4.01.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv   = $_
            $rules = Get-AzSqlServerFirewallRule -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName
            $open  = $rules | Where-Object { $_.StartIpAddress -eq '0.0.0.0' }
            $status = if ($open) { 'FAIL' } else { 'PASS' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "Open rules: $($open.FirewallRuleName -join ', ')"
        }
    }
}

function Test-CIS_4_01_03 {
    <# Ensure SQL Server TDE protector is encrypted with Customer-managed key #>
    param([object[]]$Subscriptions)
    $id = '4.01.03'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv       = $_
            $protector = Get-AzSqlServerTransparentDataEncryptionProtector -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status    = if ($protector.Type -eq 'AzureKeyVault') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "TDE type=$($protector.Type)"
        }
    }
}

function Test-CIS_4_01_04 {
    <# Ensure Azure AD Admin is Configured for SQL Servers #>
    param([object[]]$Subscriptions)
    $id = '4.01.04'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $admin  = Get-AzSqlServerActiveDirectoryAdministrator -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status = if ($admin) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "AAD admin: $($admin.DisplayName)"
        }
    }
}

function Test-CIS_4_01_05 {
    <# Ensure 'Data encryption' is set to 'On' on SQL Databases #>
    param([object[]]$Subscriptions)
    $id = '4.01.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv = $_
            Get-AzSqlDatabase -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName |
                Where-Object { $_.DatabaseName -ne 'master' } | ForEach-Object {
                    $tde    = Get-AzSqlDatabaseTransparentDataEncryption -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -DatabaseName $_.DatabaseName
                    $status = if ($tde.State -eq 'Enabled') { 'PASS' } else { 'FAIL' }
                    Out-CISResult $id $sub.Name "$($srv.ServerName)/$($_.DatabaseName)" $status "TDE=$($tde.State)"
                }
        }
    }
}

function Test-CIS_4_02_02 {
    <# Ensure VA is enabled on SQL Server by setting a Storage Account #>
    param([object[]]$Subscriptions)
    $id = '4.02.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $va     = Get-AzSqlServerVulnerabilityAssessmentSetting -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status = if ($va.StorageAccountName) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "VA storage=$($va.StorageAccountName)"
        }
    }
}

function Test-CIS_4_02_03 {
    <# Ensure VA 'Periodic recurring scans' is 'On' #>
    param([object[]]$Subscriptions)
    $id = '4.02.03'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $va     = Get-AzSqlServerVulnerabilityAssessmentSetting -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status = if ($va.RecurringScansInterval -ne 'None' -and $va.RecurringScansInterval) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "RecurringScan=$($va.RecurringScansInterval)"
        }
    }
}

function Test-CIS_4_02_04 {
    <# Ensure VA 'Send scan reports to' is configured #>
    param([object[]]$Subscriptions)
    $id = '4.02.04'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $va     = Get-AzSqlServerVulnerabilityAssessmentSetting -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status = if ($va.NotificationEmail) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "VA emails=$($va.NotificationEmail -join '; ')"
        }
    }
}

function Test-CIS_4_02_05 {
    <# Ensure VA 'Also send email notifications to admins and subscription owners' is set #>
    param([object[]]$Subscriptions)
    $id = '4.02.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzSqlServer | ForEach-Object {
            $srv    = $_
            $va     = Get-AzSqlServerVulnerabilityAssessmentSetting -ResourceGroupName $srv.ResourceGroupName -ServerName $srv.ServerName -ErrorAction SilentlyContinue
            $status = if ($va.EmailAdmins) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($srv.ResourceGroupName)/$($srv.ServerName)" $status "EmailAdmins=$($va.EmailAdmins)"
        }
    }
}

function Test-CIS_4_04_01 {
    <# Ensure 'Enforce SSL connection' is set to 'Enabled' for MySQL #>
    param([object[]]$Subscriptions)
    $id = '4.04.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzMySqlServer -ErrorAction SilentlyContinue | ForEach-Object {
            $status = if ($_.SslEnforcement -eq 'Enabled') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "SslEnforcement=$($_.SslEnforcement)"
        }
        Get-AzMySqlFlexibleServer -ErrorAction SilentlyContinue | ForEach-Object {
            $p      = Get-AzMySqlFlexibleServerConfiguration -Name require_secure_transport -ResourceGroupName $_.ResourceGroupName -ServerName $_.Name -ErrorAction SilentlyContinue
            $status = if ($p.Value -eq 'ON') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "$($_.Name) (Flex)" $status "require_secure_transport=$($p.Value)"
        }
    }
}

function Test-CIS_4_04_02 {
    <# Ensure TLS Version is set to TLSV1.2 for MySQL Flexible Server #>
    param([object[]]$Subscriptions)
    $id = '4.04.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzMySqlFlexibleServer -ErrorAction SilentlyContinue | ForEach-Object {
            $p      = Get-AzMySqlFlexibleServerConfiguration -Name tls_version -ResourceGroupName $_.ResourceGroupName -ServerName $_.Name -ErrorAction SilentlyContinue
            $status = if ($p.Value -match 'TLSv1\.2') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "tls_version=$($p.Value)"
        }
    }
}

function Get-AzCosmosDBAccountAll {
    <# Helper: returns all Cosmos DB accounts in the current subscription without prompting #>
    Get-AzResourceGroup | ForEach-Object {
        Get-AzCosmosDBAccount -ResourceGroupName $_.ResourceGroupName -ErrorAction SilentlyContinue
    }
}

function Test-CIS_4_05_01 {
    <# Ensure Cosmos DB Firewalls & Networks is limited to Selected Networks #>
    param([object[]]$Subscriptions)
    $id = '4.05.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $accounts = Get-AzCosmosDBAccountAll
        if (-not $accounts) {
            Out-CISResult $id $sub.Name 'N/A' 'UNKNOWN' 'No Cosmos DB accounts found'
            continue
        }
        $accounts | ForEach-Object {
            $status = if ($_.IsVirtualNetworkFilterEnabled -or $_.IpRules.Count -gt 0) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "VNetFilter=$($_.IsVirtualNetworkFilterEnabled) IpRules=$($_.IpRules.Count)"
        }
    }
}

function Test-CIS_4_05_02 {
    <# Ensure Private Endpoints are used for Cosmos DB #>
    param([object[]]$Subscriptions)
    $id = '4.05.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $accounts = Get-AzCosmosDBAccountAll
        if (-not $accounts) {
            Out-CISResult $id $sub.Name 'N/A' 'UNKNOWN' 'No Cosmos DB accounts found'
            continue
        }
        $accounts | ForEach-Object {
            $pe     = $_.PrivateEndpointConnections | Where-Object { $_.PrivateLinkServiceConnectionState.Status -eq 'Approved' }
            $status = if ($pe) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "Approved PEs: $($pe.Count)"
        }
    }
}

#endregion

#region ── 05 – LOGGING AND MONITORING ────────────────────────────────────────

function Test-CIS_5_01_02 {
    <# Ensure Diagnostic Setting captures appropriate categories #>
    param([object[]]$Subscriptions)
    $id   = '5.01.02'
    $cats = @('Administrative','Security','ServiceHealth','Alert','Recommendation','Policy','Autoscale','ResourceHealth')
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $diag = Get-AzDiagnosticSetting -ResourceId "/subscriptions/$($sub.Id)" -ErrorAction SilentlyContinue
        if ($diag) {
            foreach ($d in $diag) {
                $enabled = $d.Logs | Where-Object { $_.Enabled } | Select-Object -ExpandProperty Category
                $missing = $cats | Where-Object { $_ -notin $enabled }
                $status  = if ($missing.Count -eq 0) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $d.Name $status "Missing: $($missing -join ', ')"
            }
        } else {
            Out-CISResult $id $sub.Name $sub.Name 'FAIL' 'No diagnostic settings found'
        }
    }
}

function Test-CIS_5_01_05 {
    <# Ensure logging for Azure Key Vault is Enabled #>
    param([object[]]$Subscriptions)
    $id = '5.01.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | ForEach-Object {
            $kv      = $_
            $diag    = Get-AzDiagnosticSetting -ResourceId $kv.ResourceId -ErrorAction SilentlyContinue
            $enabled = $diag.Logs | Where-Object { $_.Enabled -and $_.Category -eq 'AuditEvent' }
            $status  = if ($enabled) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $kv.VaultName $status "AuditEvent logging=$([bool]$enabled)"
        }
    }
}

function Test-CIS_5_01_06 {
    <# Ensure NSG Flow logs are captured and sent to Log Analytics #>
    param([object[]]$Subscriptions)
    $id = '5.01.06'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzNetworkSecurityGroup | ForEach-Object {
            $nsg = $_
            $nw  = Get-AzNetworkWatcher -ErrorAction SilentlyContinue | Where-Object { $_.Location -eq $nsg.Location }
            if ($nw) {
                $flow   = Get-AzNetworkWatcherFlowLog -NetworkWatcher $nw[0] -TargetResourceId $nsg.Id -ErrorAction SilentlyContinue
                $la     = $flow.FlowAnalyticsConfiguration.NetworkWatcherFlowAnalyticsConfiguration.Enabled
                $status = if ($flow.Enabled -and $la) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $nsg.Name $status "FlowLog=$($flow.Enabled) LAEnabled=$la"
            } else {
                Out-CISResult $id $sub.Name $nsg.Name 'UNKNOWN' 'No Network Watcher in region'
            }
        }
    }
}

function Test-CIS_5_01_07 {
    <# Ensure HTTP logs for Azure App Service are enabled #>
    param([object[]]$Subscriptions)
    $id = '5.01.07'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $app    = $_
            $diag   = Get-AzDiagnosticSetting -ResourceId $app.Id -ErrorAction SilentlyContinue
            $http   = $diag.Logs | Where-Object { $_.Enabled -and $_.Category -eq 'AppServiceHTTPLogs' }
            $status = if ($http) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $app.Name $status "HTTPLogs enabled=$([bool]$http)"
        }
    }
}

function Test-ActivityLogAlert {
    <# Shared helper for all 5.02.xx activity log alert checks #>
    param([string]$ControlID, [string]$OperationName, [object[]]$Subscriptions)
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $alerts = Get-AzActivityLogAlert -ErrorAction SilentlyContinue
        $match  = $alerts | Where-Object {
            $_.Enabled -and ($_.Condition.AllOf | Where-Object { $_.Field -eq 'operationName' -and $_.Equals -eq $OperationName })
        }
        $status = if ($match) { 'PASS' } else { 'FAIL' }
        Out-CISResult $ControlID $sub.Name $sub.Name $status "Alert for '$OperationName': $([bool]$match)"
    }
}

function Test-CIS_5_02_01 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.01' 'Microsoft.Authorization/policyAssignments/write'  $Subscriptions }
function Test-CIS_5_02_02 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.02' 'Microsoft.Authorization/policyAssignments/delete' $Subscriptions }
function Test-CIS_5_02_03 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.03' 'Microsoft.Network/networkSecurityGroups/write'    $Subscriptions }
function Test-CIS_5_02_04 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.04' 'Microsoft.Network/networkSecurityGroups/delete'   $Subscriptions }
function Test-CIS_5_02_05 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.05' 'Microsoft.Security/securitySolutions/write'       $Subscriptions }
function Test-CIS_5_02_06 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.06' 'Microsoft.Security/securitySolutions/delete'      $Subscriptions }
function Test-CIS_5_02_07 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.07' 'Microsoft.Sql/servers/firewallRules/write'        $Subscriptions }
function Test-CIS_5_02_08 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.08' 'Microsoft.Sql/servers/firewallRules/delete'       $Subscriptions }
function Test-CIS_5_02_09 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.09' 'Microsoft.Network/publicIPAddresses/write'        $Subscriptions }
function Test-CIS_5_02_10 { param([object[]]$Subscriptions); Test-ActivityLogAlert '5.02.10' 'Microsoft.Network/publicIPAddresses/delete'       $Subscriptions }

function Test-CIS_5_03_01 {
    <# Ensure Application Insights are Configured #>
    param([object[]]$Subscriptions)
    $id = '5.03.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $aiKey  = $_.SiteConfig.AppSettings | Where-Object { $_.Name -in @('APPINSIGHTS_INSTRUMENTATIONKEY','APPLICATIONINSIGHTS_CONNECTION_STRING') }
            $status = if ($aiKey) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "AppInsights configured=$([bool]$aiKey)"
        }
    }
}

function Test-CIS_5_05 {
    <# Ensure Basic/Consumption SKU not used on monitored artifacts #>
    param([object[]]$Subscriptions)
    $id = '5.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzOperationalInsightsWorkspace -ErrorAction SilentlyContinue | ForEach-Object {
            $sku    = $_.Sku.Name
            $status = if ($sku -notin @('Free','Standalone')) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "LAW:$($_.Name)" $status "SKU=$sku"
        }
        Get-AzEventHubNamespace -ErrorAction SilentlyContinue | ForEach-Object {
            $tier   = $_.Sku.Tier
            $status = if ($tier -ne 'Basic') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name "EH:$($_.Name)" $status "Tier=$tier"
        }
    }
}

#endregion

#region ── 06 – NETWORKING ─────────────────────────────────────────────────────

function Test-NSGPortAccess {
    param([string]$ControlID, $Port, [string]$Protocol, [string]$Description, [object[]]$Subscriptions)
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzNetworkSecurityGroup | ForEach-Object {
            $nsg       = $_
            $openRules = $nsg.SecurityRules | Where-Object {
                $_.Direction -eq 'Inbound' -and $_.Access -eq 'Allow' -and
                ($_.SourceAddressPrefix -in @('*','Internet','0.0.0.0/0')) -and
                ($_.Protocol -eq $Protocol -or $_.Protocol -eq '*') -and
                ($Port -eq '*' -or $_.DestinationPortRange -eq '*' -or
                 ($_.DestinationPortRange -in ($Port -as [string[]])) -or
                 ($_.DestinationPortRanges -in ($Port -as [string[]])))
            }
            $status = if ($openRules) { 'FAIL' } else { 'PASS' }
            Out-CISResult $ControlID $sub.Name $nsg.Name $status "$Description open rules: $($openRules.Name -join ', ')"
        }
    }
}

function Test-CIS_6_01 { param([object[]]$Subscriptions); Test-NSGPortAccess '6.01' 3389 'TCP' 'RDP'      $Subscriptions }
function Test-CIS_6_02 { param([object[]]$Subscriptions); Test-NSGPortAccess '6.02' 22   'TCP' 'SSH'      $Subscriptions }
function Test-CIS_6_03 { param([object[]]$Subscriptions); Test-NSGPortAccess '6.03' '*'  'UDP' 'UDP'      $Subscriptions }
function Test-CIS_6_04 { param([object[]]$Subscriptions); Test-NSGPortAccess '6.04' @(80,443) 'TCP' 'HTTP/S' $Subscriptions }

function Test-CIS_6_05 {
    <# Ensure NSG Flow Log retention is greater than 90 days #>
    param([object[]]$Subscriptions)
    $id = '6.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzNetworkSecurityGroup | ForEach-Object {
            $nsg = $_
            $nw  = Get-AzNetworkWatcher -ErrorAction SilentlyContinue | Where-Object { $_.Location -eq $nsg.Location }
            if ($nw) {
                $flow   = Get-AzNetworkWatcherFlowLog -NetworkWatcher $nw[0] -TargetResourceId $nsg.Id -ErrorAction SilentlyContinue
                $ret    = $flow.RetentionPolicy.Days
                $status = if ($flow.RetentionPolicy.Enabled -and $ret -gt 90) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name $nsg.Name $status "Retention=$ret days Enabled=$($flow.RetentionPolicy.Enabled)"
            } else {
                Out-CISResult $id $sub.Name $nsg.Name 'UNKNOWN' 'No Network Watcher in region'
            }
        }
    }
}

#endregion

#region ── 07 – VIRTUAL MACHINES ──────────────────────────────────────────────

function Test-CIS_7_01 {
    <# Ensure an Azure Bastion Host Exists #>
    param([object[]]$Subscriptions)
    $id = '7.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        $bastions = Get-AzBastion -ErrorAction SilentlyContinue
        $status   = if ($bastions) { 'PASS' } else { 'FAIL' }
        Out-CISResult $id $sub.Name $sub.Name $status "Bastion hosts: $($bastions.Count)"
    }
}

function Test-CIS_7_03 {
    <# Ensure OS and Data disks are encrypted with Customer Managed Key (CMK) #>
    param([object[]]$Subscriptions)
    $id = '7.03'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzDisk | ForEach-Object {
            $enc    = $_.Encryption.Type
            $status = if ($enc -eq 'EncryptionAtRestWithCustomerKey') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "EncryptionType=$enc State=$($_.DiskState)"
        }
    }
}

function Test-CIS_7_04 {
    <# Ensure Unattached disks are encrypted with CMK #>
    param([object[]]$Subscriptions)
    $id = '7.04'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzDisk | Where-Object { $_.DiskState -eq 'Unattached' } | ForEach-Object {
            $enc    = $_.Encryption.Type
            $status = if ($enc -eq 'EncryptionAtRestWithCustomerKey') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "EncryptionType=$enc"
        }
    }
}

#endregion

#region ── 08 – KEY VAULT ─────────────────────────────────────────────────────

function Test-CIS_8_02 {
    <# Ensure Expiration Date is set for all Keys in Non-RBAC Key Vaults #>
    param([object[]]$Subscriptions)
    $id = '8.02'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | Where-Object { -not $_.EnableRbacAuthorization } | ForEach-Object {
            $vn = $_.VaultName
            Get-AzKeyVaultKey -VaultName $vn -ErrorAction SilentlyContinue | ForEach-Object {
                $status = if ($_.Attributes.Expires) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name "$vn/$($_.Name)" $status "Expires=$($_.Attributes.Expires)"
            }
        }
    }
}

function Test-CIS_8_03 {
    <# Ensure Expiration Date is set for all Secrets in RBAC Key Vaults #>
    param([object[]]$Subscriptions)
    $id = '8.03'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | Where-Object { $_.EnableRbacAuthorization } | ForEach-Object {
            $vn = $_.VaultName
            Get-AzKeyVaultSecret -VaultName $vn -ErrorAction SilentlyContinue | ForEach-Object {
                $status = if ($_.Attributes.Expires) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name "$vn/$($_.Name)" $status "Expires=$($_.Attributes.Expires)"
            }
        }
    }
}

function Test-CIS_8_04 {
    <# Ensure Expiration Date is set for all Secrets in Non-RBAC Key Vaults #>
    param([object[]]$Subscriptions)
    $id = '8.04'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | Where-Object { -not $_.EnableRbacAuthorization } | ForEach-Object {
            $vn = $_.VaultName
            Get-AzKeyVaultSecret -VaultName $vn -ErrorAction SilentlyContinue | ForEach-Object {
                $status = if ($_.Attributes.Expires) { 'PASS' } else { 'FAIL' }
                Out-CISResult $id $sub.Name "$vn/$($_.Name)" $status "Expires=$($_.Attributes.Expires)"
            }
        }
    }
}

function Test-CIS_8_05 {
    <# Ensure the Key Vault is Recoverable #>
    param([object[]]$Subscriptions)
    $id = '8.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | ForEach-Object {
            $status = if ($_.EnableSoftDelete -and $_.EnablePurgeProtection) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.VaultName $status "SoftDelete=$($_.EnableSoftDelete) PurgeProtection=$($_.EnablePurgeProtection)"
        }
    }
}

function Test-CIS_8_06 {
    <# Ensure Role Based Access Control for Azure Key Vault #>
    param([object[]]$Subscriptions)
    $id = '8.06'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | ForEach-Object {
            $status = if ($_.EnableRbacAuthorization) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.VaultName $status "RBACEnabled=$($_.EnableRbacAuthorization)"
        }
    }
}

function Test-CIS_8_07 {
    <# Ensure Private Endpoints are Used for Azure Key Vault #>
    param([object[]]$Subscriptions)
    $id = '8.07'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzKeyVault | ForEach-Object {
            $pe     = $_.PrivateEndpointConnections | Where-Object { $_.PrivateLinkServiceConnectionState.Status -eq 'Approved' }
            $status = if ($pe) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.VaultName $status "Approved PEs: $($pe.Count)"
        }
    }
}

#endregion

#region ── 09 – APP SERVICES ───────────────────────────────────────────────────

function Test-CIS_9_01 {
    <# Ensure App Service Authentication is set up #>
    param([object[]]$Subscriptions)
    $id = '9.01'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $app    = $_
            $auth   = Get-AzWebAppAuthSettings -ResourceGroupName $app.ResourceGroup -Name $app.Name -ErrorAction SilentlyContinue
            $status = if ($auth.Enabled) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $app.Name $status "AuthEnabled=$($auth.Enabled)"
        }
    }
}

function Test-CIS_9_04 {
    <# Ensure 'Client Certificates' is 'On' #>
    param([object[]]$Subscriptions)
    $id = '9.04'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $status = if ($_.ClientCertEnabled) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "ClientCertEnabled=$($_.ClientCertEnabled)"
        }
    }
}

function Test-CIS_9_05 {
    <# Ensure Register with Azure Active Directory is enabled on App Service #>
    param([object[]]$Subscriptions)
    $id = '9.05'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $status = if ($_.Identity -and $_.Identity.Type -ne 'None') { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $_.Name $status "ManagedIdentity=$($_.Identity.Type)"
        }
    }
}

function Test-CIS_9_09 {
    <# Ensure HTTP Version is the Latest #>
    param([object[]]$Subscriptions)
    $id = '9.09'
    foreach ($sub in $Subscriptions) {
        Set-AzContext -SubscriptionId $sub.Id | Out-Null
        Get-AzWebApp | ForEach-Object {
            $app    = $_
            $cfg    = Get-AzWebAppConfiguration -ResourceGroupName $app.ResourceGroup -Name $app.Name -ErrorAction SilentlyContinue
            $status = if ($cfg.Http20Enabled) { 'PASS' } else { 'FAIL' }
            Out-CISResult $id $sub.Name $app.Name $status "Http20Enabled=$($cfg.Http20Enabled)"
        }
    }
}

#endregion

#region ── MASTER RUNNER ──────────────────────────────────────────────────────

function Invoke-AllCISChecks {
    <#
    .SYNOPSIS
        Runs all 61 CIS v2.0 checks and exports results to CSV.
    .PARAMETER SubscriptionId
        Optional. Pass a specific subscription ID to skip the interactive menu.
    .PARAMETER OutputPath
        Path for CSV output. Default: .\CIS_Results_<timestamp>.csv
    .EXAMPLE
        Invoke-AllCISChecks
    .EXAMPLE
        Invoke-AllCISChecks -SubscriptionId 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'
    #>
    param(
        [string]$SubscriptionId,
        [string]$OutputPath = ".\CIS_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    )

    Connect-AzGov
    $targetSubs = Select-AzGovSubscription -SubscriptionId $SubscriptionId
    if (-not $targetSubs) { return }

    Write-Host "`nRunning CIS checks against: $($targetSubs.Name -join ', ')" -ForegroundColor Cyan

    $checks = @(
        { Test-CIS_2_01_03 $targetSubs }, { Test-CIS_2_01_14 $targetSubs },
        { Test-CIS_2_01_15 $targetSubs }, { Test-CIS_2_01_19 $targetSubs },
        { Test-CIS_3_01 $targetSubs },    { Test-CIS_3_02 $targetSubs },
        { Test-CIS_3_05 $targetSubs },    { Test-CIS_3_08 $targetSubs },
        { Test-CIS_3_09 $targetSubs },    { Test-CIS_3_10 $targetSubs },
        { Test-CIS_3_11 $targetSubs },    { Test-CIS_3_12 $targetSubs },
        { Test-CIS_3_13 $targetSubs },    { Test-CIS_3_15 $targetSubs },
        { Test-CIS_4_01_01 $targetSubs }, { Test-CIS_4_01_02 $targetSubs },
        { Test-CIS_4_01_03 $targetSubs }, { Test-CIS_4_01_04 $targetSubs },
        { Test-CIS_4_01_05 $targetSubs }, { Test-CIS_4_02_02 $targetSubs },
        { Test-CIS_4_02_03 $targetSubs }, { Test-CIS_4_02_04 $targetSubs },
        { Test-CIS_4_02_05 $targetSubs }, { Test-CIS_4_04_01 $targetSubs },
        { Test-CIS_4_04_02 $targetSubs }, { Test-CIS_4_05_01 $targetSubs },
        { Test-CIS_4_05_02 $targetSubs }, { Test-CIS_5_01_02 $targetSubs },
        { Test-CIS_5_01_05 $targetSubs }, { Test-CIS_5_01_06 $targetSubs },
        { Test-CIS_5_01_07 $targetSubs }, { Test-CIS_5_02_01 $targetSubs },
        { Test-CIS_5_02_02 $targetSubs }, { Test-CIS_5_02_03 $targetSubs },
        { Test-CIS_5_02_04 $targetSubs }, { Test-CIS_5_02_05 $targetSubs },
        { Test-CIS_5_02_06 $targetSubs }, { Test-CIS_5_02_07 $targetSubs },
        { Test-CIS_5_02_08 $targetSubs }, { Test-CIS_5_02_09 $targetSubs },
        { Test-CIS_5_02_10 $targetSubs }, { Test-CIS_5_03_01 $targetSubs },
        { Test-CIS_5_05 $targetSubs },    { Test-CIS_6_01 $targetSubs },
        { Test-CIS_6_02 $targetSubs },    { Test-CIS_6_03 $targetSubs },
        { Test-CIS_6_04 $targetSubs },    { Test-CIS_6_05 $targetSubs },
        { Test-CIS_7_01 $targetSubs },    { Test-CIS_7_03 $targetSubs },
        { Test-CIS_7_04 $targetSubs },    { Test-CIS_8_02 $targetSubs },
        { Test-CIS_8_03 $targetSubs },    { Test-CIS_8_04 $targetSubs },
        { Test-CIS_8_05 $targetSubs },    { Test-CIS_8_06 $targetSubs },
        { Test-CIS_8_07 $targetSubs },    { Test-CIS_9_01 $targetSubs },
        { Test-CIS_9_04 $targetSubs },    { Test-CIS_9_05 $targetSubs },
        { Test-CIS_9_09 $targetSubs }
    )

    $results = @()
    $i = 0
    foreach ($check in $checks) {
        $i++
        Write-Progress -Activity "Running CIS Checks" -Status "Check $i of $($checks.Count)" -PercentComplete (($i / $checks.Count) * 100)
        $results += & $check
    }
    Write-Progress -Activity "Running CIS Checks" -Completed

    $results | Export-Csv -Path $OutputPath -NoTypeInformation
    Write-Host "`nResults exported to: $OutputPath" -ForegroundColor Cyan
    Write-Host "`nSummary:" -ForegroundColor Cyan
    $results | Group-Object Status | Sort-Object Name | ForEach-Object {
        $color = switch ($_.Name) { 'PASS' {'Green'} 'FAIL' {'Red'} default {'Yellow'} }
        Write-Host "  $($_.Name): $($_.Count)" -ForegroundColor $color
    }
    return $results
}

#endregion