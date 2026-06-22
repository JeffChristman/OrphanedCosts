<#
.SYNOPSIS
    Inventories public outbound IP addresses for all Azure App Services across
    all accessible subscriptions.

.DESCRIPTION
    Iterates every subscription available to the authenticated Azure account,
    retrieves all App Service web apps in each subscription, and extracts their
    outbound IP addresses. Private (RFC 1918) addresses are filtered out, and
    only public IPs are retained. Results are displayed in the console and
    exported to a CSV file for audit or firewall allowlist purposes.

.NOTES
    Name:       Get-AppServicePublicIPs.ps1
    Author:     <Author>
    Created:    <Date>
    Version:    1.0

    Prerequisites:
        - Az PowerShell module must be installed:
              Install-Module -Name Az -Scope CurrentUser
        - Authenticate before running:
              Connect-AzAccount
          The Connect-AzAccount line in this script is commented out by default
          so it can be run in sessions where authentication already exists.

    Private IP Ranges Filtered (RFC 1918):
        10.0.0.0/8
        172.16.0.0/12
        192.168.0.0/16

    Output:
        Console   — Formatted table of all public IPs found.
        CSV file  — ./AppServices_PublicVIPs.csv (relative to working directory).

    Important:
        OutboundIpAddresses reflects the shared pool of IPs the App Service
        *may* use for outbound traffic — not a single dedicated VIP. For a
        complete list including IPs used during scale-out, see the
        PossibleOutboundIpAddresses property instead.

    Permissions Required:
        Reader role (or equivalent) on each subscription being scanned.
#>


# ---------------------------------------------------------------------------
# Prerequisites
# ---------------------------------------------------------------------------

# Uncomment the line below if not already authenticated in this session.
#Connect-AzAccount


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Output CSV path. Defaults to the current working directory.
$csvOutputPath = "./AppServices_PublicVIPs.csv"

# RFC 1918 private IP ranges used to filter out non-public addresses.
# Matches: 10.x.x.x | 192.168.x.x | 172.16.x.x–172.31.x.x
$privateIpPattern = '^(10\.|192\.168\.|172\.(1[6-9]|2[0-9]|3[0-1])\.)'


# ---------------------------------------------------------------------------
# Main Processing
# ---------------------------------------------------------------------------

# Retrieve all subscriptions accessible to the authenticated account.
$subscriptions = Get-AzSubscription

# Array to accumulate result objects across all subscriptions and apps.
$results = @()

foreach ($sub in $subscriptions) {

    Write-Host "Processing subscription: $($sub.Name)" -ForegroundColor Cyan

    # Switch the Az context to the current subscription so all subsequent
    # cmdlets operate within it.
    Set-AzContext -SubscriptionId $sub.Id

    # Retrieve all App Service web apps in this subscription.
    $apps = Get-AzWebApp

    foreach ($app in $apps) {

        # Fetch the full app configuration. Get-AzWebApp without -Name returns
        # a summary object; the full object is required to access IP properties.
        $config = Get-AzWebApp -ResourceGroupName $app.ResourceGroup -Name $app.Name

        # OutboundIpAddresses is a comma-separated string of IPs in the shared
        # outbound pool. Split into an array for individual evaluation.
        $vipList = $config.OutboundIpAddresses -split ','

        foreach ($ip in $vipList) {

            # Skip RFC 1918 private addresses — only public IPs are relevant.
            if ($ip -match $privateIpPattern) {
                continue
            }

            $results += [PSCustomObject]@{
                Subscription    = $sub.Name
                AppName         = $app.Name
                ResourceGroup   = $app.ResourceGroup
                Location        = $app.Location
                PublicIPAddress = $ip
            }
        }
    }
}


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

# Display results in the console as a formatted table.
$results | Format-Table -AutoSize

# Export results to CSV for audit, reporting, or firewall allowlist use.
$results | Export-Csv -Path $csvOutputPath -NoTypeInformation

Write-Host "`nExport complete: $csvOutputPath" -ForegroundColor Green