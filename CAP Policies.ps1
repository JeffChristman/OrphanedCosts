# Check for required module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

# Import module
Import-Module Microsoft.Graph

# Connect to Microsoft Graph with required scopes
try {
    Connect-MgGraph -Scopes "Policy.Read.All", "Directory.Read.All"
    Write-Host "✅ Connected to Microsoft Graph"
}
catch {
    Write-Error "❌ Could not connect. Please verify your role (Global Reader, Security Reader) and Graph API permissions (Policy.Read.All)."
    return
}

# Try to get Conditional Access policies
try {
    $policies = Get-MgConditionalAccessPolicy
    if (-not $policies) {
        Write-Warning "⚠️ No Conditional Access policies found, or insufficient permissions."
        return
    }
}
catch {
    Write-Error "❌ Failed to retrieve Conditional Access policies. You may lack permission."
    return
}

# Parse and store key information
$results = foreach ($policy in $policies) {
    [PSCustomObject]@{
        DisplayName       = $policy.DisplayName
        State             = $policy.State
        Users             = ($policy.Conditions.Users.IncludeUsers -join ", ")
        ExcludeUsers      = ($policy.Conditions.Users.ExcludeUsers -join ", ")
        Applications      = ($policy.Conditions.Applications.IncludeApplications -join ", ")
        Platforms         = ($policy.Conditions.Platforms.IncludePlatforms -join ", ")
        Locations         = ($policy.Conditions.Locations.IncludeLocations -join ", ")
        SignInRiskLevels  = ($policy.Conditions.SignInRiskLevels -join ", ")
        GrantControls     = ($policy.GrantControls.BuiltInControls -join ", ")
        SessionControls   = ($policy.SessionControls | ConvertTo-Json -Depth 3)
        Created           = $policy.CreatedDateTime
        Modified          = $policy.ModifiedDateTime
        ID                = $policy.Id
    }
}

# Export to CSV
$results | Export-Csv -Path ".\VA_ConditionalAccess_PolicyExport.csv" -NoTypeInformation -Encoding UTF8

Write-Host "✅ Export complete: VA_ConditionalAccess_PolicyExport.csv"

