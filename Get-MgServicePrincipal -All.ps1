Get-MgServicePrincipal -All

# Check API permissions (delegated + application)
Get-MgServicePrincipal -All | ForEach-Object {
    $sp = $_
    $perms = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id
    [PSCustomObject]@{
        DisplayName = $sp.DisplayName
        AppId       = $sp.AppId
        ObjectId    = $sp.Id
        Permissions = ($perms | Select-Object -ExpandProperty ResourceDisplayName -Unique) -join ", "
    }
}
