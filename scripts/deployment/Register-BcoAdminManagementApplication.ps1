[CmdletBinding()]
Param(
    [Guid] $TenantId
)

$appId = "a86b9632-42bf-4dfe-83c8-bbc95145504b"

try
{
    Write-Host 'Registering an admin management application';

    # Login interactively with a tenant administrator for Power Platform
    Add-PowerAppsAccount -Endpoint prod -TenantID $TenantId 

    # Register a new application, this gives the SPN / client application same permissions as a tenant admin
    New-PowerAppManagementApp -ApplicationId $appId

    Write-Host 'SUCCESS!';
    Write-Host 
}
catch
{
    $Error[0]
    Write-Host 'Error registering application'
    exit(1);
}
