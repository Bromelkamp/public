[CmdletBinding()]
Param(
    # gather permission requests but don't create any AppId nor ServicePrincipal
    [switch] $DryRun = $false
)

$appId = "a86b9632-42bf-4dfe-83c8-bbc95145504b"


function ensureModules {
    $dependencies = @(
        # the more general and modern "Az" a "AzureRM" do not have proper support to manage permissions
        @{ Name = "Microsoft.PowerApps.Administration.PowerShell"; Version = [Version]"2.0.154"; "InstallWith" = "Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Force" },
        @{ Name = "Microsoft.PowerApps.PowerShell"; Version = [Version]"1.0.26"; "InstallWith" = "Install-Module -Name Microsoft.PowerApps.PowerShell -AllowClobber -Force" }
    )
    $missingDependencies = $false
    $dependencies | ForEach-Object -Process {
        $moduleName = $_.Name
        $deps = (Get-Module -ListAvailable -Name $moduleName `
            | Sort-Object -Descending -Property Version)
        if ($null -eq $deps) {
            Write-Host @"

ERROR: Required module not installed; install from an elevated PowerShell prompt with:
>>  $($_.InstallWith) -MinimumVersion $($_.Version)

"@
            $missingDependencies = $true
            return
        }
        $dep = $deps[0]
        if ($dep.Version -lt $_.Version) {
            Write-Host @"

ERROR: Required module installed but does not meet minimal required version:
       found: $($dep.Version), required: >= $($_.Version); to fix, please run:
>>  Update-Module $($_.Name) -Scope CurrentUser -RequiredVersion $($_.Version)

"@
            $missingDependencies = $true
            return
        }
        Import-Module $moduleName -MinimumVersion $_.Version
    }
    if ($missingDependencies) {
        throw "Missing required dependencies!"
    }
}

function checkIsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { 
        Write-Output $true 
    }
    else { 
        Write-Output $false 
    }
}

function Test-IsGuid
{
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        [string]$StringGuid
    )
 
   $ObjectGuid = [System.Guid]::empty
   return [System.Guid]::TryParse($StringGuid,[System.Management.Automation.PSReference]$ObjectGuid) # Returns True if successfully parsed
}

function connectPowerApps {
    Write-Host @"

    Connecting to Power Platform...
Please log in, using your Dynamics 365 / Power Platform tenant ADMIN credentials:
"@
    try {
        Add-PowerAppsAccount -Endpoint "prod" -TenantID $tenantId -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to login: $($_.Exception.Message)"
    }
}


########################################################################

# Make sure current user has admin rights

$hasAdmin = checkIsElevated
if (!$hasAdmin) {
    throw "This action requires administrator privileges."
}

if ($PSVersionTable.PSEdition -ne "Desktop") {
    throw "This script must be run on PowerShell Desktop/Windows; the AzureAD module is not supported for PowershellCore yet!"
}

# Install required modules

ensureModules

# Get user inputs

$tenantId = Read-Host -Prompt "Enter the tenant ID (GUID)"
$isTenantIdGuid = Test-IsGuid $tenantId
if (!$isTenantIdGuid) {
    throw "Invalid tenant ID"
}

$crmEnvironmentId = Read-Host -Prompt "Enter the CRM Environment ID (GUID)"
$isCrmEnvironmentIdGuid = Test-IsGuid $crmEnvironmentId
if (!$isCrmEnvironmentIdGuid) {
    throw "Invalid CRM Environment ID"
}      

# Connect to PowerApps

connectPowerApps

Test-PowerAppsAccount

$ErrorActionPreference = "Stop"


# Register "BCO akoyaGO Integration" SPN in tenant

Write-Host @"

Registering "BCO akoyaGO Integration" ($appId) SPN in tenant

"@

$registerReport = New-PowerAppManagementApp -ApplicationId $appId
Write-Host $registerReport

Write-Host @"

Complete.
"@

Write-Host @"


Report start ----------------------------------------------------------------

 Tenant ID: $tenantId
 CRM Environment ID: $crmEnvironmentId

"@

# Get and report selected environment

Write-Host @"


## Environment details:

"@

$environmentReport = Get-AdminPowerAppEnvironment -EnvironmentName $crmEnvironmentId
$environmentJson = ConvertTo-Json @($environmentReport)
#$environmentJson | Format-Table -AutoSize
Write-Host $environmentJson

# Get and report Dataverse connections

Write-Host @"


Microsoft Dataverse connections:

"@

$dataverseReport = Get-AdminPowerAppConnection -EnvironmentName $crmEnvironmentId -ConnectorName shared_commondataserviceforapps
$dataverseJson = ConvertTo-Json @($dataverseReport)
Write-Host $dataverseJson

# Get and report Dataverse legacy connections

Write-Host @"


Microsoft Dataverse (legacy) connections:

"@

$dataverseLegacyReport = Get-AdminPowerAppConnection -EnvironmentName $crmEnvironmentId -ConnectorName shared_commondataservice
$dataverseLegacyJson = ConvertTo-Json @($dataverseLegacyReport)
Write-Host $dataverseLegacyJson

# Get and report SharePoint connections

Write-Host @"


SharePoint connections:

"@

$sharepointReport = Get-AdminPowerAppConnection -EnvironmentName $crmEnvironmentId -ConnectorName shared_sharepointonline
$sharepointJson = ConvertTo-Json @($sharepointReport)
Write-Host $sharepointJson

Write-Host @"


Report end ----------------------------------------------------------------
"@

Write-Host @"

Done.

"@

