[CmdletBinding()]
Param(
    # gather permission requests but don't create any AppId nor ServicePrincipal
    [switch] $DryRun = $false,
    # other possible Azure environments, see: https://docs.microsoft.com/en-us/powershell/module/azuread/connect-azuread?view=azureadps-2.0#parameters
    [string] $AzureEnvironment = "AzureCloud",

    [ValidateSet(
        "UnitedStates",
        "Preview(UnitedStates)",
        "Europe",
        "EMEA",
        "Asia",
        "Australia",
        "Japan",
        "SouthAmerica",
        "India",
        "Canada",
        "UnitedKingdom",
        "France"
    )]
    [string] $TenantLocation = "UnitedStates"
)

function ensureModules {
    $dependencies = @(
        # the more general and modern "Az" a "AzureRM" do not have proper support to manage permissions
        @{ Name = "AzureAD"; Version = [Version]"2.0.2.137"; "InstallWith" = "Install-Module -Name AzureAD -AllowClobber -Scope CurrentUser" }
    )
    $missingDependencies = $false
    $dependencies | ForEach-Object -Process {
        $moduleName = $_.Name
        $deps = (Get-Module -ListAvailable -Name $moduleName `
            | Sort-Object -Descending -Property Version)
        if ($deps -eq $null) {
            Write-Host @"
ERROR: Required module not installed; install from PowerShell prompt with:
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

function connectAAD {
    Write-Host @"

Connecting to AzureAD: Please log in, using your Dynamics365 / Power Platform tenant ADMIN credentials:

"@
    try {
        Connect-AzureAD -AzureEnvironmentName $AzureEnvironment -ErrorAction Stop | Out-Null
    }
    catch {
        throw "Failed to login: $($_.Exception.Message)"
    }
    return Get-AzureADCurrentSessionInfo
}

function reconnectAAD {
    # for tenantID, see DirectoryID here: https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/Overview
    try {
        $session = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
        if ($session.Environment.Name -ne $AzureEnvironment) {
            Disconnect-AzureAd
            $session = connectAAD
        }
    }
    catch [Microsoft.Open.Azure.AD.CommonLibrary.AadNeedAuthenticationException] {
        $session = connectAAD
    }
    $tenantId = $session.TenantId

    Write-Host @"
Connected to AAD tenant: $($session.TenantDomain) ($($tenantId))

"@
    return $tenantId
}

function getAppConsentUri($tenantDomain) {
    "https://login.microsoftonline.com/$tenantDomain/oauth2/authorize?client_id=63b6e8df-e035-446b-938e-5173ebe4ac69&response_type=code&redirect_uri=https://akoyago.com&nonce=doesntmatter&resource=https://graph.microsoft.com&prompt=admin_consent"
}




$hasAdmin = checkIsElevated
if (!$hasAdmin) {
    throw "This action requires administrator privileges."
}

if ($PSVersionTable.PSEdition -ne "Desktop") {
    throw "This script must be run on PowerShell Desktop/Windows; the AzureAD module is not supported for PowershellCore yet!"
}

ensureModules

$ErrorActionPreference = "Stop"

$tenantId = reconnectAAD
$session = Get-AzureADCurrentSessionInfo -ErrorAction SilentlyContinue
$spnDisplayName = "BCO akoyaGO GOmanager User Connector"

if (!$DryRun) {
    $spn = New-AzureADServicePrincipal -AccountEnabled $true -AppId "63b6e8df-e035-446b-938e-5173ebe4ac69" -AppRoleAssignmentRequired $true -DisplayName "$spnDisplayName" -Tags {WindowsAzureActiveDirectoryIntegratedApp}
    $spnId = $spn.ObjectId
    Write-Host "Created SPN '$spnDisplayName' with objectId: $spnId"
}
else {
    Write-Host "Skipping SPN creation because DryRun is 'true'"
}

Write-Host @"

Copy and paste the following URL in a browser to grant consent.

"@

Write-Host $(getAppConsentUri $session.TenantDomain)

Write-Host @"

Done.

"@
