# This script will connect to a CRM to identify and remove all 
#    sdkmessageprocessingsteps
#    plugintypes
#    pluginassemblys
# associated with any assemblies named with the given name

[CmdletBinding()]
Param(
    [string] $EnvironmentUrl,
    [Guid] $OAuthClientId,
    [string] $ClientSecret,
    [string] $AssemblyName 
)

Write-Host "Input params:"
Write-Host "  EnvironmentUrl: $EnvironmentUrl"
Write-Host "  OAuthClientId: $OAuthClientId"
Write-Host "  ClientSecret: $ClientSecret"
Write-Host "  AssemblyName: $AssemblyName"
Write-Host ""

# Install modules
Write-Host 'Installing required modules';
Install-Module  Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force;

#set up the connection to be used by the below commands
$connection = Connect-CrmOnline -ServerUrl $EnvironmentUrl -OAuthClientId $OAuthClientId -ClientSecret $ClientSecret -OAuthRedirectUri “https://tempuri.org”;

Write-Host "Connection established";
Write-Host "";

#get the steps
$stepsFetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='sdkmessageprocessingstep'><link-entity name='plugintype' from='plugintypeid' to='plugintypeid' visible='false' link-type='inner' alias='step'><order attribute='typename' descending='false' /><filter type='and'><condition attribute='assemblyname' operator='eq' value='$AssemblyName' /></filter></link-entity></entity></fetch>";
$stepsRecords = (Get-CrmRecordsByFetch $stepsFetch).CrmRecords;
$stepsCount = $stepsRecords.Count;
Write-Host "Retrieved $stepsCount step(s)";

#remove them
$stepsRecords | % {
  Remove-CrmRecord -CrmRecord $_
}    

Write-Host "Step(s) removed";
Write-Host "";

#get the types
$typesFetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='plugintype'><order attribute='typename' descending='false' /><filter type='and'><condition attribute='assemblyname' operator='eq' value='$AssemblyName' /></filter></entity></fetch>";
$typesRecords = (Get-CrmRecordsByFetch $typesFetch).CrmRecords;
$typesCount = $typesRecords.Count;
Write-Host "Retrieved $typesCount type(s)";

#remove them
$typesRecords | % {
  Remove-CrmRecord -CrmRecord $_
}    

Write-Host "Type(s) removed";
Write-Host "";

#get the assemblies
$assemblyFetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='pluginassembly'><filter type='and'><condition attribute='name' operator='eq' value='$AssemblyName' /></filter></entity></fetch>";
$assemblyRecords = (Get-CrmRecordsByFetch $assemblyFetch).CrmRecords;
$assemblyCount = $assemblyRecords.Count;
Write-Host "Retrieved $assemblyCount assembly(ies)";

#remove them
$assemblyRecords | % {
  Remove-CrmRecord -CrmRecord $_
}

Write-Host "Assembly(ies) removed";
Write-Host "";

Write-Host "Process complete";
