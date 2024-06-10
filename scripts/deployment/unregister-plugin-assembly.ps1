# This script will connect to a CRM to identify and remove all 
#    sdkmessageprocessingsteps
#    plugintypes
#    pluginassemblys
# associated with any assemblies named with the given name
# only IF the target environment's version does not match the base environment

[CmdletBinding()]
Param(
    [string] $BaseEnvironmentUrl,
    [string] $EnvironmentUrl,
    [Guid] $OAuthClientId,
    [string] $ClientSecret,
    [string] $AssemblyName 
)

Write-Host "Input params:"
Write-Host "  BaseEnvironmentUrl: $BaseEnvironmentUrl"
Write-Host "  EnvironmentUrl: $EnvironmentUrl"
Write-Host "  OAuthClientId: $OAuthClientId"
Write-Host "  ClientSecret: $ClientSecret"
Write-Host "  AssemblyName: $AssemblyName"
Write-Host ""

# Install modules
Write-Host 'Installing required modules';
Install-Module  Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force;

#set up the base connection
$baseConnection = Connect-CrmOnline -ServerUrl $BaseEnvironmentUrl -OAuthClientId $OAuthClientId -ClientSecret $ClientSecret -OAuthRedirectUri �https://tempuri.org�;
Write-Host "Base connection established";

$assemblyFetch = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'><entity name='pluginassembly'><filter type='and'><condition attribute='name' operator='eq' value='$AssemblyName' /></filter></entity></fetch>";

#get base assemblies
$baseAssemblyRecords = (Get-CrmRecordsByFetch $assemblyFetch).CrmRecords;
$baseAssemblyCount = $baseAssemblyRecords.Count;
Write-Host "Retrieved $baseAssemblyCount assembly(ies)";
#get the base version
$baseAssemblyVersion = "";
$baseAssemblyRecords | % {
  $baseAssemblyVersion =  (Get-CrmRecord -EntityLogicalName pluginassembly -Fields version -Id $_.pluginassemblyid -conn $baseConnection).version;
}

Write-Host "";

#set up the target connection
$connection = Connect-CrmOnline -ServerUrl $EnvironmentUrl -OAuthClientId $OAuthClientId -ClientSecret $ClientSecret -OAuthRedirectUri �https://tempuri.org�;
Write-Host "Connection established";

#get the assemblies
$assemblyRecords = (Get-CrmRecordsByFetch $assemblyFetch).CrmRecords;
$assemblyCount = $assemblyRecords.Count;
Write-Host "Retrieved $assemblyCount assembly(ies)";
#get the base version
$assemblyVersion = "";
$assemblyRecords | % {
  $assemblyVersion =  (Get-CrmRecord -EntityLogicalName pluginassembly -Fields version -Id $_.pluginassemblyid -conn $connection).version;
}

Write-Host "";

Write-Host "Base version  : $baseAssemblyVersion";
Write-Host "Target version: $assemblyVersion";

#if the base and target version are the same, don't remove
if ($assemblyVersion -eq $baseAssemblyVersion) {
  Write-Host "Versions are the same.  Leave the plugins alone.";
  Write-Host "";
}
#if they aren't the same, remove existing
else {
  Write-Host "Versions are not the same.  Remove existing assembly, types, and steps.";
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
}

Write-Host "Process complete";
