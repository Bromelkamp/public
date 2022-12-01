# This script required an AAD App that has been registered using 'New-PowerAppManagementApp'.
# https://learn.microsoft.com/en-us/power-platform/admin/powershell-create-service-principal#registering-an-admin-management-application

# Requires: Akoyanet flow connections references to be mapped to tenant's connections

[CmdletBinding()]
Param(
    [string] $EnvironmentUrl,
    [Guid] $TenantId,
    [Guid] $ApplicationId,
    [string] $ClientSecret 
)

$BuildToolsConnectionString = "AuthType=ClientSecret;url=$EnvironmentUrl;ClientId=$ApplicationId;ClientSecret=$ClientSecret";
$BuildToolsSolutionName = 'Akoyanet';

Write-Host "Input params:"
Write-Host "  EnvironmentUrl: $EnvironmentUrl"
Write-Host "  TenantId: $TenantId"
Write-Host "  ApplicationId: $ApplicationId"
Write-Host "  ClientSecret: $ClientSecret"
Write-Host ""

# Install modules
Write-Host 'Installing required modules';
Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -AllowClobber -Scope CurrentUser -Force;
Install-Module  Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force;

# Login to PowerApps for the Admin commands
Write-Host 'Login to PowerApps for the Admin commands';
Add-PowerAppsAccount -Endpoint prod -TenantID $TenantId -ApplicationId $ApplicationId -ClientSecret $ClientSecret -Verbose;
 
# Login to PowerApps for the Xrm.Data commands
Write-Host "Login to PowerApps for the Xrm.Data commands";
$conn = Get-CrmConnection -ConnectionString $BuildToolsConnectionString;
if (!$conn)
{
    Write-Host "##vso[task.logissue type=error]Unable to get CRM Connection";
    Write-Warning "Unable to get CRM Connection";
    exit(1);
}

# Get the Orgid
$org = (Get-CrmRecords -conn $conn -EntityLogicalName organization).CrmRecords[0];
if (!$org)
{ 
    Write-Host "##vso[task.logissue type=error]Unable to get CRM Organization";
    Write-Warning "Unable to get CRM Organization";
    exit(1);
}
$orgid = $org.organizationid;

Write-Host "";
Write-Host "Connected to:" $conn.ConnectedOrgFriendlyName;
Write-Host "Environment ID:" $conn.EnvironmentId;
Write-Host "Org unique name:" $conn.ConnectedOrgUniqueName;
Write-Host "Org URL:" $conn.CrmConnectOrgUriActual.Host;
Write-Host "";

# Get connection references in the solution that are connected
Write-Host "Get Connected Connection References";
$connectionrefFetch = @"
<fetch>
  <entity name='connectionreference'>
    <attribute name='connectionreferencedisplayname' />
    <attribute name='connectionreferenceid' />
    <attribute name='connectionid' />
    <attribute name='connectorid' />
    <attribute name='owningbusinessunit' />
    <attribute name='owninguser' />
    <attribute name='ownerid' />
    <attribute name='createdonbehalfby' />
    <attribute name='createdby' />
    <attribute name='modifiedby' />
    <attribute name='modifiedonbehalfby' />
    <filter>
      <condition attribute='connectionid' operator='not-null' />
    </filter>
    <link-entity name='solutioncomponent' from='objectid' to='connectionreferenceid'>
      <link-entity name='solution' from='solutionid' to='solutionid'>
        <filter>
          <condition attribute='uniquename' operator='eq' value='Akoyanet' />
        </filter>
      </link-entity>
    </link-entity>
  </entity>
</fetch>
"@;
$connectionsrefs = (Get-CrmRecordsByFetch  -conn $conn -Fetch $connectionrefFetch -Verbose).CrmRecords;
 
# If there are no connection refeferences that are connected then exit
if ($connectionsrefs.Count -eq 0)
{
    Write-Host "##vso[task.logissue type=error]No Connection References that are connected in the solution '$BuildToolsSolutionName'";
    Write-Warning "No Connection References that are connected in the solution '$BuildToolsSolutionName'";
    exit(1);
}
 
$existingconnectionreferences = (ConvertTo-Json ($connectionsrefs | Select-Object -Property connectionreferencedisplayname, connectorid, connectionreferenceid, connectionid, owningbusinessunit, owninguser, ownerid, createdonbehalfby, createdby, modifiedby, modifiedonbehalfby));
#Write-Host "##vso[task.setvariable variable=CONNECTION_REFS]$existingconnectionreferences"
Write-Host "Connection References: $existingconnectionreferences";

# Get the first connection reference connector that is not null and load it to find who it was created by
#$connections = Get-AdminPowerAppConnection -EnvironmentName $conn.EnvironmentId -Filter $connectionsrefs[0].connectionid
$connections = Get-AdminPowerAppConnection -EnvironmentName $conn.EnvironmentId;
if (!$connections)
{
    Write-Host "##vso[task.logissue type=error]Unable to get Admin Power App connection";
    Write-Warning "Unable to get Admin Power App connection";
    exit(1);
}

Write-Host "";
Write-Host "All environment connections:";
$connectionsJson = (ConvertTo-Json ($connections | Select-Object -Property ConnectionName, ConnectionId, FullConnectorName, DisplayName, CreatedBy));
Write-Host $connectionsJson;

Write-Host "";
Write-Host "Getting CreatedBy User from first connection:";
$firstConnectionJson = (ConvertTo-Json ($connections[0] | Select-Object -Property ConnectionName, ConnectionId, FullConnectorName, DisplayName, CreatedBy));
Write-Host $firstConnectionJson;

$user = Get-CrmRecords -conn $conn -EntityLogicalName systemuser -FilterAttribute azureactivedirectoryobjectid -FilterOperator eq -FilterValue $connections[0].CreatedBy.id;
if (!$user)
{
    Write-Host "##vso[task.logissue type=error]Unable to get CreatedBy user. CreatedBy.id:" $connections[0].CreatedBy.id;
    Write-Warning "Unable to create to get CreatedBy user. CreatedBy.id:" $connections[0].CreatedBy.id;
    exit(1);
}
 
# Create a new Connection to impersonate the creator of the connection reference
$impersonatedconn = Get-CrmConnection -ConnectionString $BuildToolsConnectionString;
if (!$impersonatedconn)
{
    Write-Host "##vso[task.logissue type=error]Unable to create impersonated connection";
    Write-Warning "Unable to create impersonated connection";
    exit(1);
}
$impersonatedconn.OrganizationWebProxyClient.CallerId = $user.CrmRecords[0].systemuserid;

# Get the flows that are turned off
Write-Host "";
Write-Host "Get Flows that are turned off";
$fetchFlows = @"
<fetch>
    <entity name='workflow'>
    <attribute name='category' />
    <attribute name='name' />
    <attribute name='statecode' />
    <filter>
        <condition attribute='category' operator='eq' value='5' />
        <condition attribute='statecode' operator='eq' value='0' />
    </filter>
    <link-entity name='solutioncomponent' from='objectid' to='workflowid'>
        <link-entity name='solution' from='solutionid' to='solutionid'>
        <filter>
            <condition attribute='uniquename' operator='eq' value='$BuildToolsSolutionName' />
        </filter>
        </link-entity>
    </link-entity>
    </entity>
</fetch>
"@;
 
$flows = (Get-CrmRecordsByFetch -conn $conn -Fetch $fetchFlows -Verbose).CrmRecords;
if ($flows.Count -eq 0)
{
    Write-Host "##vso[task.logissue type=warning]No Flows that are turned off in '$BuildToolsSolutionName.'";
    Write-Host "No Flows that are turned off in '$BuildToolsSolutionName'";
    exit(0);
}

$flowJson = (ConvertTo-Json ($flows | Select-Object -Property workflowid, category, name, statecode));
Write-Host "";
Write-Host "Flows that are disabled: $flowJson";
Write-Host "";

try
{
    #statecode: 0 (Draft)
    #statecode: 1 (Activated)
    #
    #statuscode: 1 (Draft)
    #statuscode: 2 (Activated)

    Write-Host "GOapply Add All Applications to Review Group";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '2c618902-7de2-ec11-bb3d-0022480c598c' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Add Request to Review Group";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '5b3644c7-00f7-ec11-bb3d-000d3a5c0906' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Add to Review Group";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id 'ef6d98cd-c1c4-ec11-a7b6-000d3a3be81d' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Advance to Final Report Phase";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '01472d0b-f217-ed11-b83e-002248081438' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply AutoFill Next Phase";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id 'c566a946-22a6-ec11-983f-0022480b67c1' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Document Move Version 2.0";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '93e74ef0-3976-ec11-8d21-000d3a5cf59d' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Duplicate Application";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '4e3cd5d3-c58a-ec11-93b0-000d3a5b26c0' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Duplicate Review Group";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id 'e98d7973-d2e1-ec11-bb3d-0022480c5141' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Duplicate Review Group and Reviewers";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '4a5431f9-d5e1-ec11-bb3d-0022480c598c' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Invite Constituent";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '76ae7a00-4e75-ec11-8943-00224809eefd' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOapply Invite Contact";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '33ccdc60-dd77-ec11-8d21-000d3a5cfdba' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOfund Invitation Email";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id 'eba79434-563b-ed11-9db1-000d3a59640f' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

    Write-Host "GOfund Move Documents Flow";
    Set-CrmRecordState -conn $impersonatedconn -EntityLogicalName workflow -Id '21197efa-da39-ed11-9db1-000d3a5a6918' -StateCode Activated -StatusCode Activated -Verbose -Debug;
    Write-Host "Completed.";

}
catch
{
    Write-Warning "An error occored when activating flow"; #:$(($flow).name)"
    Write-Warning $_;
    exit(1);
}

Write-Host "";
Write-Host "DONE.";

exit;
