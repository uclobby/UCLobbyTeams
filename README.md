# UCLobbyTeams

Please use the PowerShell Gallery to install this module:
<br/>
<br/>PowerShell Gallery – UcLobbyTeams
<br/> https://www.powershellgallery.com/packages/UcLobbyTeams/
<br/>
<br/>Available cmdlets:

<br/>Get-UcM365Domains
<br/>Get-UcM365TenantId 

<br/>Get-UcTeamsDevice
<br/>Get-UcTeamsForest
<br/>Get-UcTeamsVersion
<br/>Get-UcTeamsVersionBatch

<br>Update-UcTeamsDevice

<br/>Get-UcTeamsWithSingleOwner
<br/>Get-UcTeamUsersEmail

<br/>Get-UcConditionalAccessPolicy

<br/>Test-UcTeamsDevicesConditionalAccessPolicy
<br/>Test-UcTeamsDevicesCompliancePolicy
<br/>Test-UcTeamsOnlyModeReadiness

<br/>Get-UcArch
<br/>Test-UcModuleUpdateAvailable

<br/>More info:
<br/>https://uclobby.com/uclobby-teams-powershell-module/

<br/>Change Log:

<br/>0.3.3 - 2023
<ul>
<li>Test-UcTeamsDevicesConditionalAccessPolicy
<br/>Fix: Detailed output was showing deviceFilter for GrantControl settings.
</li>
</ul>
<br/>0.3.2 - 2023/05/01
<ul>
<li>Get-UcTeamsDevice
<br/>Fix: Issue where only the first MS graph page from a request was returned.
</li>
<li>Update-UcTeamsDevice
<br/>Fix: Issue where only the first MS graph page from a request was returned.
</li>
</ul>
<br/>0.3.1 - 2023/04/25
<ul>
<li>General
  <br/>Change: Added a check, in all cmdlets, if there is a newer version of module available (Test-UcModuleUpdateAvailable).
</li>
<li>Update-UcTeamsDevice
  <br/>New cmdlet: Allow to send update commands to Teams Android Devices using MS Graph.
</li>
<li>Test-UcModuleUpdateAvailable
  <br/>New cmdlet: Check if a PowerShell module has a new update in PowerShell Gallery and if it's installed.
</li>
<li>Test-UcTeamsDevicesConditionalAccessPolicy
  <br/>Fix: In some cases the number of groups was too large, so now we only get the display name for the groups included/excluded.
  <br/>Feature: Ability to export to CSV with a new parameter -ExportCSV
</li>
<li>Test-UcTeamsDevicesCompliancePolicy
  <br/>Fix: In some cases the number of groups was too large, so now we only get the display name for the groups included/excluded.
  <br/>Feature: Ability to export to CSV with a new parameter -ExportCSV
</li>
</ul>
<br/>0.2.7 - 2023/04/13
<ul>
  <li>General
  <br/>Change: We need to connect to Microsoft Graph before using any cmdlet in the module. In the previous versions a connection was attempted, we remove this since in some scenarios we might want to use different authentication methods or environment.
  </li>
  <li>Get-UcTeamsDevice
  <br/>Fix: An exception could be raised if the User was null.
  <br/>Fix: Cycle all pages in MS Graph Response 
  <br/>Feature: Ability to export to CSV with a new parameter -ExportCSV
  </li>
  <li>Get-UcTeamsVersionBatch
  <br />New cmdlet: This allows to get the teams version from a list of computers from a CSV file. 
  </li>
</ul>
<br/>0.2.6 - 2023/02/10
<ul>
  <li>Get-UcM365TenantId
  <br/>Added support for Multi Geo Tenants. 
  </li>
  <li>Test-UcTeamsDevicesConditionalAccessPolicy
  <br/>Fixed missing CloudApps setting value.
  </li>
</ul>
<br/>0.2.5 - 2023/02/03
<ul>
  <li>Get-UcTeamsVersion
  <br/>Added Credential parameter that will be used to connect to the remote computer.
  </li>
  <li>Get-UcTeamsDevice
  <br/>Remove Device ID from parameters and output since that ID is the object in MS Graph and not the Device ID in Azure AD.
  </li>
  <li>Test-UcTeamsDevicesCompliancePolicy
  <br/>Change in the default behavior, now without any switch only policies that are assigned to a group will be checked.
  <br/>Output now includes which groups are included/excluded.
  <br/>Display warning with the number of compliance policy skipped (Not associated with a group).
  <br/>Added All switch to allow check policies even the policies without group assignment.
  <br/>Added User UPN parameter to check Compliance policies applied to the specified user.
  <br/>Added Device ID parameter to check Compliance policies applied to specific device.
  <br/>Detailed switch will only output the unsupported settings.
  <br/>Added IncludedSupported switch to show all checked policy settings for each policy.
  <br/>Added Setting Description in the Detailed output to make it easier identify it in Microsoft Endpoint Manager admin center.
  <br/>Added check for unsupported settings for MTR Windows (Windows Compliance Policy). 
  </li>
  <li>Test-UcTeamsDevicesConditionalAccessPolicy
  <br/>Change in the default behavior, now without any switch only policies that are assigned to a group will be checked.
  <br/>Output now includes which groups are included/excluded.
  <br/>Display warning with the number of compliance policy skipped (Not associated with a group or Teams Application).
  <br/>Added All switch to allow check policies even the policies without group assignment.
  <br/>Added User UPN parameter to only check Conditional Access policies applied to the specified user.
  <br/>Detailed switch will only output the unsupported settings.
  <br/>Added IncludedSupported switch to show all checked policy settings for each policy.
  <br/>Added Setting Description in the Detailed output to make it easier identify it in Microsoft Endpoint Manager admin center.
  </li>
  <li>Get-UcM365TenantId
  <br/>The output will also include the OnMicrosoft.com Domain for that tenant.
  </li>
</ul>
<br/>0.2.0 - 2022/10/20
<ul>
  <li>Get-UcTeamsVersion
  <br/>Added Computer parameter to get Teams version on a remote machine.
  <br/>Added Path parameter to specify a path that contains Teams log files.</li>
  <li>Get-UcTeamsDevice
  <br/>New cmdlet that gets Microsoft Teams Devices information using MS Graph API.
  </li>
  <li>Test-UcTeamsDevicesConditionalAccessPolicy
  <br/>New cmdlet that that validates which Conditional Access policies are supported by Microsoft Teams Android Devices.</li>
  <li>Test-UcTeamsDevicesCompliancePolicy
  <br/>New cmdlet that validates which Intune Compliance policies are supported by Microsoft Teams Android Devices.</li>
</ul>
<br/>0.1.3 - 2022/06/10
<ul>
  <li>Get-UcTeamVersion
  <br/>Fixed the issue where the version was limited to 4 digits.
  <br/>Added information for Ring, Environment, Region.</li>
  <li>Get-UcTeamUsersEmail
  <br/>This function returns a list of users email address that are part of a Team.</li>
  </li>
  <li>Get-UcTeamsWithSingleOwner
  <br/>This function returns a list of Teams that only have a single owner.</li>
  </li>
</ul>
<br/>0.1.2 - 2022/05/23
<ul>
  <li>Test-UcTeamsOnlyModeReadiness
  <br/>Add an additional check to skip non SIP enabled domains;
  <br/>Add progress status.</li>
  </li>
  <li>Get-UcTeamsForest
  <br/>New cmdlet that returns the Teams Forest, this is helpful for Skype for Business OnPrem to Teams migrations.</li>
</ul>
<br/>0.1.0 - 2022/03/25
<ul>
  <li>Initial Release uploaded to PowerShell Gallery</li>
</ul>