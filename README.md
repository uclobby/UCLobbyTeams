# UCLobbyTeams

Please use the PowerShell Gallery to install this module:
<br/>
<br/>PowerShell Gallery – UcLobbyTeams
<br/> https://www.powershellgallery.com/packages/UcLobbyTeams/
<br/>
<br/>Available cmdlets:
<br/>
<br/>Get-UcTeamsVersion
<br/>Get-UcM365TenantId 
<br/>Get-UcTeamsForest
<br/>Get-UcM365Domains
<br/>Test-UcTeamsOnlyModeReadiness
<br/>Get-UcTeamUsersEmail
<br/>Get-UcTeamsWithSingleOwner
<br/>Get-UcArch

<br/>More info:
<br/>https://uclobby.com/uclobby-teams-powershell-module/

<br/>
<br/>Change Log:
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
  <br/>New cmdlet that returns the Teams Forest, this is helpfull for Skype for Busines OnPrem to Teams migrations.</li>
</ul>
<br/>0.1.0 - 2022/03/25
<ul>
  <li>Initial Release uploaded to PowerShell Gallery</li>
</ul>
