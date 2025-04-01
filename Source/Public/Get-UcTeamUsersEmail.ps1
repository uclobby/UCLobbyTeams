function Get-UcTeamUsersEmail {
    <#
        .SYNOPSIS
        Get Users Email Address that are in a Team

        .DESCRIPTION
        This function returns a list of users email address that are part of a Team.

        Requirements:   Microsoft Teams PowerShell Module (Install-Module MicrosoftTeams)

        .PARAMETER TeamName
        Specifies Team Name.

        .PARAMETER Role
        Specifies which roles to filter (Owner, User, Guest)

        .EXAMPLE
        PS> Get-UcTeamUsersEmail

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -TeamName "Marketing"

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -Role "Guest"

        .EXAMPLE
        PS> Get-UcTeamUsersEmail -TeamName "Marketing" -Role "Guest"
    #>
    [cmdletbinding(SupportsShouldProcess)]
    param(
        [string]$TeamName,
        [ValidateSet("Owner", "User", "Guest")] 
        [string]$Role
    )
    
    #region 2025-03-31: Check if connected with Teams PowerShell module. 
    if (!(Test-UcServiceConnection -Type TeamsPowerShell)) {
        return
    }
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    $output = [System.Collections.ArrayList]::new()
    if ($TeamName) {
        $Teams = Get-Team -DisplayName $TeamName
    }
    else {
        if ($ConfirmPreference) {
            $title = 'Confirm'
            $question = 'Are you sure that you want to list all Teams?'
            $choices = '&Yes', '&No'
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        }
        else {
            $decision = 0
        }
        if ($decision -eq 0) {
            $Teams = Get-Team
        }
        else {
            return
        }
    }
    foreach ($Team in $Teams) { 
        if ($Role) {
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID -Role $Role
        }
        else {
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID 
        }
        foreach ($TeamMember in $TeamMembers) {
            $Email = ( Get-csOnlineUser $TeamMember.User | Select-Object @{Name = 'PrimarySMTPAddress'; Expression = { $_.ProxyAddresses -cmatch '^SMTP:' -creplace 'SMTP:' } }).PrimarySMTPAddress
            $Member = [PSCustomObject][Ordered]@{
                TeamGroupID     = $Team.GroupID
                TeamDisplayName = $Team.DisplayName
                TeamVisibility  = $Team.Visibility
                UPN             = $TeamMember.User
                Role            = $TeamMember.Role
                Email           = $Email
            }
            $Member.PSObject.TypeNames.Insert(0, 'TeamUsersEmail')
            [void]$output.Add($Member) 
        }
    }
    return $output
}