
<#
.SYNOPSIS
Get Users Email Address that are in a Team

.DESCRIPTION
This function returns a list of users email address that are part of a Team.

.PARAMETER TeamName
Specifies Team Name

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
Function Get-UcTeamUsersEmail {
    [cmdletbinding(SupportsShouldProcess)]
    Param(
        [Parameter(Mandatory = $false)]
        [string]$TeamName,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Owner", "User", "Guest")] 
        [string]$Role
    )
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
            $Member = New-Object -TypeName PSObject -Property @{
                TeamGroupID     = $Team.GroupID
                TeamDisplayName = $Team.DisplayName
                TeamVisibility  = $Team.Visibility
                UPN             = $TeamMember.User
                Role            = $TeamMember.Role
                Email           = $Email
            }
            $Member.PSObject.TypeNames.Insert(0, 'TeamUsersEmail')
            $output.Add($Member) | Out-Null
        }
    }
    return $output
}