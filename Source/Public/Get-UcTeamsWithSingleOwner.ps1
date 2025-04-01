function Get-UcTeamsWithSingleOwner {
    <#
        .SYNOPSIS
        Get Teams that have a single owner

        .DESCRIPTION
        This function returns a list of Teams that only have a single owner.

        Requirements:   Microsoft Teams PowerShell Module (Install-Module MicrosoftTeams)

        .EXAMPLE
        PS> Get-UcTeamsWithSingleOwner
    #>
    Get-UcTeamUsersEmail -Role Owner -Confirm:$false | Group-Object -Property TeamDisplayName | Where-Object { $_.Count -lt 2 } | Select-Object -ExpandProperty Group
}