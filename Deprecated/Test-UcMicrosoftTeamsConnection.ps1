function Test-UcMicrosoftTeamsConnection {
    <#
        .SYNOPSIS
        Test connection to Microsoft Teams PowerShell

        .DESCRIPTION
        This function will validate if the current session is connected to Microsoft Teams PowerShell.
    #>
    #Checking if MicrosoftTeams module is installed
    if (!(Get-Module MicrosoftTeams -ListAvailable)) {
        Write-Warning ("Missing MicrosoftTeams PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module MicrosoftTeams") 
        return $false
    }
    #We need to use a cmdlet to know if we are connected to MicrosoftTeams PowerShell
    try {
        Get-CsTenant -ErrorAction SilentlyContinue | Out-Null
        return $true
    }
    catch [System.UnauthorizedAccessException] {
        Write-Warning ("Please connect to Microsoft Teams PowerShell with Connect-MicrosoftTeams before running this cmdlet")
        return $false
    }
}