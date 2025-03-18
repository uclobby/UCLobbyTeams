function Test-UcAPIConnection {
    param(
        [Parameter(mandatory=$true)]
        [ValidateSet("TeamsModule","TeamsDeviceTAC")]
        [string]$Type
    )
    
    switch ($Type) {
        "TeamsModule" { 
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
        "TeamsDeviceTAC" {
            #Checking if EntraAuth module is installed
            if (!(Get-Module EntraAuth -ListAvailable)) {
                Write-Warning ("Missing EntraAuth PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module EntraAuth") 
                return $false
            }

            if(Get-EntraToken TeamsDeviceTAC){
                return $true
            } else {
                Write-Warning "Please connect to Teams TAC API with Connect-UcTeamsDeviceTAC before running this cmdlet"
            }
        }
        Default {
            return $false
        }
    }
}