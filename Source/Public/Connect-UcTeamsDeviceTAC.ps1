function Connect-UcTeamsDeviceTAC {
    param(
        [string]$TenantID
    )
    #Create TeamsDeviceTAC Entra Service if doens't exist
    if (!(Get-EntraService -Name TeamsDeviceTAC)) {
        $teamsDeviceTacCfg = @{
            Name          = 'TeamsDeviceTAC'
            ServiceUrl    = 'https://admin.devicemgmt.teams.microsoft.com/'
            Resource      = 'https://admin.devicemgmt.teams.microsoft.com'
            DefaultScopes = @("https://devicemgmt.teams.microsoft.com/.default")
            HelpUrl       = ''
            Header        = @{ }
            NoRefresh     = $false
        }
        Register-EntraService @teamsDeviceTacCfg
    }
    if ($TenantID) {
        Connect-EntraService  -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Service TeamsDeviceTAC -TenantID $TenantID
    }
    else {
        Connect-EntraService  -ClientId '12128f48-ec9e-42f0-b203-ea49fb6af367' -Service TeamsDeviceTAC
    }
}