function Test-UcMgGraphConnection {
    param(
        [Parameter(mandatory = $false)]
        [string[]]$Scopes,
        [string[]]$AltScopes,
        [ValidateSet("AppOnly", "Delegated")]
        [string]$AuthType
    )
    <#
        .SYNOPSIS
        Test connection to Microsoft Graph

        .DESCRIPTION
        This function will validate if the current connection to Microsoft Graph has the required Scopes.
 
        Contributors: Daniel Jelinek, David Paulino

        Requirements:   Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

        .PARAMETER Scopes
        When present it will get detailed information from Teams Devices

        .EXAMPLE 
        PS> Connect-UcMgGraph "TeamworkDevice.Read.All","Directory.Read.All"

    #>
    #Checking if Microsoft.Graph is installed
    if (!(Get-Module Microsoft.Graph.Authentication -ListAvailable)) {
        Write-Warning ("Missing Microsoft.Graph.Authentication PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module Microsoft.Graph.Authentication") 
        return $false
    }
    $MgGraphContext = Get-MgContext

    if($AuthType -and $MgGraphContext.AuthType -ne $AuthType){
        Write-Warning ("Wrong Permission Type: " + $MgGraphContext.AuthType + ", this PowerShell cmdlet requires: $AuthType") 
        return $false
    }

    #Checking if we have the required scopes.
    $currentScopes = $MgGraphContext.Scopes
 
    $strScope = ""
    $strAltScope = ""
    $missingScopes = ""
    $missingAltScopes = ""
    $missingScope = $false
    $missingAltScope = $false

    foreach ($scope in $Scopes) {
        $strScope += "`"" + $scope + "`","
        if ($scope -notin $currentScopes) {
            $missingScope = $true
            $missingScopes += $scope + ","
        }
    }

    #20231013 - Added option to specify alternative scopes
    if ($missingScope -and $AltScopes) {
        foreach ($altScope in $AltScopes) {
            $strAltScope += "`"" + $altScope + "`","
            if ($altScope -notin $currentScopes) {
                $missingAltScope = $true
                $missingAltScopes += $altScope + ","
            }
        }
        #20231117 - Issue when AltScopes wasnt submitted
    }
    else {
        $missingAltScope = $true
    }

    if (!($currentScopes)) {
        Write-Warning  ("Not Connected to Microsoft Graph" + [Environment]::NewLine + "Please connect to Microsoft Graph before running this cmdlet." + [Environment]::NewLine + "Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " -Environment USGov")
        return $false
    }

    #If scopes are missing we need to connect using the required scopes
    if ($missingScope -and $missingAltScope) {
        if ($Scopes -and $AltScopes) {
            Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + " and missing alternative Scope(s): " + $missingAltScopes.Substring(0, $missingAltScopes.Length - 1) + `
                    [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                    [Environment]::NewLine + "Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-MgGraph -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) + `
                    [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-MgGraph -Environment USGov -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + " or Connect-MgGraph -Environment USGov -Scopes " + $strAltScope.Substring(0, $strAltScope.Length - 1) )
        }
        else {
            Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0, $missingScopes.Length - 1) + `
                    [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
                    [Environment]::NewLine + "Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0, $strScope.Length - 1) + `
                    [Environment]::NewLine + "US Gov (GCC-H) Tenant: Connect-MgGraph -Environment USGov -Scopes " + $strScope.Substring(0, $strScope.Length - 1))
        }
        return $false
    }
    else {
        return $true
    }
}