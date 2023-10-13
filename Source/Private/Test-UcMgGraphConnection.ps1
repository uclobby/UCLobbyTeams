<#
.SYNOPSIS
Test connection to Microsoft Graph

.DESCRIPTION
This function will validate if the current connection to Microsoft Graph has the required Scopes and will attempt to connect if we are missing one.
 
Contributors: Daniel Jelinek, David Paulino

Requirements:   Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

.PARAMETER Scopes
When present it will get detailed information from Teams Devices

.EXAMPLE 
PS> Connect-UcMgGraph "TeamworkDevice.Read.All","Directory.Read.All"

#>

function Test-UcMgGraphConnection {
    Param(
    [string[]]$Scopes,
    [string[]]$AltScopes
    )
    #Checking if Microsoft.Graph is installed
    if(!(Get-Module Microsoft.Graph.Authentication -ListAvailable)){
        Write-Warning ("Missing Microsoft.Graph PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module Microsoft.Graph") 
        return $false
    }
    #Checking if we have the required scopes.
    $currentScopes = (Get-MgContext).Scopes
 
    $strScope =""
    $missingScopes = ""
    $missingScope = $false    
    foreach($scope in $Scopes){
        $strScope += "`"" + $scope + "`","
        if(($scope -notin $currentScopes) -and $currentScopes){
            $missingScope = $true
            $missingScopes += $scope + ","
        }
    }

    #20231013 - Added option to specify alternative scopes
    if(!($missingScope)){
        $strAltScope =""
        $missingAltScopes = ""
        $missingAltScope = $false
        foreach($altScope in $AltScopes){
            $strAltScope += "`"" + $altScope + "`","
            if(($altScope -notin $currentScopes) -and $currentScopes){
                $missingAltScope = $true
                $missingAltScopes += $altScope + ","
            }
        }
    }

    if(!($currentScopes)){
        Write-Warning  ("Not Connected to Microsoft Graph" + [Environment]::NewLine  + "Please connect to Microsoft Graph before running this cmdlet." + [Environment]::NewLine +"Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1) + [Environment]::NewLine  + "US Gov (GCC-H) Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1)  + " -Environment USGov")
        return $false
    }

    #If scopes are missing we need to connect using the required scopes
    if($missingScope -and $missingAltScope){
        if($Scopes -and $AltScopes){
            Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0,$missingScopes.Length -1) + " and missing alternative Scope(s): " + $missingAltScopes.Substring(0,$missingAltScopes.Length -1) + `
            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
            [Environment]::NewLine + "Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1) + " or Connect-MgGraph -Scopes " + $strAltScope.Substring(0,$strAltScope.Length -1)  + `
            [Environment]::NewLine  + "US Gov (GCC-H) Tenant: Connect-MgGraph -Environment USGov -Scopes " + $strScope.Substring(0,$strScope.Length -1)  + " or  Connect-MgGraph -Environment USGov -Scopes " + $strAltScope.Substring(0,$strAltScope.Length -1) )
        } else {
            Write-Warning  ("Missing scope(s): " + $missingScopes.Substring(0,$missingScopes.Length -1) + `
            [Environment]::NewLine + "Please reconnect to Microsoft Graph before running this cmdlet." + `
            [Environment]::NewLine +"Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1) + `
            [Environment]::NewLine  + "US Gov (GCC-H) Tenant: Connect-MgGraph -Environment USGov -Scopes " + $strScope.Substring(0,$strScope.Length -1))
        }
        return $false
    } else {
        return $true
    }
}