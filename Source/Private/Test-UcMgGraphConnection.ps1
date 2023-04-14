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
    [string[]]$Scopes
    )
    #Checking if Microsoft.Graph is installed
    if(!(Get-Module Microsoft.Graph.Authentication -ListAvailable)){
        Write-Warning ("Missing Microsoft.Graph PowerShell module. Please install it with:" + [Environment]::NewLine + "Install-Module Microsoft.Graph") 
        return $false
    }
    #Checking if we have the required scopes.
    $currentScopes = (Get-MgContext).Scopes
    $strScope =""
    $missingScope = $false
    foreach($scope in $Scopes){
        $strScope += "`"" + $scope + "`","
        if(($scope -notin $currentScopes) -and $currentScopes){
            $missingScope = $true
            Write-Warning -Message ("Missing Scope in Microsoft Graph: " + $scope) 
        }
    }

    if(!($currentScopes)){
        Write-Warning  ("Not Connected to Microsoft Graph" + [Environment]::NewLine  + "Please connect to Microsoft Graph before running this cmdlet." + [Environment]::NewLine +"Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1) + [Environment]::NewLine  + "US Gov (GCC-H) Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1)  + " -Environment USGov")
        return $false
    }

    #If scopes are missing we need to connect using the required scopes
    if($missingScope){
        Write-Warning  ("Some scopes are missing please reconnect to Microsoft Graph before running this cmdlet." + [Environment]::NewLine +"Commercial Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1) + [Environment]::NewLine  + "US Gov (GCC-H) Tenant: Connect-MgGraph -Scopes " + $strScope.Substring(0,$strScope.Length -1)  + " -Environment USGov")
        return $false
    } else {
        return $true
    }
}