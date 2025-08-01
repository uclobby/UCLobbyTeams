function Get-UcM365Domains {
    <#
        .SYNOPSIS
        Get Microsoft 365 Domains from a Tenant

        .DESCRIPTION
        This function returns a list of domains that are associated with a Microsoft 365 Tenant.

        .PARAMETER Domain
        Specifies a domain registered with Microsoft 365.

        .EXAMPLE
        PS> Get-UcM365Domains -Domain uclobby.com
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )

    $regex = "^(.*@)(.*[.].*)$"
    $outDomains = [System.Collections.ArrayList]::new()
    try {
        #2025-01-31: Only need to check this once per PowerShell session
        if (!($global:UCLobbyTeamsModuleCheck)) {
            Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
            $global:UCLobbyTeamsModuleCheck = $true
        }
        $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
    }
    catch [System.Net.WebException] {
        if ($PSItem.Exception.Message -eq "The remote server returned an error: (400) Bad Request.") {
            Write-Warning "The domain $Domain is not part of a Microsoft 365 Tenant."
        }
        else {
            Write-Warning $PSItem.Exception.Message
        }
    }
    catch {
        #2024-03-18: Support for GCC High tenants.
        try {
            $AllowedAudiences = Invoke-WebRequest -Uri ("https://login.microsoftonline.us/" + $Domain + "/metadata/json/1") -UseBasicParsing | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
        }
        catch {
            Write-Warning "Unknown error while checking domain: $Domain"
        }
    }
    try {
        foreach ($AllowedAudience in $AllowedAudiences) {
            $temp = [regex]::Match($AllowedAudience , $regex).captures.groups
            if ($temp.count -ge 2) {
                $tempObj = New-Object -TypeName PSObject -Property @{
                    Name = $temp[2].value
                }
                $outDomains.Add($tempObj) | Out-Null
            }
        }
    }
    catch {
        Write-Warning "Unknown error while checking domain: $Domain"
    }
    return $outDomains
}