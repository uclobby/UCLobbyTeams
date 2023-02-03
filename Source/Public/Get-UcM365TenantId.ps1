<#
.SYNOPSIS
Get Microsoft 365 Tenant Id 

.DESCRIPTION
This function returns the Tenant ID associated with a domain that is part of a Microsoft 365 Tenant.

.PARAMETER Domain
Specifies a domain registered with Microsoft 365

.EXAMPLE
PS> Get-UcM365TenantId -Domain uclobby.com
#>
Function Get-UcM365TenantId {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )
    $regexTenantID = "^(.*@)(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})$"
    $regexOnMicrosoftDomain = "^(.*@)(?!.*mail)(.*.onmicrosoft.com)$"

    try {
        $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/" + $Domain + "/metadata/json/1") | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
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
        Write-Warning "Unknown error while checking domain: $Domain"
    }
    $output = [System.Collections.ArrayList]::new()
    $TenantID = ""
    foreach ($AllowedAudience in $AllowedAudiences) {
        $tempTID = [regex]::Match($AllowedAudience , $regexTenantID).captures.groups
        $tempID = [regex]::Match($AllowedAudience , $regexOnMicrosoftDomain).captures.groups
        if ($tempTID.count -ge 2) {
            $TenantID = $tempTID[2].value 
        }
        if ($tempID.count -ge 2) {
            $OnMicrosoftDomain = $tempID[2].value
        }
        if($TenantID -and $OnMicrosoftDomain){
            $M365TidPSObj = New-Object -TypeName PSObject -Property @{ TenantID = $TenantID
                OnMicrosoftDomain = $OnMicrosoftDomain}
            $M365TidPSObj.PSObject.TypeNames.Insert(0, 'M365TenantId')
            [void]$output.Add($M365TidPSObj)
        }
    }
    return $output
}
