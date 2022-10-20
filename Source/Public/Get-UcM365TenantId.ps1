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
    $regex = "^(.*@)(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})$"
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

    foreach ($AllowedAudience in $AllowedAudiences) {
        $temp = [regex]::Match($AllowedAudience , $regex).captures.groups
        if ($temp.count -ge 2) {
            return   New-Object -TypeName PSObject -Property @{ TenantID = $temp[2].value }
        }
    }
}
