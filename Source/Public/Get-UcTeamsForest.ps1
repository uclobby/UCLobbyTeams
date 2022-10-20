<#
.SYNOPSIS
Get Teams Forest 

.DESCRIPTION
This function returns the forest for a SIP enabled domain.

.PARAMETER Domain
Specifies a domain registered with Microsoft 365

.EXAMPLE
PS> Get-UcTeamsForest -Domain uclobby.com
#>
Function Get-UcTeamsForest {
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Domain
    )
    $regex = "^.*[redirect].*(webdir)(\w*)(.online.lync.com).*$"
    try {
        $WebRequest = Invoke-WebRequest -Uri ("https://webdir.online.lync.com/AutoDiscover/AutoDiscoverservice.svc/root?originalDomain=" + $Domain)
        $temp = [regex]::Match($WebRequest, $regex).captures.groups
        $result = New-Object -TypeName PSObject -Property @{
                
            Domain       = $Domain
            Forest       = $temp[2].Value
            MigrationURL = "https://admin" + $temp[2].Value + ".online.lync.com/HostedMigration/hostedmigrationService.svc"
        }
        return $result
    }
    catch {
 
        if ($Error[0].Exception.Message -like "*404*") {
            Write-Warning ($Domain + " is not enabled for SIP." )
        }
    }
}