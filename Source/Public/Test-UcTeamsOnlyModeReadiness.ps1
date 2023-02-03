<#
.SYNOPSIS
Check if a Tenant can be changed to TeamsOnly

.DESCRIPTION
This function will check if there is a lyncdiscover DNS Record that prevents a tenant to be switched to TeamsOnly.

.PARAMETER Domain
Specifies a domain registered with Microsoft 365

.EXAMPLE
PS> Test-UcTeamsOnlyModeReadiness

.EXAMPLE
PS> Test-UcTeamsOnlyModeReadiness -Domain uclobby.com
#>
Function Test-UcTeamsOnlyModeReadiness {
    Param(
        [Parameter(Mandatory = $false)]
        [string]$Domain
    )
    $outTeamsOnly = [System.Collections.ArrayList]::new()
    $connectedMSTeams = $false
    if ($Domain) {
        $365Domains = Get-UcM365Domains -Domain $Domain
    }
    else {
        try {
            $365Domains = Get-CsOnlineSipDomain
            $connectedMSTeams = $true
        }
        catch {
            Write-Error "Please Connect to before running this cmdlet with Connect-MicrosoftTeams"
        }
    }
    $DomainCount = ($365Domains.Count)
    $i = 1
    foreach ($365Domain in $365Domains) {
        $tmpDomain = $365Domain.Name
        Write-Progress -Activity "Teams Only Mode Readiness" -Status "Checking: $tmpDomain - $i of $DomainCount"  -PercentComplete (($i / $DomainCount) * 100)
        $i++
        $status = "Not Ready"
        $DNSRecord = ""
        $hasARecord = $false
        $enabledSIP = $false
        
        # 20220522 - Skipping if the domain is not SIP Enabled.
        if (!($connectedMSTeams)) {
            try {
                Invoke-WebRequest -Uri ("https://webdir.online.lync.com/AutoDiscover/AutoDiscoverservice.svc/root?originalDomain=" + $tmpDomain) | Out-Null
                $enabledSIP = $true
            }
            catch {
 
                if ($Error[0].Exception.Message -like "*404*") {
                    $enabledSIP = $false
                }
            }
        }
        else {
            if ($365Domain.Status -eq "Enabled") {
                $enabledSIP = $true
            }
        }
        if ($enabledSIP) {
            $DiscoverFQDN = "lyncdiscover." + $365Domain.Name
            $FederationFQDN = "_sipfederationtls._tcp." + $365Domain.Name
            $DNSResultA = (Resolve-DnsName $DiscoverFQDN -ErrorAction Ignore -Type A)
            $DNSResultSRV = (Resolve-DnsName $FederationFQDN -ErrorAction Ignore -Type SRV).NameTarget
            foreach ($tmpRecord in $DNSResultA) {
                if (($tmpRecord.NameHost -eq "webdir.online.lync.com") -and ($DNSResultSRV -eq "sipfed.online.lync.com")) {
                    break
                }
        
                if ($tmpRecord.Type -eq "A") {
                    $hasARecord = $true
                    $DNSRecord = $tmpRecord.IPAddress
                }
            }
            if (!($hasARecord)) {
                $DNSResultCNAME = (Resolve-DnsName $DiscoverFQDN -ErrorAction Ignore -Type CNAME)
                if ($DNSResultCNAME.count -eq 0) {
                    $status = "Ready"
                }
                if (($DNSResultCNAME.NameHost -eq "webdir.online.lync.com") -and ($DNSResultSRV -eq "sipfed.online.lync.com")) {
                    $status = "Ready"
                    $DNSRecord = $DNSResultCNAME.NameHost
                }
                else {
                    $DNSRecord = $DNSResultCNAME.NameHost
                }
                if ($DNSResultCNAME.Type -eq "SOA") {
                    $status = "Ready"
                }
            }
            $Validation = New-Object -TypeName PSObject -Property @{
                DiscoverRecord   = $DNSRecord
                FederationRecord = $DNSResultSRV
                Status           = $status
                Domain           = $365Domain.Name
            }

        }
        else {
            $Validation = New-Object -TypeName PSObject -Property @{
                DiscoverRecord   = ""
                FederationRecord = ""
                Status           = "Not SIP Enabled"
                Domain           = $365Domain.Name
            }
        }
        $Validation.PSObject.TypeNames.Insert(0, 'TeamsOnlyModeReadiness')
        $outTeamsOnly.Add($Validation) | Out-Null
    }
    return $outTeamsOnly
}
