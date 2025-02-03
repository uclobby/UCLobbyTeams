function Test-UcTeamsOnlyDNSRequirements {
    <#
        .SYNOPSIS
        Check if the DNS records are OK for a TeamsOnly Tenant.

        .DESCRIPTION
        This function will check if the DNS entries that were previously required.

        .PARAMETER Domain
        Specifies a domain registered with Microsoft 365

        .EXAMPLE
        PS> Test-UcTeamsOnlyDNSRequirements

        .EXAMPLE
        PS> Test-UcTeamsOnlyDNSRequirements -Domain uclobby.com
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$Domain,
        [switch]$All
    )

    $outDNSRecords = [System.Collections.ArrayList]::new()
    if ($Domain) {
        $365Domains = Get-UcM365Domains -Domain $Domain | Where-Object { $_.Name -notlike "*.onmicrosoft.com" }
    }
    else {
        try {
            #2025-01-31: Only need to check this once per PowerShell session
            if (!($global:UCLobbyTeamsModuleCheck)) {
                Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
                $global:UCLobbyTeamsModuleCheck = $true
            }
            #We only need to validate the Enable domains and exclude *.onmicrosoft.com
            $365Domains = Get-CsOnlineSipDomain | Where-Object { $_.Status -eq "Enabled" -and $_.Name -notlike "*.onmicrosoft.com" }
        }
        catch {
            Write-Host "Error: Please connect to Microsoft Teams PowerShell before running this cmdlet with Connect-MicrosoftTeams" -ForegroundColor Red
            return
        }
    }
    $DomainCount = ($365Domains.Count)
    $i = 1
    foreach ($365Domain in $365Domains) {
        $tmpDomain = $365Domain.Name
        Write-Progress -Activity "Teams Only DNS Requirements" -Status "Checking: $tmpDomain - $i of $DomainCount"
        $i++
        $DiscoverFQDN = "lyncdiscover." + $365Domain.Name
        $SIPFQDN = "sip." + $365Domain.Name
        $SIPTLSFQDN = "_sip._tls." + $365Domain.Name
        $FederationFQDN = "_sipfederationtls._tcp." + $365Domain.Name
        $DNSResultDiscover = (Resolve-DnsName $DiscoverFQDN -Type CNAME -ErrorAction Ignore).NameHost
        $DNSResultSIP = (Resolve-DnsName $SIPFQDN -Type CNAME -ErrorAction Ignore).NameHost
        $DNSResultSIPTLS = (Resolve-DnsName $SIPTLSFQDN -Type SRV -ErrorAction Ignore).NameTarget
        $DNSResultFederation = (Resolve-DnsName $FederationFQDN -Type SRV -ErrorAction Ignore)

        $DNSDiscover = ""
        if ($DNSResultDiscover -and ($All -or $DNSResultDiscover.contains("online.lync.com") -or $DNSResultDiscover.contains("online.gov.skypeforbusiness.us"))) {
            $DNSDiscover = $DNSResultDiscover
        }

        $DNSSIP = ""
        if ($DNSResultSIP -and ($All -or $DNSResultSIP.contains("online.lync.com") -or $DNSResultSIP.contains("online.gov.skypeforbusiness.us"))) {
            $DNSSIP = $DNSResultSIP
        }
        
        $DNSSIPTLS = ""
        if ($DNSResultSIPTLS -and ($All -or $DNSResultSIPTLS.contains("online.lync.com") -or $DNSResultSIPTLS.contains("online.gov.skypeforbusiness.us"))) {
            $DNSSIPTLS = $DNSResultSIPTLS
        }

        $DNSFederation = ""
        if ([string]::IsNullOrEmpty($DNSResultFederation.NameTarget)) {
            $DNSFederation = "Not configured"
        }
        elseif (($DNSResultFederation.NameTarget.equals("sipfed.online.lync.com") -or $DNSResultFederation.NameTarget.equals("sipfed.online.gov.skypeforbusiness.us")) -and $DNSResultFederation.Port -eq 5061) {
            $DNSFederation = "OK"
        }
        else {
            $DNSFederation = "NOK - " + $DNSResultFederation.NameTarget + ":" + $DNSResultFederation.Port 
        }

        if ($DNSDiscover -or $DNSSIP -or $DNSSIPTLS -or $DNSFederation) {
            $tmpDNSRecord = New-Object -TypeName PSObject -Property @{
                Domain           = $365Domain.Name
                DiscoverRecord   = $DNSDiscover
                SIPRecord        = $DNSSIP
                SIPTLSRecord     = $DNSSIPTLS
                FederationRecord = $DNSFederation
            }
            $tmpDNSRecord.PSObject.TypeNames.Insert(0, 'TeamsOnlyDNSRequirements')
            [void]$outDNSRecords.Add($tmpDNSRecord)
        }
    }
    return $outDNSRecords | Sort-Object Domain
}