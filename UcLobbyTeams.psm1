<#
	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.
#>


<#
.SYNOPSIS
Funcion to get the Architecture from .exe file

.DESCRIPTION
Based on PowerShell script Get-ExecutableType.ps1 by David Wyatt, please check the complete script in:

Identify 16-bit, 32-bit and 64-bit executables with PowerShell
https://gallery.technet.microsoft.com/scriptcenter/Identify-16-bit-32-bit-and-522eae75
#>
Function Get-UcArch([string]$FilePath) {
try
    {
        $stream = New-Object System.IO.FileStream(
        $FilePath,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read )
        $exeType = 'Unknown'
        $bytes = New-Object byte[](4)
        if ($stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) -eq 0x3C -and $stream.Read($bytes, 0, 4) -eq 4)
        {
            if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 4) }
            $peHeaderOffset = [System.BitConverter]::ToUInt32($bytes, 0)

            if ($stream.Length -ge $peHeaderOffset + 6 -and
                $stream.Seek($peHeaderOffset, [System.IO.SeekOrigin]::Begin) -eq $peHeaderOffset -and
                $stream.Read($bytes, 0, 4) -eq 4 -and
                $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0 -and $bytes[3] -eq 0)
            {
                $exeType = 'Unknown'
                if ($stream.Read($bytes, 0, 2) -eq 2)
                {
                    if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 2) }
                    $machineType = [System.BitConverter]::ToUInt16($bytes, 0)
                    switch ($machineType)
                    {
                        0x014C { $exeType = 'x86' }
                        0x8664 { $exeType = 'x64' }
                    }
                }
            }
        }
        return $exeType
    }
    catch
    {
        return "Unknown"
    }
    finally
    {
        if ($null -ne $stream) { $stream.Dispose() }
    }
}

<#
.SYNOPSIS
Get Microsoft Teams Desktop Version
.DESCRIPTION
This function returns the installed Microsoft Teams desktop version for each user profile.
#>
Function Get-UcTeamsVersion {
    $regexVersion = '("version":")([0-9.]*)'
    $regexRing = '("ring":")(\w*)'
    $regexEnv = '("environment":")(\w*)'
    $regexCloudEnv = '("cloudEnvironment":")(\w*)'
    $regexRegion = '("region":")([a-zA-Z0-9._-]*)'
    
    $outTeamsVersion = [System.Collections.ArrayList]::new()
    
    $currentDateFormat = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
    $Profiles = Get-childItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object {Get-ItemProperty $_.pspath } | Where-Object {$_.fullprofile -eq 1}
    foreach($Profile in $Profiles){
        $TeamsSettingPath = $Profile.ProfileImagePath + "\AppData\Roaming\Microsoft\Teams\settings.json"
        if(Test-Path $TeamsSettingPath -ErrorAction SilentlyContinue) {
            $TeamsSettings = Get-Content -Path $TeamsSettingPath
            $Version = ""
            $Ring = ""
            $Env = ""
            $CloudEnv = ""
            $Region = ""
            try {
                $VersionTemp = [regex]::Match($TeamsSettings,$regexVersion).captures.groups
                if($VersionTemp.Count -ge 2){
                    $Version = $VersionTemp[2].value
                }
                $RingTemp = [regex]::Match($TeamsSettings,$regexRing).captures.groups
                if($RingTemp.Count -ge 2){
                    $Ring = $RingTemp[2].value
                }
                $EnvTemp = [regex]::Match($TeamsSettings,$regexEnv).captures.groups
                if($EnvTemp.Count -ge 2){
                    $Env = $EnvTemp[2].value
                }
                $CloudEnvTemp = [regex]::Match($TeamsSettings,$regexCloudEnv).captures.groups
                if($CloudEnvTemp.Count -ge 2){
                    $CloudEnv = $CloudEnvTemp[2].value
                }
                $RegionTemp = [regex]::Match($TeamsSettings,$regexRegion).captures.groups
                if($RegionTemp.Count -ge 2){
                    $Region = $RegionTemp[2].value
                }
            } catch { }
            $TeamsApp = $Profile.ProfileImagePath + "\AppData\Local\Microsoft\Teams\current\Teams.exe"
            $InstallDateStr = Get-Content ($Profile.ProfileImagePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt")
            $TeamsVersion = New-Object –TypeName PSObject -Property @{
                Profile = (New-Object System.Security.Principal.SecurityIdentifier($Profile.PSChildName)).Translate( [System.Security.Principal.NTAccount]).Value
                ProfilePath = $Profile.ProfileImagePath
                Version = $Version
                Ring = $Ring
                Environment = $Env
                CloudEnvironment = $CloudEnv
                Region = $Region
                Arch = Get-UcArch $TeamsApp
                InstallDate = [Datetime]::ParseExact($InstallDateStr, 'M/d/yyyy', $null) | Get-Date -Format $currentDateFormat
            }
            
            $TeamsVersion.PSObject.TypeNAmes.Insert(0,'TeamsVersion')
            $outTeamsVersion.Add($TeamsVersion) | Out-Null
        }
    }
    return $outTeamsVersion
}

<#
.SYNOPSIS
Get Microsoft 365 Tenant Id 
.DESCRIPTION
This function returns the Tenant ID associated with a domain that is part of a Microsoft 365 Tenant.
#>
Function Get-UcM365TenantId{
Param(
 [Parameter(Mandatory=$true)]
 [string]$Domain
)
    $regex = "^(.*@)(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})$"
    try{
    $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/"+$Domain+"/metadata/json/1") | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
    } catch [System.Net.WebException]{
        if($PSItem.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
            Write-Warning "The domain $Domain is not part of a Microsoft 365 Tenant."
        } else {
            Write-Warning $PSItem.Exception.Message
        }
    } catch {
        Write-Warning "Unknown error while checking domain: $Domain"
    }

    foreach($AllowedAudience in $AllowedAudiences){
        $temp = [regex]::Match($AllowedAudience ,$regex).captures.groups
        if($temp.count -ge 2) {
            return  $TeamsVersion = New-Object –TypeName PSObject -Property @{ TenantID = $temp[2].value}
        }
    }
}

<#
.SYNOPSIS
Get Teams Forest 
.DESCRIPTION
This function returns the forest for a SIP enabled domain.
#>
Function Get-UcTeamsForest{
Param(
 [Parameter(Mandatory=$true)]
 [string]$Domain
)
    $regex = "^.*[redirect].*(webdir)(\w*)(.online.lync.com).*$"
    try{
        $WebRequest = Invoke-WebRequest -Uri ("https://webdir.online.lync.com/AutoDiscover/AutoDiscoverservice.svc/root?originalDomain="+$Domain)
        $temp = [regex]::Match($WebRequest,$regex).captures.groups
        $forest = 

        $result =  New-Object –TypeName PSObject -Property @{
                
                Domain = $Domain
                Forest = $temp[2].Value
                MigrationURL = "https://admin" + $temp[2].Value + ".online.lync.com/HostedMigration/hostedmigrationService.svc"
            }
        return $result
    } catch {
 
        if ($Error[0].Exception.Message -like "*404*"){
             Write-Warning ($Domain + " is not enabled for SIP." )
        }
    }
}

<#
.SYNOPSIS
Get Microsoft 365 Domains from a Tenant
.DESCRIPTION
This function returns a list of domains that are associated with a Microsoft 365 Tenant.
#>
Function Get-UcM365Domains{
Param(
 [Parameter(Mandatory=$true)]
 [string]$Domain
)
    $regex = "^(.*@)(.*[.].*)$"
    $outDomains = [System.Collections.ArrayList]::new()

    try{
        $AllowedAudiences = Invoke-WebRequest -Uri ("https://accounts.accesscontrol.windows.net/"+$Domain+"/metadata/json/1") | ConvertFrom-Json | Select-Object -ExpandProperty allowedAudiences
        foreach($AllowedAudience in $AllowedAudiences){
            $temp = [regex]::Match($AllowedAudience ,$regex).captures.groups
            if($temp.count -ge 2) {
                $tempObj =  New-Object –TypeName PSObject -Property @{
                    Name = $temp[2].value
                }
                $outDomains.Add($tempObj) | Out-Null
            }
        }
    } catch [System.Net.WebException]{
        if($PSItem.Exception.Message -eq "The remote server returned an error: (400) Bad Request."){
            Write-Warning "The domain $Domain is not part of a Microsoft 365 Tenant."
        } else {
            Write-Warning $PSItem.Exception.Message
        }
    } catch {
        Write-Warning "Unknown error while checking domain: $Domain"
    }
    return $outDomains
}

<#
.SYNOPSIS
Check if a Tenant can be changed to TeamsOnly
.DESCRIPTION
This function will check if there is a lyncdiscover DNS Record that prevents a tenant to be switched to TeamsOnly.
#>
Function Test-UcTeamsOnlyModeReadiness {
Param(
 [Parameter(Mandatory=$false)]
 [string]$Domain
)
    $outTeamsOnly = [System.Collections.ArrayList]::new()
    $connectedMSTeams = $false
    if($Domain){
        $365Domains = Get-UcM365Domains -Domain $Domain
    } else {
        try{
            $365Domains =  Get-CsOnlineSipDomain
            $connectedMSTeams = $true
        } catch {
            Write-Error "Please Connect to before running this cmdlet with Connect-MicrosoftTeams"
        }
    }
    $DomainCount = ($365Domains.Count)
    $i= 1
    foreach($365Domain in $365Domains) {
        $tmpDomain = $365Domain.Name
        Write-Progress -Activity "Teams Only Mode Readiness" -Status "Checking: $tmpDomain - $i of $DomainCount"  -PercentComplete (($i / $DomainCount) * 100)
        $i++
        $status = "Not Ready"
        $DNSRecord = ""
        $hasARecord = $false
        $enabledSIP = $false
        
        # 20220522 - Skipping if the domain is not SIP Enabled.
        if(!($connectedMSTeams)) {
            try{
                Invoke-WebRequest -Uri ("https://webdir.online.lync.com/AutoDiscover/AutoDiscoverservice.svc/root?originalDomain="+$tmpDomain) | Out-Null
                $enabledSIP = $true
            } catch {
 
                if ($Error[0].Exception.Message -like "*404*"){
                    $enabledSIP = $false
                }
            }
        } else {
            if($365Domain.Status -eq "Enabled"){
                $enabledSIP = $true
            }
        }
        if($enabledSIP){
            $DiscoverFQDN = "lyncdiscover." + $365Domain.Name
            $FederationFQDN = "_sipfederationtls._tcp." + $365Domain.Name
            $DNSResultA = (Resolve-DnsName $DiscoverFQDN -ErrorAction Ignore -Type A)
            $DNSResultSRV = (Resolve-DnsName $FederationFQDN -ErrorAction Ignore -Type SRV).NameTarget
            foreach ($tmpRecord in $DNSResultA){
                if(($tmpRecord.NameHost -eq "webdir.online.lync.com")-and ($DNSResultSRV -eq "sipfed.online.lync.com")) {
                    break
                }
        
                if($tmpRecord.Type -eq "A"){
                    $hasARecord = $true
                    $DNSRecord = $tmpRecord.IPAddress
                }
            }
            if(!($hasARecord)){
                $DNSResultCNAME = (Resolve-DnsName $DiscoverFQDN -ErrorAction Ignore -Type CNAME)
                if($DNSResultCNAME.count -eq 0){
                    $status = "Ready"
                }
                if(($DNSResultCNAME.NameHost -eq "webdir.online.lync.com") -and ($DNSResultSRV -eq "sipfed.online.lync.com")) {
                    $status = "Ready"
                    $DNSRecord = $DNSResultCNAME.NameHost
                } else {
                    $DNSRecord = $DNSResultCNAME.NameHost
                }
                if($DNSResultCNAME.Type -eq "SOA"){
                    $status = "Ready"
                }
            }
            $Validation =  New-Object –TypeName PSObject -Property @{
                DiscoverRecord = $DNSRecord
                FederationRecord = $DNSResultSRV
                Status = $status
                Domain = $365Domain.Name
            }

        } else {
            $Validation =  New-Object –TypeName PSObject -Property @{
                DiscoverRecord = ""
                FederationRecord = ""
                Status = "Not SIP Enabled"
                Domain = $365Domain.Name
            }
        }
        $Validation.PSObject.TypeNAmes.Insert(0,'TeamsOnlyModeReadiness')
        $outTeamsOnly.Add($Validation) | Out-Null
    }
    return $outTeamsOnly
}

<#
.SYNOPSIS
Get Users Email Address that are in a Team
.DESCRIPTION
This function returns a list of users email address that are part of a Team.
#>
Function Get-UcTeamUsersEmail{
    [cmdletbinding(SupportsShouldProcess)]
Param(
 [Parameter(Mandatory=$false)]
 [string]$TeamName,
 [Parameter(Mandatory=$false)]
 [ValidateSet("Owner", "User", "Guest")] 
 [string]$Role
)
    $output = [System.Collections.ArrayList]::new()
    if($TeamName){
        $Teams = Get-Team -DisplayName $TeamName
    } else {
        if($ConfirmPreference){
            $title    = 'Confirm'
            $question = 'Are you sure that you want to list all Teams?'
            $choices  = '&Yes', '&No'
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
        } else {
            $decision = 0
        }
        if ($decision -eq 0) {
            $Teams = Get-Team
        } else {
            return
        }
    }
    foreach($Team in $Teams) { 
        if($Role){
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID -Role $Role
        } else {
            $TeamMembers = Get-TeamUser -GroupId $Team.GroupID 
        }
        foreach ($TeamMember in $TeamMembers){
            $Email =( Get-csOnlineUser $TeamMember.User |Select-Object @{Name='PrimarySMTPAddress';Expression={$_.ProxyAddresses -cmatch '^SMTP:' -creplace 'SMTP:'}}).PrimarySMTPAddress
            $Member =  New-Object –TypeName PSObject -Property @{
                    TeamGroupID = $Team.GroupID
                    TeamDisplayName = $Team.DisplayName
                    TeamVisibility = $Team.Visibility
                    UPN = $TeamMember.User
                    Role = $TeamMember.Role
                    Email = $Email
            }
            $Member.PSObject.TypeNAmes.Insert(0,'TeamUsersEmail')
            $output.Add($Member) | Out-Null
        }
    }
    return $output
}

<#
.SYNOPSIS
Get Teams that have a single owner
.DESCRIPTION
This function returns a list of Teams that only have a single owner.
#>
Function Get-UcTeamsWithSingleOwner{
    Get-UcTeamUsersEmail -Role Owner -Confirm:$false | Group-Object -Property TeamDisplayName | Where-Object {$_.Count -lt 2} | Select-Object -ExpandProperty Group
}