
<#
.SYNOPSIS
Get Microsoft Teams Desktop Version

.DESCRIPTION
This function returns the installed Microsoft Teams desktop version for each user profile.

.PARAMETER $Path
Specify the path with Teams Log Files

.PARAMETER $Computer
Specify the remote computer

.EXAMPLE
PS> Get-UcTeamsVersion

.EXAMPLE
PS> Get-UcTeamsVersion -Path C:\Temp\

.EXAMPLE
PS> Get-UcTeamsVersion -Computer workstation124

.EXAMPLE
PS> $cred = Get-Credential
PS> Get-UcTeamsVersion -Computer workstation124 -Credential $cred

#>
Function Get-UcTeamsVersion {
    Param(
        [string]$Path,
        [string]$Computer,
        [System.Management.Automation.PSCredential]$Credential
    )
    $regexVersion = '("version":")([0-9.]*)'
    $regexRing = '("ring":")(\w*)'
    $regexEnv = '("environment":")(\w*)'
    $regexCloudEnv = '("cloudEnvironment":")(\w*)'
    $regexRegion = '("region":")([a-zA-Z0-9._-]*)'

    $regexWindowsUser = '("upnWindowUserUpn":")([a-zA-Z0-9@._-]*)'
    $regexTeamsUserName = '("userName":")([a-zA-Z0-9@._-]*)'

    $outTeamsVersion = [System.Collections.ArrayList]::new()

    if($Path){
        if (Test-Path $Path -ErrorAction SilentlyContinue) {
            $TeamsSettingsFiles = Get-ChildItem -Path $Path -Include "settings.json" -Recurse
            foreach ($TeamsSettingsFile in $TeamsSettingsFiles){
                $TeamsSettings = Get-Content -Path $TeamsSettingsFile.FullName
                $Version = ""
                $Ring = ""
                $Env = ""
                $CloudEnv = ""
                $Region = ""
                try {
                    $VersionTemp = [regex]::Match($TeamsSettings, $regexVersion).captures.groups
                    if ($VersionTemp.Count -ge 2) {
                        $Version = $VersionTemp[2].value
                    }
                    $RingTemp = [regex]::Match($TeamsSettings, $regexRing).captures.groups
                    if ($RingTemp.Count -ge 2) {
                        $Ring = $RingTemp[2].value
                    }
                    $EnvTemp = [regex]::Match($TeamsSettings, $regexEnv).captures.groups
                    if ($EnvTemp.Count -ge 2) {
                        $Env = $EnvTemp[2].value
                    }
                    $CloudEnvTemp = [regex]::Match($TeamsSettings, $regexCloudEnv).captures.groups
                    if ($CloudEnvTemp.Count -ge 2) {
                        $CloudEnv = $CloudEnvTemp[2].value
                    }
                    $RegionTemp = [regex]::Match($TeamsSettings, $regexRegion).captures.groups
                    if ($RegionTemp.Count -ge 2) {
                        $Region = $RegionTemp[2].value
                    }
                }
                catch { }
                $TeamsDesktopSettingsFile = $TeamsSettingsFile.Directory.FullName + "\desktop-config.json"
                if (Test-Path $TeamsDesktopSettingsFile -ErrorAction SilentlyContinue) {
                    $TeamsDesktopSettings = Get-Content -Path $TeamsDesktopSettingsFile
                    $WindowsUser = ""
                    $TeamsUserName =""
                    $RegexTemp = [regex]::Match($TeamsDesktopSettings, $regexWindowsUser).captures.groups
                    if ($RegexTemp.Count -ge 2) {
                        $WindowsUser = $RegexTemp[2].value
                    }
                    $RegexTemp = [regex]::Match($TeamsDesktopSettings, $regexTeamsUserName).captures.groups
                    if ($RegexTemp.Count -ge 2) {
                        $TeamsUserName = $RegexTemp[2].value
                    }
                }
                $TeamsVersion = New-Object -TypeName PSObject -Property @{
                    WindowsUser             = $WindowsUser
                    TeamsUser               = $TeamsUserName
                    Version                 = $Version
                    Ring                    = $Ring
                    Environment             = $Env
                    CloudEnvironment        = $CloudEnv
                    Region                  = $Region
                    Path                    = $TeamsSettingsFile.Directory.FullName
                }
                $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersionFromPath')
                $outTeamsVersion.Add($TeamsVersion) | Out-Null
            }
        } else {
            Write-Error -Message ("Invalid Path, please check if path: " + $path + " is correct and exists.")
        }

    } else {
        $currentDateFormat = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
        if($Computer) {
            $RemotePath = "\\" + $Computer + "\C$\Users"

            if($Credential){
                New-PSDrive -Root $RemotePath -Name ($Computer+"TmpTeamsVersion") -PSProvider FileSystem -Credential $Credential | Out-Null
            }

            if(Test-Path -Path $RemotePath) {
                $Profiles = Get-ChildItem -Path $RemotePath -ErrorAction SilentlyContinue
            } else {
                Write-Error -Message ("Cannot get users on : " + $computer + " check if name is correct and if the current user has permissions.")
            }
        } else {
            $Profiles = Get-childItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Get-ItemProperty $_.pspath } | Where-Object { $_.fullprofile -eq 1 }
        }
        
        foreach ($UserProfile in $Profiles) {
            if($Computer){
                $ProfilePath = $UserProfile.FullName
                $ProfileName = $UserProfile.Name
            } else {
                $ProfilePath = $UserProfile.ProfileImagePath
                $ProfileName  = (New-Object System.Security.Principal.SecurityIdentifier($UserProfile.PSChildName)).Translate( [System.Security.Principal.NTAccount]).Value
            }
            $TeamsSettingPath = $ProfilePath + "\AppData\Roaming\Microsoft\Teams\settings.json"
            if (Test-Path $TeamsSettingPath -ErrorAction SilentlyContinue) {
                $TeamsSettings = Get-Content -Path $TeamsSettingPath
                $Version = ""
                $Ring = ""
                $Env = ""
                $CloudEnv = ""
                $Region = ""
                try {
                    $VersionTemp = [regex]::Match($TeamsSettings, $regexVersion).captures.groups
                    if ($VersionTemp.Count -ge 2) {
                        $Version = $VersionTemp[2].value
                    }
                    $RingTemp = [regex]::Match($TeamsSettings, $regexRing).captures.groups
                    if ($RingTemp.Count -ge 2) {
                        $Ring = $RingTemp[2].value
                    }
                    $EnvTemp = [regex]::Match($TeamsSettings, $regexEnv).captures.groups
                    if ($EnvTemp.Count -ge 2) {
                        $Env = $EnvTemp[2].value
                    }
                    $CloudEnvTemp = [regex]::Match($TeamsSettings, $regexCloudEnv).captures.groups
                    if ($CloudEnvTemp.Count -ge 2) {
                        $CloudEnv = $CloudEnvTemp[2].value
                    }
                    $RegionTemp = [regex]::Match($TeamsSettings, $regexRegion).captures.groups
                    if ($RegionTemp.Count -ge 2) {
                        $Region = $RegionTemp[2].value
                    }
                }
                catch { }
                $TeamsApp = $ProfilePath + "\AppData\Local\Microsoft\Teams\current\Teams.exe"
                $InstallDateStr = Get-Content ($ProfilePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt")
                $TeamsVersion = New-Object -TypeName PSObject -Property @{
                    Profile          = $ProfileName
                    ProfilePath      = $ProfilePath
                    Version          = $Version
                    Ring             = $Ring
                    Environment      = $Env
                    CloudEnvironment = $CloudEnv
                    Region           = $Region
                    Arch             = Get-UcArch $TeamsApp
                    InstallDate      = [Datetime]::ParseExact($InstallDateStr, 'M/d/yyyy', $null) | Get-Date -Format $currentDateFormat
                }
                $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersion')
                $outTeamsVersion.Add($TeamsVersion) | Out-Null
            }
        }
        if($Credential -and $Computer){
            try{
                Remove-PSDrive -Name ($Computer+"TmpTeamsVersion") -ErrorAction SilentlyContinue
            } catch {}
        }
    }
    return $outTeamsVersion
}