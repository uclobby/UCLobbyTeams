function Get-UcTeamsVersion {
    param(
        [string]$Path,
        [string]$Computer,
        [System.Management.Automation.PSCredential]$Credential,
        [switch]$SkipModuleCheck
    )
    <#
        .SYNOPSIS
        Get Microsoft Teams Desktop Version

        .DESCRIPTION
        This function returns the installed Microsoft Teams desktop version for each user profile.

        .PARAMETER Path
        Specify the path with Teams Log Files

        .PARAMETER Computer
        Specify the remote computer

        .PARAMETER Credential
        Specify the credential to be used to connect to the remote computer

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

    $regexVersion = '("version":")([0-9.]*)'
    $regexRing = '("ring":")(\w*)'
    $regexEnv = '("environment":")(\w*)'
    $regexCloudEnv = '("cloudEnvironment":")(\w*)'
    
    $regexWindowsUser = '("upnWindowUserUpn":")([a-zA-Z0-9@._-]*)'
    $regexTeamsUserName = '("userName":")([a-zA-Z0-9@._-]*)'

    #20240309 - REGEX to get New Teams version from log file DesktopApp: Version: 23202.1500.2257.3700
    $regexNewVersion = '(DesktopApp: Version: )(\d{5}.\d{4}.\d{4}.\d{4})'

    $outTeamsVersion = [System.Collections.ArrayList]::new()
    if (!$SkipModuleCheck) {
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
    }

    if ($Path) {
        if (Test-Path $Path -ErrorAction SilentlyContinue) {
            #region Teams Classic Path
            $TeamsSettingsFiles = Get-ChildItem -Path $Path -Include "settings.json" -Recurse
            foreach ($TeamsSettingsFile in $TeamsSettingsFiles) {
                $TeamsSettings = Get-Content -Path $TeamsSettingsFile.FullName
                $Version = ""
                $Ring = ""
                $Env = ""
                $CloudEnv = ""
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
                }
                catch { }
                $TeamsDesktopSettingsFile = $TeamsSettingsFile.Directory.FullName + "\desktop-config.json"
                if (Test-Path $TeamsDesktopSettingsFile -ErrorAction SilentlyContinue) {
                    $TeamsDesktopSettings = Get-Content -Path $TeamsDesktopSettingsFile
                    $WindowsUser = ""
                    $TeamsUserName = ""
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
                    WindowsUser      = $WindowsUser
                    TeamsUser        = $TeamsUserName
                    Type             = "Teams Classic"
                    Version          = $Version
                    Ring             = $Ring
                    Environment      = $Env
                    CloudEnvironment = $CloudEnv
                    Path             = $TeamsSettingsFile.Directory.FullName
                }
                $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersionFromPath')
                $outTeamsVersion.Add($TeamsVersion) | Out-Null
            }
            #endregion
            #region New Teams Path
            $TeamsSettingsFiles = Get-ChildItem -Path $Path -Include "tma_settings.json" -Recurse
            foreach ($TeamsSettingsFile in $TeamsSettingsFiles) {
                if (Test-Path $TeamsSettingsFile -ErrorAction SilentlyContinue) {
                    $NewTeamsSettings = Get-Content -Path $TeamsSettingsFile | ConvertFrom-Json
                    $tmpAccountID = $NewTeamsSettings.primary_user.accounts.account_id
                    try {
                        $Version = ""
                        $MostRecentTeamsLogFile = Get-ChildItem -Path $TeamsSettingsFile.Directory.FullName -Include "MSTeams_*.log" -Recurse | Sort-Object -Property CreationTime -Descending | Select-Object -First 1
                        $TeamLogContents = Get-Content $MostRecentTeamsLogFile
                        $RegexTemp = [regex]::Match($TeamLogContents, $regexNewVersion).captures.groups
                        if ($RegexTemp.Count -ge 2) {
                            $Version = $RegexTemp[2].value
                        }
                    }
                    catch {}

                    $TeamsVersion = New-Object -TypeName PSObject -Property @{
                        WindowsUser      = "NA"
                        TeamsUser        = $NewTeamsSettings.primary_user.accounts.account_upn
                        Type             = "New Teams"
                        Version          = $Version
                        Ring             = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.ring
                        Environment      = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.environment
                        CloudEnvironment = $NewTeamsSettings.primary_user.accounts.cloud
                        Path             = $TeamsSettingsFile.Directory.FullName
                    }
                    $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersionFromPath')
                    [void]$outTeamsVersion.Add($TeamsVersion)
                }
            }
            #endregion
        }
        else {
            Write-Error -Message ("Invalid Path, please check if path: " + $path + " is correct and exists.")
        }
    }
    else {
        $currentDateFormat = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
        if ($Computer) {
            $RemotePath = "\\" + $Computer + "\C$\Users"
            $ComputerName = $Computer
            if ($Credential) {
                if ($Computer.IndexOf('.') -gt 0) {
                    $PSDriveName = $Computer.Substring(0, $Computer.IndexOf('.')) + "_TmpTeamsVersion"
                }
                else {
                    $PSDriveName = $Computer + "_TmpTeamsVersion"
                }
                New-PSDrive -Root $RemotePath -Name $PSDriveName -PSProvider FileSystem -Credential $Credential | Out-Null
            }

            if (Test-Path -Path $RemotePath) {
                $Profiles = Get-ChildItem -Path $RemotePath -ErrorAction SilentlyContinue
            }
            else {
                Write-Error -Message ("Error: Cannot get users on " + $computer + ", please check if name is correct and if the current user has permissions.")
            }
        }
        else {
            $ComputerName = $Env:COMPUTERNAME
            $Profiles = Get-childItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { Get-ItemProperty $_.pspath } | Where-Object { $_.fullprofile -eq 1 }
        }
       
        foreach ($UserProfile in $Profiles) {
            if ($Computer) {
                $ProfilePath = $UserProfile.FullName
                $ProfileName = $UserProfile.Name
            }
            else {
                $ProfilePath = $UserProfile.ProfileImagePath
                #20231013 Added exception handeling, only known case is when a windows profile was created when the machine was joined to a previous domain.
                try {
                    $ProfileName = (New-Object System.Security.Principal.SecurityIdentifier($UserProfile.PSChildName)).Translate( [System.Security.Principal.NTAccount]).Value
                }
                catch {
                    $ProfileName = "Unknown Windows User"
                }
            }
            #region classic teams
            $TeamsSettingPath = $ProfilePath + "\AppData\Roaming\Microsoft\Teams\settings.json"
            if (Test-Path $TeamsSettingPath -ErrorAction SilentlyContinue) {
                $TeamsSettings = Get-Content -Path $TeamsSettingPath
                $Version = ""
                $Ring = ""
                $Env = ""
                $CloudEnv = ""
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
                }
                catch { }
                $TeamsApp = $ProfilePath + "\AppData\Local\Microsoft\Teams\current\Teams.exe"
                $TeamsInstallTimePath = $ProfilePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt"
                #20240228 - In some cases the install file can be missing.
                $tmpInstallDate = ""
                if (Test-Path $TeamsInstallTimePath -ErrorAction SilentlyContinue) {
                    $InstallDateStr = Get-Content ($ProfilePath + "\AppData\Roaming\Microsoft\Teams\installTime.txt")
                    $tmpInstallDate = [Datetime]::ParseExact($InstallDateStr, 'M/d/yyyy', $null) | Get-Date -Format $currentDateFormat
                }
                
                $TeamsVersion = New-Object -TypeName PSObject -Property @{
                    Computer         = $ComputerName
                    Profile          = $ProfileName
                    ProfilePath      = $ProfilePath
                    Type             = "Teams Classic"
                    Version          = $Version
                    Ring             = $Ring
                    Environment      = $Env
                    CloudEnvironment = $CloudEnv
                    Arch             = Get-UcArch $TeamsApp
                    InstallDate      = $tmpInstallDate
                }
                $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersion')
                [void]$outTeamsVersion.Add($TeamsVersion)
            }
            #endregion

            #region New Teams
            $NewTeamsSettingPath = $ProfilePath + "\AppData\Local\Publishers\8wekyb3d8bbwe\TeamsSharedConfig\tma_settings.json"
            if (Test-Path $NewTeamsSettingPath -ErrorAction SilentlyContinue) {
                $NewTeamsSettings = Get-Content -Path $NewTeamsSettingPath | ConvertFrom-Json
                $tmpAccountID = $NewTeamsSettings.primary_user.accounts.account_id
                if ($Computer) {
                    $newTeamsLocation = Get-ChildItem -Path ( $RemotePath + "\..\Program Files\Windowsapps" ) -Filter "ms-teams.exe" -Recurse -Depth 1 | Sort-Object -Property CreationTime -Descending | Select-Object -First 1                    
                }
                else {
                    #20240103 - Using Get-AppPackage drops the requirement to run with Administrative Rights
                    $newTeamsInstallPath = (Get-AppPackage MSTeams).InstallLocation + ".\ms-teams.exe"
                    $newTeamsLocation = Get-ItemProperty -Path ($newTeamsInstallPath)
                }
                if (Test-Path -Path $newTeamsLocation.FullName -ErrorAction SilentlyContinue) {
                    $TeamsVersion = New-Object -TypeName PSObject -Property @{
                        Computer         = $ComputerName
                        Profile          = $ProfileName
                        ProfilePath      = $ProfilePath
                        Type             = "New Teams"
                        Version          = $newTeamsLocation.VersionInfo.ProductVersion
                        Ring             = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.ring
                        Environment      = $NewTeamsSettings.tma_ecs_settings.$tmpAccountID.environment
                        CloudEnvironment = $NewTeamsSettings.primary_user.accounts.cloud
                        Arch             = Get-UcArch $newTeamsLocation.FullName
                        InstallDate      = $newTeamsLocation.CreationTime | Get-Date -Format $currentDateFormat
                    }
                    $TeamsVersion.PSObject.TypeNames.Insert(0, 'TeamsVersion')
                    [void]$outTeamsVersion.Add($TeamsVersion)
                }
            }
            #endregion
        }
        if ($Credential -and $PSDriveName) {
            try {
                Remove-PSDrive -Name $PSDriveName -ErrorAction SilentlyContinue
            }
            catch {}
        }
    }
    return $outTeamsVersion
}