function Test-UcTeamsDevicesEnrollmentProfile {
    <#
        .SYNOPSIS
        Validate Intune Enrollment Profiles that are supported by Microsoft Teams Android Devices

        .DESCRIPTION
        This function will validate if the settings in a Enrollment Profile are supported:

        Contributors: David Paulino, Gonçalo Sepúlveda and Eileen Beato

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)
                        or
                        Microsoft Graph Authentication PowerShell Module (Install-Module Microsoft.Graph.Authentication)

                        Microsoft Graph Scopes:
                            "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All", "Directory.Read.All"

        .PARAMETER UserUPN
        Specifies a UserUPN that we want to check for a user enrollment profiles.

        .PARAMETER PlatformType
        Platform Type:
            AndroidDeviceAdministrator
            AOSP


        .PARAMETER Detailed
        Displays test results for unsupported settings in each Intune Enrollment Profile.

        .PARAMETER ExportCSV
        When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

        .PARAMETER OutputPath
        Allows to specify the path where we want to save the results.

        .EXAMPLE 
        PS> Test-UcTeamsDevicesEnrollmentProfile

        .EXAMPLE 
        PS> Test-UcTeamsDevicesEnrollmentProfile -UserUPN
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [string]$UserUPN,
        [ValidateSet("AndroidDeviceAdministrator", "AOSP", "Any")]
        [string]$PlatformType = "Any",
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath
    )   
    
    $GraphPathUsers = "/users"
    $GraphPathGroups = "/groups"
    $GraphPathEnrollmentProfiles = "/deviceManagement/deviceEnrollmentConfigurations/?`$expand=assignments"
    $GraphPathAOSPEnrollmentProfiles = "/deviceManagement/androidDeviceOwnerEnrollmentProfiles?`$select=displayName,tokenExpirationDateTime&`$filter=enrollmentMode eq 'corporateOwnedAOSPUserAssociatedDevice' and isTeamsDeviceProfile eq true" 
    $output = [System.Collections.ArrayList]::new()

    #region Graph Connection, Scope validation and module version
    if (!(Test-UcServiceConnection -Type MSGraph -Scopes "DeviceManagementServiceConfig.Read.All", "DeviceManagementConfiguration.Read.All", "Directory.Read.All")) {
        return
    }
    #2025-01-31: Only need to check this once per PowerShell session
    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }
    #endregion

    #region ExportCSV defaults to Detaileld and Output Path Test
    if ($ExportCSV) {
        $Detailed = $true
        $outFileName = "TeamsDevices_EnrollmentProfile_Report_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
        if ($OutputPath) {
            if (!(Test-Path $OutputPath -PathType Container)) {
                Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
                return
            } 
            $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $outFileName)
        }
        else {                
            $OutputFullPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFileName)
        }
    }
    #endregion
    try {
        Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Getting Intune Enrollment Policies"
        $EnrollmentProfiles = Invoke-UcGraphRequest -Path $GraphPathEnrollmentProfiles -Beta
        if ($PlatformType -in ("Any", "AOSP") ) {
            Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Getting AOSP Enrollment Profiles"
            $AOSPEnrollmentProfiles = Invoke-UcGraphRequest -Path $GraphPathAOSPEnrollmentProfiles -Beta
        }
    }
    catch [System.Net.Http.HttpRequestException] {
        if ($PSItem.Exception.Response.StatusCode -eq "Unauthorized") {
            Write-Error "Access Denied, please make sure the user connecing to MS Graph is part of one of the following Global Reader/Intune Service Administrator/Global Administrator roles"
        }
        else {
            Write-Error $PSItem.Exception.Message
        }
    }
    catch {
        Write-Error $PSItem.Exception.Message
    }

    if ($UserUPN) {
        try {
            Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentProfile" -Status "Retriving User Group Membership"
            $UserGroups = (Invoke-UcGraphRequest -Path ($GraphPathUsers + "/" + $UserUPN + "/transitiveMemberOf?`$select=id")).id
        }
        catch [System.Net.Http.HttpRequestException] {
            if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                Write-warning -Message ("User Not Found: " + $UserUPN)
            }
            return
        }
    }
                 
    foreach ($EnrollmentProfile in $EnrollmentProfiles) {
        $Status = "Not Supported"
        #This is the default enrollment policy that is applied to all users/devices
        if ($EnrollmentProfile."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration") {
            if ($PlatformType -in ("Any", "AndroidDeviceAdministrator")) {
                if (!($EnrollmentProfile.androidRestriction.platformBlocked) -and !($EnrollmentProfile.androidRestriction.personalDeviceEnrollmentBlocked)) {
                    $Status = "Supported"
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                              = 9999
                    EnrollmentProfileName           = $EnrollmentProfile.displayName
                    Priority                        = "Default"
                    PlatformType                    = "Android device administrator"
                    AssignedToGroup                 = "All devices"
                    TeamsDevicesStatus              = $Status
                    PlatformBlocked                 = $EnrollmentProfile.androidRestriction.platformBlocked
                    PersonalDeviceEnrollmentBlocked = $EnrollmentProfile.androidRestriction.personalDeviceEnrollmentBlocked
                    osMinimumVersion                = $EnrollmentProfile.androidRestriction.osMinimumVersion
                    osMaximumVersion                = $EnrollmentProfile.androidRestriction.osMaximumVersion
                    blockedManufacturers            = $EnrollmentProfile.androidRestriction.blockedManufacturers
                    TokenExpirationDate             = ""
                }
                $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
                [void]$output.Add($SettingPSObj)
            }
        }

        $Status = "Not Supported"
        if (($EnrollmentProfile."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration") -and ($EnrollmentProfile.platformType -eq "android") ) {
            $AssignedToGroup = [System.Collections.ArrayList]::new()

            if ($UserUPN -and !($EnrollmentProfile.assignments.target.groupId | Where-Object { $_ -in $UserGroups })) {
                #Skipping this policy since not assigned to a user
                Write-Verbose ("Skipping Enrollment Profile " + $EnrollmentProfile.displayName + " because user $UserUPN is not part of the Included Groups.")
                continue
            }

            #2025-03-28: Moving the Group Name here since we use expand Assignments in the graph request.
            foreach ($AssignedGroupID in $EnrollmentProfile.assignments.target.groupId) {
                $GroupInfo = Invoke-UcGraphRequest -Path "$GraphPathGroups/$AssignedGroupID/?`$select=id,displayname"
                $GroupEntry = [PSCustomObject][Ordered]@{
                    GroupID          = $AssignedGroupID
                    GroupDisplayName = $GroupInfo.displayname
                }
                [void]$AssignedToGroup.Add($GroupEntry)
            }

            if ($AssignedToGroup.Count -gt 0) {
                $outAssignedToGroup = "None"
                if ($AssignedToGroup.count -eq 1) {
                    $outAssignedToGroup = $AssignedToGroup[0].GroupDisplayName
                }
                elseif ($AssignedToGroup.count -gt 1) {
                    $outAssignedToGroup = "" + $AssignedToGroup.count + " group(s)"
                }
                if (!($EnrollmentProfile.platformRestriction.platformBlocked) -and !($EnrollmentProfile.platformRestriction.personalDeviceEnrollmentBlocked)) {
                    $Status = "Supported"
                }
                $SettingPSObj = [PSCustomObject][Ordered]@{
                    ID                              = $EnrollmentProfile.priority
                    EnrollmentProfileName           = $EnrollmentProfile.displayName
                    Priority                        = $EnrollmentProfile.priority
                    PlatformType                    = "Android device administrator"
                    AssignedToGroup                 = $outAssignedToGroup
                    AssignedToGroupList             = $AssignedToGroup
                    TeamsDevicesStatus              = $Status 
                    PlatformBlocked                 = $EnrollmentProfile.platformRestriction.platformBlocked
                    PersonalDeviceEnrollmentBlocked = $EnrollmentProfile.platformRestriction.personalDeviceEnrollmentBlocked
                    osMinimumVersion                = $EnrollmentProfile.platformRestriction.osMinimumVersion
                    osMaximumVersion                = $EnrollmentProfile.platformRestriction.osMaximumVersion
                    blockedManufacturers            = $EnrollmentProfile.platformRestriction.blockedManufacturers
                    TokenExpirationDate             = ""
                }
                $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
                [void]$output.Add($SettingPSObj)
            }
        }
    }
    #region 2024-09-12: Support for Android AOSP Enrollment
    if ($PlatformType -in ("Any", "AOSP") ) {
        $PolicyName = ""
        $TokenExpirationDate = ""
        $CurrentDate = Get-Date
        if ($AOSPEnrollmentProfiles.Length -ge 1) {
            #In some cases we can have multiple AOSP profiles enabled for Teams Devices, currently we only support one valid at a time.
            $ValidAOSPEnrollmentProfiles = $AOSPEnrollmentProfiles | Where-Object { $_.tokenExpirationDateTime -gt $CurrentDate }
            if ($ValidAOSPEnrollmentProfiles.Length -eq 1) {
                $PolicyName = $ValidAOSPEnrollmentProfiles.displayName
                $TokenExpirationDate = $ValidAOSPEnrollmentProfiles.tokenExpirationDateTime
                $TeamsDevicesStatus = "Supported - Token valid for " + ($TokenExpirationDate - $CurrentDate).days + " day(s)"
            }
            elseif ($ValidAOSPEnrollmentProfiles.Length -eq 0) {
                $ExpiredAOSPEnrollmentProfiles = $AOSPEnrollmentProfiles | Sort-Object -Property tokenExpirationDateTime -Descending
                $TeamsDevicesStatus = "Not Supported - Token Expired on " + ($ExpiredAOSPEnrollmentProfiles[0].tokenExpirationDateTime.ToString([cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern))
            }
            else {
                $TeamsDevicesStatus = "Not Supported - Multiple AOSP Enrollment Profile enabled for Teams Devices"
            }
        }
        else {
            $TeamsDevicesStatus = "Not Supported - Missing AOSP Enrollment"
        }
        $SettingPSObj = [PSCustomObject][Ordered]@{
            ID                              = "10000"
            EnrollmentProfileName           = $PolicyName
            Priority                        = "ASOP"
            PlatformType                    = "Android Open Source Project (AOSP)"
            TeamsDevicesStatus              = $TeamsDevicesStatus
            PlatformBlocked                 = ""
            PersonalDeviceEnrollmentBlocked = ""
            osMinimumVersion                = ""
            osMaximumVersion                = ""
            blockedManufacturers            = ""
            TokenExpirationDate             = $TokenExpirationDate
        }
        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentProfile')
        [void]$output.Add($SettingPSObj)
    }
    #endregion
    #region Output
    if ($ExportCSV) {
        $output | Sort-Object ID | Select-Object -ExcludeProperty ID | Export-Csv -path $OutputFullPath -NoTypeInformation
        Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        return
    }
    if ($Detailed) {
        return $output | Sort-Object ID | Select-Object -ExcludeProperty ID 
    }
    else {
        return $output | Sort-Object ID
    }
    #endregion
}