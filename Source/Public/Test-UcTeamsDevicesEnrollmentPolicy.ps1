<#
.SYNOPSIS
Validate Intune Enrollment Policies that are supported by Microsoft Teams Android Devices

.DESCRIPTION
This function will validate each setting in a Conditional Access Policy to make sure they are in line with the supported settings:

Contributors: David Paulino, Gonçalo Sepulveda

Requirements: Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)

.PARAMETER UserUPN
Specifies a UserUPN that we want to check for a user enrollment policies.

.PARAMETER Detailed
Displays test results for unsupported settings in each Intune Enrollment Policy.

.PARAMETER ExportCSV
When present will export the detailed results to a CSV file. By defautl will save the file under the current user downloads, unless we specify the OutputPath.

.PARAMETER OutputPath
Allows to specify the path where we want to save the results.

.EXAMPLE 
PS> Test-UcTeamsDevicesEnrollmentPolicy

.EXAMPLE 
PS> Test-UcTeamsDevicesEnrollmentPolicy -UserUPN

#>

Function Test-UcTeamsDevicesEnrollmentPolicy {
    Param(
        [string]$UserUPN,
        [switch]$Detailed,
        [switch]$ExportCSV,
        [string]$OutputPath
    )

    $GraphURI_Users = "https://graph.microsoft.com/v1.0/users"
    $GraphURI_EnrollmentPolicies = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations"
    $output = [System.Collections.ArrayList]::new()

    if (Test-UcMgGraphConnection -Scopes "DeviceManagementServiceConfig.Read.All", "Directory.Read.All") {
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
        $outFileName = "TeamsDevices_EnrollmentPolicy_Report_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
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
        try {
            Write-Progress -Activity "Test-UcTeamsDevicesEnrollmentPolicy" -Status "Getting Intune Enrollment Policies"
            $EnrollmentPolicies = (Invoke-MgGraphRequest -Uri $GraphURI_EnrollmentPolicies -Method GET).value
            $connectedMSGraph = $true
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
            Write-Error 'Please connect to MS Graph with Connect-MgGraph -Scopes "DeviceManagementServiceConfig.Read.All","Directory.Read.All" before running this script'
        }
        if ($connectedMSGraph) {
            if ($UserUPN) {
                try {
                    $UserGroups = (Invoke-MgGraphRequest -Uri ($GraphURI_Users + "/" + $userUPN + "/transitiveMemberOf?`$select=id") -Method GET).value.id
                }
                catch [System.Net.Http.HttpRequestException] {
                    if ($PSItem.Exception.Response.StatusCode -eq "NotFound") {
                        Write-warning -Message ("User Not Found: " + $UserUPN)
                    }
                    return
                }
            }
            #We need to cycle this in order to get the Group IDs that are assigned to the enrollment policies
            $graphRequests = [System.Collections.ArrayList]::new()
            foreach ($EnrollmentPolicy in $EnrollmentPolicies) {
                #Only Android Policies
                if(($EnrollmentPolicy."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration") -and ($EnrollmentPolicy.platformType -eq "android") ){
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id =  $EnrollmentPolicy.id
                        method = "GET"
                        url = "/deviceManagement/deviceEnrollmentConfigurations/" + $EnrollmentPolicy.id + "/assignments?`$select=target"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }

            $PolicyGroupAssigment = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Test-UcTeamsDevicesEnrollmentPolicy, fetching enrollment policies assignment" -IncludeBody)
        
            $graphRequests = [System.Collections.ArrayList]::new()
            foreach($Group in $PolicyGroupAssigment.body.value.target.GroupID){
                if(!($UserUPN) -or ($PolicyGroupAssigment.body.value.target.GroupID -in $UserGroups)){
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id =  $Group
                        method = "GET"
                        url = "/groups/" + $Group + "?`$select=id,displayName"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }
            if($graphRequests.Count -gt 0){
                $Groups = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Test-UcTeamsDevicesEnrollmentPolicy, getting group information") | Select-Object Id,displayName
            }
            
            foreach ($EnrollmentPolicy in $EnrollmentPolicies) {
                $Status = "Not Supported"
                #This is the default enrollment policy that is applied to all users/devices
                if($EnrollmentPolicy."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionsConfiguration"){
                    if(!($EnrollmentPolicy.androidRestriction.platformBlocked)) {
                        if(!($EnrollmentPolicy.androidRestriction.personalDeviceEnrollmentBlocked)){
                            $Status = "Supported"
                        }
                    }
                    $SettingPSObj = [PSCustomObject]@{
                        PID                             = 9999
                        PolicyName                      = $EnrollmentPolicy.displayName
                        PolicyPriority                  = "Default"
                        PlatformType                    = "Android device administrator"
                        AssignedToGroup                 = "All devices"
                        TeamsDevicesStatus              = $Status
                        PlatformBlocked                 = $EnrollmentPolicy.androidRestriction.platformBlocked
                        PersonalDeviceEnrollmentBlocked = $EnrollmentPolicy.androidRestriction.personalDeviceEnrollmentBlocked
                        osMinimumVersion                = $EnrollmentPolicy.androidRestriction.osMinimumVersion
                        osMaximumVersion                = $EnrollmentPolicy.androidRestriction.osMaximumVersion
                        blockedManufacturers            = $EnrollmentPolicy.androidRestriction.blockedManufacturers
                    }
                    $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentPolicy')
                    [void]$output.Add($SettingPSObj)
                }

                $Status = "Not Supported"
                if(($EnrollmentPolicy."@odata.type" -eq "#microsoft.graph.deviceEnrollmentPlatformRestrictionConfiguration") -and ($EnrollmentPolicy.platformType -eq "android") ){
                    $AssignedToGroup = [System.Collections.ArrayList]::new()
                    $AssignedGroupsTemp = ($PolicyGroupAssigment | Where-Object -Property "id" -Value $EnrollmentPolicy.id -EQ).body.value.target.GroupID

                    foreach($AssignedGroup in $AssignedGroupsTemp){
                        if(!($UserUPN) -or ($AssignedGroup -in $UserGroups)){
                            $GroupDisplayName = ($Groups | Where-Object -Property "id" -Value $AssignedGroup -EQ).DisplayName
                            $GroupEntry = New-Object -TypeName PSObject -Property @{
                                GroupID          = $AssignedGroup
                                GroupDisplayName = $GroupDisplayName
                            }
                            [void]$AssignedToGroup.Add($GroupEntry)
                        }
                    }
                    if($AssignedToGroup.Count -gt 0){
                        $outAssignedToGroup = "None"
                        if ($AssignedToGroup.count -eq 1){
                            $outAssignedToGroup = $AssignedToGroup[0].GroupDisplayName
                        }
                        elseif ($AssignedToGroup.count -gt 1) {
                            $outAssignedToGroup = "" + $AssignedToGroup.count + " group(s)"
                        }
                        if(!($EnrollmentPolicy.platformRestriction.platformBlocked)) {
                            if(!($EnrollmentPolicy.platformRestriction.personalDeviceEnrollmentBlocked)){
                                $Status = "Supported"
                            }
                        }

                        $SettingPSObj = [PSCustomObject]@{
                            PID                             = $EnrollmentPolicy.priority
                            PolicyName                      = $EnrollmentPolicy.displayName
                            PolicyPriority                  = $EnrollmentPolicy.priority
                            PlatformType                    = "Android device administrator"
                            AssignedToGroup                 = $outAssignedToGroup
                            AssignedToGroupList             = $AssignedToGroup
                            TeamsDevicesStatus              = $Status 
                            PlatformBlocked                 = $EnrollmentPolicy.platformRestriction.platformBlocked
                            PersonalDeviceEnrollmentBlocked = $EnrollmentPolicy.platformRestriction.personalDeviceEnrollmentBlocked
                            osMinimumVersion                = $EnrollmentPolicy.platformRestriction.osMinimumVersion
                            osMaximumVersion                = $EnrollmentPolicy.platformRestriction.osMaximumVersion
                            blockedManufacturers            = $EnrollmentPolicy.platformRestriction.blockedManufacturers
                        }
                        $SettingPSObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceEnrollmentPolicy')
                        [void]$output.Add($SettingPSObj)
                    }
                }
            }
            if($Detailed){
                if ($ExportCSV) {
                    $output |Sort-Object PID | Select-Object PolicyName, PolicyPriority, PlatformType, AssignedToGroup, TeamsDevicesStatus, PlatformBlocked, PersonalDeviceEnrollmentBlocked, osMinimumVersion, osMaximumVersion, blockedManufacturers | Export-Csv -path $OutputFullPath -NoTypeInformation
                    Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
                    return
                }
                #20231116 - Fix for empty output.
                else {
                    $output | Sort-Object PID | Format-List
                }

            } else {
                $output | Sort-Object PID | Format-Table
            }
        }
    }
}

