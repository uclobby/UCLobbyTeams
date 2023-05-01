<#
.SYNOPSIS
Update Microsoft Teams Devices

.DESCRIPTION
This function will send update commands to Teams Android Devices using MS Graph.

Contributors: Eileen Beato, David Paulino and Bryan Kendrick

Requirements:   Microsoft Graph PowerShell Module (Install-Module Microsoft.Graph)
                Microsoft Graph Scopes:
                        "TeamworkDevice.ReadWrite.All"
.PARAMETER DeviceID
Specify the Teams Admin Center Device ID that we want to update.

.PARAMETER DeviceType
Specifies a filter, valid options:
    Phone - Teams Native Phones
    MTRA - Microsoft Teams Room Running Android
    Display - Microsoft Teams Displays 
    Panel - Microsoft Teams Panels

.PARAMETER UpdateType
Allow to specify which time of update we want to do:
		Firmware
        TeamsApp
        All

.PARAMETER InputCSV
When present will use this file as Input, we only need a column with Device Id. It supports files exported from Teams Admin Center (TAC).

.PARAMETER Subnet
Only available when using InputCSV and requires a “IP Address” column, it allows to only send updates to Teams Android devices within a subnet.
Format examples:
10.0.0.0/8
192.168.0.0/24


.PARAMETER OutputPath
Allows to specify the path where we want to save the results. By default will save on current user Download.

.PARAMETER ReportOnly
Will read Teams Device Android versions info and generate a report

.EXAMPLE 
PS> Update-UcTeamsDevice

.EXAMPLE 
PS> Update-UcTeamsDevice -ReportOnly

.EXAMPLE
PS> Update-UcTeamsDevice -InputCSV C:\Temp\DevicesList_2023-04-20_15-19-00-UTC.csv

#>

Function Update-UcTeamsDevice {
    [cmdletbinding(SupportsShouldProcess)]
    Param(
        [ValidateSet("Firmware", "TeamsApp", "All")]
        [string]$UpdateType = "All",
        [ValidateSet("Phone", "MTRA", "Display", "Panel")]
        [string]$DeviceType,
        [string]$DeviceID,
        [string]$InputCSV,
        [string]$Subnet,
        [string]$OutputPath,
        [switch]$ReportOnly
    )
    $regExIPAddressSubnet = "^((25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9]))\/(3[0-2]|[1-2]{1}[0-9]{1}|[1-9])$"

    if (Test-UcMgGraphConnection -Scopes "TeamworkDevice.ReadWrite.All") {
        $outTeamsDevices = [System.Collections.ArrayList]::new()
        Test-UcModuleUpdateAvailable -ModuleName UcLobbyTeams
        #Checking if the Subnet is valid
        if($Subnet){
            if(!($Subnet -match $regExIPAddressSubnet)){
                Write-Host ("Error: Subnet " + $Subnet + " is invalid, please make sure the subnet is valid and in this format 10.0.0.0/8, 192.168.0.0/24") -ForegroundColor Red
                return
            } 
        }

        if ($ReportOnly) {
            $outFileName = "UpdateTeamsDevices_ReportOnly_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
            $StatusType = "offline", "critical", "nonUrgent", "healthy"
        }
        else {
            $outFileName = "UpdateTeamsDevices_" + ( get-date ).ToString('yyyyMMdd-HHmmss') + ".csv"
            $StatusType = "critical", "nonUrgent"
        }
        #Verify if the Output Path exists
        if ($OutputPath) {
            if (!(Test-Path $OutputPath -PathType Container)) {
                Write-Host ("Error: Invalid folder " + $OutputPath) -ForegroundColor Red
                return
            }
            else {
                $OutputFullPath = [System.IO.Path]::Combine($OutputPath, $outFileName)
            }
        }
        else {                
            $OutputFullPath = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads", $outFileName)
        }

        $graphRequests = [System.Collections.ArrayList]::new()
        if ($DeviceID) {
            $gRequestTmp = New-Object -TypeName PSObject -Property @{
                id     = $DeviceID
                method = "GET"
                url    = "/teamwork/devices/" + $DeviceID
            }
            [void]$graphRequests.Add($gRequestTmp)
            $GraphResponse = Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Update-UcTeamsDevices, getting device info" -IncludeBody
            
            if ($GraphResponse.status -eq 200) {
                $TeamsDeviceList = $GraphResponse.body
            }
            elseif ($GraphResponse.status -eq 404) {
                Write-Host ("Error: Device ID $DeviceID not found.") -ForegroundColor Red
                return
            }
        }
        elseif ($InputCSV) {
            if (Test-Path $InputCSV) {
                try {
                    $TeamsDeviceInput = Import-Csv -Path $InputCSV
                }
                catch {
                    Write-Host ("Invalid CSV input file: " + $InputCSV) -ForegroundColor Red
                    return
                }
                foreach ($TeamsDevice in $TeamsDeviceInput) {
                    $includeDevice = $true
                    if ($Subnet) {
                        $includeDevice = Test-UcIPaddressInSubnet -IPAddress $TeamsDevice.'IP Address' -Subnet $Subnet
                    }

                    if ($includeDevice) {
                        $gRequestTmp = New-Object -TypeName PSObject -Property @{
                            id     = $TeamsDevice.'Device Id'
                            method = "GET"
                            url    = "/teamwork/devices/" + $TeamsDevice.'Device Id'
                        }
                        [void]$graphRequests.Add($gRequestTmp)
                    }
                }
                if ($graphRequests.Count -gt 0) {
                    $TeamsDeviceList = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Update-UcTeamsDevices, getting device info" )
                } 
            }
            else {
                Write-Host ("Error: File not found " + $InputCSV) -ForegroundColor Red
                return
            }
        }
        else {
            #Currently only Android based Teams devices are supported.
            switch ($DeviceType) {
                "Phone" { 
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "ipPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "lowCostPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "MTRA" {            
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "collaborationBar"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "touchConsole"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "Display" {
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "teamsDisplay"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                "Panel" {
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "teamsPanel"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
                Default {
                    #This is the only way to exclude MTRW and SurfaceHub by creating a request per device type.
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "ipPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'ipPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp) 
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "lowCostPhone"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'lowCostPhone'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "collaborationBar"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'collaborationBar'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "touchConsole"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'touchConsole'"
                    }
                    [void]$graphRequests.Add($gRequestTmp) 
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "teamsDisplay"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsDisplay'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                    $gRequestTmp = New-Object -TypeName PSObject -Property @{
                        id     = "teamsPanel"
                        method = "GET"
                        url    = "/teamwork/devices/?`$filter=deviceType eq 'teamsPanel'"
                    }
                    [void]$graphRequests.Add($gRequestTmp)
                }
            }
            #Using new cmdlet to get a list of devices
            $TeamsDeviceList = (Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Update-UcTeamsDevices, getting device info").value
        }

        $devicesWithUpdatePending = 0
        $graphRequests = [System.Collections.ArrayList]::new()
        foreach ($TeamsDevice in $TeamsDeviceList) {
            if ($TeamsDevice.healthStatus -in $StatusType) {
                $devicesWithUpdatePending++
                $gRequestTmp = New-Object -TypeName PSObject -Property @{
                    id     = $TeamsDevice.id + "-health"
                    method = "GET"
                    url    = "/teamwork/devices/" + $TeamsDevice.id + "/health"
                }
                [void]$graphRequests.Add($gRequestTmp)
            }
        }
        if ($graphRequests.Count -gt 0) {
            $graphResponseExtra = Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Update-UcTeamsDevices, getting device health info" -IncludeBody
        }

        #In case we detect more than 5 devices with updates pending we will request confirmation that we can continue.
        if (($devicesWithUpdatePending -gt 5) -and !$ReportOnly) {
            if ($ConfirmPreference) {
                $title = 'Confirm'
                $question = "There are " + $devicesWithUpdatePending + " Teams Devices pending update. Are you sure that you want to continue?"
                $choices = '&Yes', '&No'
                $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)
            }
            else {
                $decision = 0
            }
            if ($decision -ne 0) {
                return
            }
        }
        $graphRequests = [System.Collections.ArrayList]::new()
        foreach ($TeamsDevice in $TeamsDeviceList) {
            if ($TeamsDevice.healthStatus -in $StatusType) {
                $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body
                
                #Valid types are: adminAgent, operatingSystem, teamsClient, firmware, partnerAgent, companyPortal.
                #Currently we only consider Firmware and TeamsApp(teamsClient)
                
                #Firmware
                if ($TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType.Equals("All") -or $UpdateType.Equals("Firmware"))) {
                    if (!($ReportOnly)) {

                        $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                        $requestHeader.Add("Content-Type", "application/json")
                        $requestBody = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                        $requestBody.Add("softwareType", "firmware")
                        $requestBody.Add("softwareVersion", $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.availableVersion)

                        $gRequestTmp = New-Object -TypeName PSObject -Property @{
                            id      = $TeamsDevice.id + "-updateFirmware"
                            method  = "POST"
                            url     = "/teamwork/devices/" + $TeamsDevice.id + "/updateSoftware"
                            body    = $requestBody
                            headers = $requestHeader
                        }
                        [void]$graphRequests.Add($gRequestTmp)
                    }
                }
                
                #TeamsApp
                if ($TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType.Equals("All") -or $UpdateType.Equals("TeamsApp"))) {
                    if (!($ReportOnly)) {
                        $requestHeader = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                        $requestHeader.Add("Content-Type", "application/json")

                        $requestBody = New-Object 'System.Collections.Generic.Dictionary[string, string]'
                        $requestBody.Add("softwareType", "teamsClient")
                        $requestBody.Add("softwareVersion", $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.availableVersion)

                        $gRequestTmp = New-Object -TypeName PSObject -Property @{
                            id      = $TeamsDevice.id + "-updateTeamsClient"
                            method  = "POST"
                            url     = "/teamwork/devices/" + $TeamsDevice.id + "/updateSoftware"
                            body    = $requestBody
                            headers = $requestHeader
                        }
                        [void]$graphRequests.Add($gRequestTmp)
                    }
                }
            }
        }
        if ($graphRequests.Count -gt 0) {
            $updateGraphResponse = Invoke-UcMgGraphBatch -Requests $graphRequests -MgProfile beta -Activity "Update-UcTeamsDevices, sending update commands" -IncludeBody
        }

        foreach ($TeamsDevice in $TeamsDeviceList) {
            if ($TeamsDevice.healthStatus -in $StatusType) {
                $TeamsDeviceHealth = ($graphResponseExtra | Where-Object { $_.id -eq ($TeamsDevice.id + "-health") }).body

                if ($ReportOnly) {

                    if ($TeamsDevice.healthStatus -notin ("critical", "nonUrgent") ) {
                        $UpdateStatus = "Report Only: No firmware or Teams App updates pending."
                    }
                    else {
                        $UpdateStatus = "Report Only:"

                        if ($TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType.Equals("All") -or $UpdateType.Equals("Firmware"))) {
                            $UpdateStatus += " Firmware Update Pending;"
                        }

                        if ($TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.softwareFreshness.Equals("updateAvailable") -and ($UpdateType.Equals("All") -or $UpdateType.Equals("TeamsApp"))) {
                            $UpdateStatus += " Teams App Update Pending;"
                        }
                    }
                    
                }
                else {
                    $tmpUpdateStatus = ($updateGraphResponse | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateFirmware") })
                    if ($tmpUpdateStatus.status -eq 202) {
                        $UpdateStatus = "Firmware update Command was sent to the device"
                    }
                    elseif ($tmpUpdateStatus.status -eq 409) {
                        $UpdateStatus = "There is a firmware update pending, please check the update status."
                    }

                    $tmpUpdateStatus = ($updateGraphResponse | Where-Object { $_.id -eq ($TeamsDevice.id + "-updateTeamsClient") })
                    if ($tmpUpdateStatus.status -eq 202) {
                        $UpdateStatus = "Teams App update Command was sent to the device"
                    }
                    elseif ($tmpUpdateStatus.status -eq 409) {
                        $UpdateStatus = "There is a Teams App update pending, please check the update status."
                    }
                }

                $TDObj = New-Object -TypeName PSObject -Property @{
                    'Device Id'                     = $TeamsDevice.id
                    DeviceType                      = Convert-UcTeamsDeviceType $TeamsDevice.deviceType
                    Manufacturer                    = $TeamsDevice.hardwaredetail.manufacturer
                    Model                           = $TeamsDevice.hardwaredetail.model
                    HealthStatus                    = $TeamsDevice.healthStatus
                    TeamsAdminAgentCurrentVersion   = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.currentVersion
                    TeamsAdminAgentAvailableVersion = $TeamsDeviceHealth.softwareUpdateHealth.adminAgentSoftwareUpdateStatus.availableVersion
                    FirmwareCurrentVersion          = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.currentVersion
                    FirmwareAvailableVersion        = $TeamsDeviceHealth.softwareUpdateHealth.firmwareSoftwareUpdateStatus.availableVersion
                    CompanyPortalCurrentVersion     = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.currentVersion
                    CompanyPortalAvailableVersion   = $TeamsDeviceHealth.softwareUpdateHealth.companyPortalSoftwareUpdateStatus.availableVersion
                    OEMAgentAppCurrentVersion       = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.currentVersion
                    OEMAgentAppAvailableVersion     = $TeamsDeviceHealth.softwareUpdateHealth.partnerAgentSoftwareUpdateStatus.availableVersion
                    TeamsAppCurrentVersion          = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.currentVersion
                    TeamsAppAvailableVersion        = $TeamsDeviceHealth.softwareUpdateHealth.teamsClientSoftwareUpdateStatus.availableVersion
                    UpdateStatus                    = $UpdateStatus
                }
                [void]$outTeamsDevices.Add($TDObj)
            }
        }

        if ( $outTeamsDevices.Count -gt 0) {
            $outTeamsDevices | Sort-Object DeviceType, Manufacturer, Model | Select-Object 'Device Id', DeviceType, Manufacturer, Model, HealthStatus, TeamsAdminAgentCurrentVersion, TeamsAdminAgentAvailableVersion, FirmwareCurrentVersion, FirmwareAvailableVersion, CompanyPortalCurrentVersion, CompanyPortalAvailableVersion, OEMAgentAppCurrentVersion, OEMAgentAppAvailableVersion, TeamsAppCurrentVersion, TeamsAppAvailableVersion, UpdateStatus  | Export-Csv -path $OutputFullPath -NoTypeInformation
            Write-Host ("Results available in: " + $OutputFullPath) -ForegroundColor Cyan
        }
        else {
            Write-Host ("No Teams Device(s) found that have pending update.") -ForegroundColor Cyan
        }
    }
}