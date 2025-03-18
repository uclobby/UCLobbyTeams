function Get-UcTeamsDeviceConfigurationProfile {
    <#
        .SYNOPSIS
        Returns all Teams Device Configuration Profiles

        .DESCRIPTION
        This function fetch Teams Devices Configuration Profiles using the TAC API.

        Requirements:   EntraAuth PowerShell Module (Install-Module EntraAuth)

        .PARAMETER Filter
        Specifies a filter, valid options:
            Phone - Teams Native Phones
            MTR - Microsoft Teams Rooms running Windows or Android
            Display - Microsoft Teams Displays 
            Panel - Microsoft Teams Panels

        .EXAMPLE 
        PS> Get-UcTeamsDeviceConfigurationProfile

        .EXAMPLE 
        PS> Get-UcTeamsDeviceConfigurationProfile -Filter MTR
    #>
    param (
        [String]$Identity,
        [ValidateSet("Phone", "MTR", "Display", "Panel")]
        [string]$Filter
    )
    $outTeamsDeviceConfiguration = [System.Collections.ArrayList]::new()
    $BaseAPIPath = "api/v2/configProfiles"

    if (!($global:UCLobbyTeamsModuleCheck)) {
        Test-UcPowerShellModule -ModuleName UcLobbyTeams | Out-Null
        $global:UCLobbyTeamsModuleCheck = $true
    }

    if (Test-UcAPIConnection -Type TeamsDeviceTAC) {
        if ($Identity) {
            $RequestPath = $BaseAPIPath + "/" + $Identity
            try{
                $ConfigurationProfiles = Invoke-EntraRequest -Path $RequestPath -Service TeamsDeviceTAC
            } catch{ }
        }
        else {
            switch ($Filter) {
                "Phone" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"IpPhone`"]}}" }
                "MTR" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsRoom`",`"collaborationBar`"]}}" }
                "Display" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsDisplay`"]}}" }
                "Panel" { $DeviceFilter = "{`"deviceTypes`":{`"eq`":[`"teamsPanel`"]}}" }
                default {
                    $DeviceFilter = $null
                }
            }
            if ($DeviceFilter) {
                $RequestPath = $BaseAPIPath + "?filterJson=" + [System.Web.HttpUtility]::UrlEncode($DeviceFilter)
            }
            else {
                $RequestPath = $BaseAPIPath
            }
            #TODO: Adding support for Pages, currently we are not expecting more than 50 Configuration Profiles.
            try{
            $ConfigurationProfiles = (Invoke-EntraRequest $RequestPath -Service TeamsDeviceTAC).paginatedConfigProfiles
            } catch {}
        }
        foreach ($ConfigurationProfile in $ConfigurationProfiles) {
            #Maintenance Window
            $tmpMaintenanceWindowDaysOfWeek = ""
            $tmpMaintenanceWindowStartTime = ""
            $tmpMaintenanceWindowDuration = ""
            foreach ($maintenanceSchedule in $ConfigurationProfile.maintenanceConfig.weeklyMaintenanceWindows) {
                $tmpMaintenanceWindowDaysOfWeek += $maintenanceSchedule.dayOfWeek + ","
                $tmpMaintenanceWindowStartTime = $maintenanceSchedule.startTime
                $tmpMaintenanceWindowDuration = $maintenanceSchedule.duration
            }

            #Restart Teams App has 3 Options: Never, When Needed, Daily - HH:MM
            $tmpRestartTeamsApp = ""
            if (![string]::IsNullOrEmpty($ConfigurationProfile.deviceAppRestartConfig.isDeviceAppRestartEnabled)) {
                $tmpRestartTeamsApp = "Never"
                if ($ConfigurationProfile.deviceAppRestartConfig.isDeviceAppRestartEnabled) {
                    $tmpRestartTeamsApp = "When needed"
                    if (!$ConfigurationProfile.deviceAppRestartConfig.isDeviceAppAutoRestartEnabled) {
                        $tmpRestartTimeMinutes = $ConfigurationProfile.deviceAppRestartConfig.scheduledDeviceAppRestartTime
                        if ([string]::IsNullOrEmpty($tmpRestartTimeMinutes)) {
                            $tmpRestartTimeMinutes = 1080
                        }
                        $tmpRestartTeamsApp = "Daily - " + (New-TimeSpan -Minutes $tmpRestartTimeMinutes).ToString("hh\:mm")
                    }
                }
            }
        
            $TDObj = [PSCustomObject][Ordered]@{
                Identity                               = $ConfigurationProfile.baseInfo.id
                DisplayName                            = $ConfigurationProfile.identity
                Description                            = $ConfigurationProfile.description
                DeviceType                             = Convert-UcTeamsDeviceType $ConfigurationProfile.deviceType
            
                #region General Settings
                #All Devices
                DeviceLock                             = $ConfigurationProfile.deviceLock
                DeviceLockTimeoutinSeconds             = $ConfigurationProfile.deviceLockTimeout
                DeviceLockPin                          = $ConfigurationProfile.deviceLockPin

                #MTRoA Only
                HDMIContentSharing                     = $ConfigurationProfile.hdmiIngestConfig.isContentSharingEnabled
                HDMIContentSharingIncludeAudio         = $ConfigurationProfile.hdmiIngestConfig.isAudioSharingEnabled
                HDMIContentSharingAutoSharing          = $ConfigurationProfile.hdmiIngestConfig.isAutoSharingEnabled

                #Phone Only
                EnforceDeviceLock                      = $ConfigurationProfile.isForcePinChangeEnabled
            
                #All Devices
                Language                               = $ConfigurationProfile.language
                Timezone                               = $ConfigurationProfile.timezone
                DateFormat                             = $ConfigurationProfile.dateformat
                TimeFormat                             = $ConfigurationProfile.timeFormat
                MaintenanceWindowDaysOfWeek            = $tmpMaintenanceWindowDaysOfWeek -replace ".$"
                MaintenanceWindowStartTime             = $tmpMaintenanceWindowStartTime -replace ".{3}$"
                MaintenanceWindowDurationInMinutes     = $tmpMaintenanceWindowDuration
                DailyDeviceRestart                     = $ConfigurationProfile.maintenanceConfig.isAutoDeviceRestartEnabled
                DailyDeviceRestartStartTime            = $ConfigurationProfile.maintenanceConfig.dailyMaintenanceWindow.startTime 
                DailyDeviceRestartDurationInMinutes    = $ConfigurationProfile.maintenanceConfig.dailyMaintenanceWindow.duration
            
                #MTRoA Only
                EnableTouchScreenControls              = $ConfigurationProfile.isTouchscreenControlsEnabled
                DualDisplayMode                        = $ConfigurationProfile.dualDisplayModeConfig.isDualDisplayModeEnabled
                DualDisplaySwapScreens                 = $ConfigurationProfile.dualDisplayModeConfig.isSwapScreensEnabled
                LogsAndFeedbackEmail                   = $ConfigurationProfile.logsAndFeedbackConfig.emailAddress
                LogsIncludeFeedback                    = $ConfigurationProfile.logsAndFeedbackConfig.isSendLogsAndFeedbackEnabled

                #Panel Only
                ShowRoomEquipment                      = $ConfigurationProfile.isRoomEquipmentEnabled

                #Phone and Display
                RestartTeamsApp                        = $tmpRestartTeamsApp
                #endregion

                #region Meeting Settings
                #Panel Only
                AllowReservationFromPanel              = $ConfigurationProfile.areRoomReservationsEnabled
                ShowMeetingNames                       = !$ConfigurationProfile.hideMeetingNames

                #Panel and Display (for Display it's under General Settings)
                ShowQRCodeForReservations              = $ConfigurationProfile.isQRCodeForSignInOrReservationEnabled

                #MTRoA and Panel
                RoomCapacityNotification               = $ConfigurationProfile.roomCapacityNotificationEnabled
                AllowReservationExtension              = $ConfigurationProfile.extendRoomReservationEnabled
            
                #MTRoA Only
                AllowWhiteBoard                        = $ConfigurationProfile.isWhiteboardEnabled
                RequirePasscodeToJoin                  = $ConfigurationProfile.isRequirePasscodeForMeetingsEnabled

                #Panel Only
                SendCheckinNotification                = $ConfigurationProfile.checkinNotification
                ReleaseRoomIfNoCheckin                 = $ConfigurationProfile.checkinRoomReleaseConfig.enabled
                ReleaseRoomIfNoCheckinTimeoutInSeconds = $ConfigurationProfile.checkinRoomReleaseConfig.timeout
                EarlyCheckout                          = $ConfigurationProfile.checkoutRoomEnabled
                #endregion

                #region Calling Settings
                #Phone Only
                AdvancedCalling                        = $ConfigurationProfile.isIPPhonePremiumCapSkuEnabled
                CallQualitySurvey                      = $ConfigurationProfile.isCallQualityFeedbackSurveyEnabled
                DisplayCallForwardingOnHomeScreen      = $ConfigurationProfile.isIPPhoneCallForwardingOnTheMainScreenEnabled 
            
                #Phone and Display, on display it's the Virtual FrontDesk
                Hotline                                = $ConfigurationProfile.hotlineConfig.hotlineIsEnabled
                HotlineDisplayName                     = $ConfigurationProfile.hotlineConfig.hotlineDisplayName

                #Display Only
                VirtualFrontDeskVideoEnabled           = $ConfigurationProfile.hotlineConfig.isVideoDefaultOn
                VirtualFrontDeskWelcomeMessage         = $ConfigurationProfile.hotlineConfig.welcomeMessage
                #endregion

                #region Device Settings
                #Panel Only
                BusyLightColor                         = $ConfigurationProfile.usageStateIndicatorBusyStateColor.name
                #MTRoA and Panel
                Background                             = $ConfigurationProfile.theme.name

                #All Devices
                DisplayScreenSaver                     = $ConfigurationProfile.displayScreenSaver
                DisplayScreenSaverTimeoutInSeconds     = $ConfigurationProfile.screenSaverTimeout

                #All Devices
                DisplayBacklightBrightness             = $ConfigurationProfile.displayBacklitBrightness
                DisplayBacklightTimeoutInSeconds       = $ConfigurationProfile.displayBacklitTimeout
                DisplayHighContrast                    = $ConfigurationProfile.displayHighContrast
                SilentMode                             = $ConfigurationProfile.silentMode
                OfficeStartHours                       = $ConfigurationProfile.officeStartHours
                OfficeEndHours                         = $ConfigurationProfile.officeEndHours
                PowerSaving                            = $ConfigurationProfile.powerSaving
                ScreenCapture                          = $ConfigurationProfile.screenCapture
            
                #MTRoA Only
                BluetoothBeaconing                     = $ConfigurationProfile.bluetoothBeconingEnabled
                BtBeaconAcceptProximityInvitations     = $ConfigurationProfile.roomQrCodeConfig.isAcceptProximityInvitationEnabled
                RemoteControlFromPersonalDevices       = $ConfigurationProfile.allowRoomRemoteEnabled
                ShowRoomQRcode                         = $ConfigurationProfile.roomQrCodeConfig.isRoomQrCodeEnabled
                QRCodeAcceptProximityInvitations       = $ConfigurationProfile.roomQrCodeConfig.isAcceptProximityInvitationEnabled
                #endregion

                #region Network Settings
                #All Devices
                DHCPEnabled                            = $ConfigurationProfile.dhcpEnabled
                LoggingEnabled                         = $ConfigurationProfile.loggingEnabled
                Hostname                               = $ConfigurationProfile.hostName
                DomainName                             = $ConfigurationProfile.domainName
                IPAddress                              = $ConfigurationProfile.ipAddress
                SubnetMask                             = $ConfigurationProfile.subnetMask
                DefaultGateway                         = $ConfigurationProfile.defaultGateway
                PrimaryDNS                             = $ConfigurationProfile.primaryDNS
                SecondaryDNS                           = $ConfigurationProfile.secondaryDNS
                DefaultAdminPassword                   = $ConfigurationProfile.deviceDefaultAdminPassword
                NetworkPCport                          = $ConfigurationProfile.networkPcPort
                #endregion

                #All Devices
                WhenCreated                            = (Get-Date "01/01/1970").AddSeconds($ConfigurationProfile.baseInfo.createdAt)
                CreatedByUser                          = $ConfigurationProfile.baseInfo.createdByUserName
                WhenChanged                            = (Get-Date "01/01/1970").AddSeconds($ConfigurationProfile.baseInfo.modifiedAt)
                ChangedByUser                          = $ConfigurationProfile.baseInfo.modifiedByUserName
            }
            $TDObj.PSObject.TypeNames.Insert(0, 'TeamsDeviceConfigurationProfile')
            [void]$outTeamsDeviceConfiguration.Add($TDObj)
        }
        return $outTeamsDeviceConfiguration | Sort-Object DeviceType
    }
}