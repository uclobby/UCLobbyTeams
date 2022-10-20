function Convert-UcTeamsDeviceType {
    param (
        [string]$DeviceType
    )
    switch ($DeviceType) {
        "ipPhone" { return "Phone" }
        "lowCostPhone" { return "Phone" }
        "teamsRoom" { return "MTR Windows" }
        "collaborationBar" { return "MTR Android" }
        "surfaceHub" { return "Surface Hub" }
        "teamsDisplay" { return "Display" }
        "touchConsole" { return "Touch Console (MTRA)" }
        "teamsPanel" { return "Panel" }
        "sip" { return "SIP Phone" }
        Default { return $DeviceType}
    }
}