function Convert-UcTeamsDeviceAutoUpdateDays {
    param (
        [int]$Days
    )
    switch ($Days) {
        0 { return "Validation" }
        30 { return "General" }
        90 { return "Final" }
        default { return $Days }
    }
}