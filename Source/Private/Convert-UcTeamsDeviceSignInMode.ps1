function Convert-UcTeamsDeviceSignInMode {
    param (
        [string]$SignInMode
    )
    switch ($SignInMode) {
        "commonArea" { return "Common Area" }
        "personal" { return "User" }
        "conference" { return "Conference" }
        default { return $SignInMode}
    }
}