function Test-UcElevatedPrivileges {
    if (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
        return $false
    }
    return $true
}