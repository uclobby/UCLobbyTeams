param(
    [version]$Version = "1.0.0"

)
#Requires -Module ModuleBuilder

$params = @{
    SourcePath = "$PSScriptRoot\Source\UCLobbyTeams.psd1"
    CopyPaths = "$PSScriptRoot\README.md"
    Version = $Version
    UnversionedOutputDirectory = $true
}

Build-Module @params
$PSScriptRoot+"\Output\Uclobbyteams"
Import-Module -Name ($PSScriptRoot+"\Output\Uclobbyteams") -Force
