param(
    [version]$Version = "0.4.3"

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


