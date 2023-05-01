param(
    [version]$Version = "0.3.2"

)
#Requires -Module ModuleBuilder

$params = @{
    SourcePath = "$PSScriptRoot\Source\UCLobbyTeams.psd1"
    CopyPaths = "$PSScriptRoot\README.md"
    Version = $Version
    UnversionedOutputDirectory = $true
}

Build-Module @params
Import-Module -Name ($PSScriptRoot+"\Output\Uclobbyteams") -Force


