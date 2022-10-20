param(
    [version]$Version = "0.2.0"

)
#Requires -Module ModuleBuilder

$params = @{
    SourcePath = "$PSScriptRoot\Source\UCLobbyTeams.psd1"
    CopyPaths = "$PSScriptRoot\README.md"
    Version = $Version
    UnversionedOutputDirectory = $true
}

Build-Module @params -Verbose
Import-Module -Name ($PSScriptRoot+"\Output\Uclobbyteams") -Verbose -Force