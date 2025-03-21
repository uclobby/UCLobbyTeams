#
# Module manifest for module 'UcLobbyTeams'
#
# Generated by: David Paulino
#
# Generated on: 10/6/2022
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'UcLobbyTeams.psm1'

# Version number of this module.
ModuleVersion = '0.1.5'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = '3984bf6e-8fcd-458c-b9fa-a53326b95798'

# Author of this module
Author = 'David Paulino'

# Company or vendor of this module
CompanyName = 'UC Lobby'

# Copyright statement for this module
Copyright = '(c) 2023 David Paulino. All rights reserved.'

# Description of the functionality provided by this module
Description = 'UC Lobby Teams PowerShell Module'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.1'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
#RequiredModules = @("Microsoft.Graph.Authentication")

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
FormatsToProcess = 'UcLobbyTeams.format.ps1xml'

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = 'Get-UcTeamsVersion', 'Get-UcM365TenantId', 'Get-UcTeamsForest', 
               'Get-UcM365Domains', 'Test-UcTeamsOnlyModeReadiness', 
               'Get-UcTeamUsersEmail', 'Get-UcTeamsWithSingleOwner', 'Get-UcArch', 
               'Get-UcTeamsDevice', 'Test-UcTeamsDevicesConditionalAccessPolicy', 
               'Test-UcTeamsDevicesCompliancePolicy'

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = '*'

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        Tags = 'MicrosoftTeams', 'Teams', 'TeamsDevices', 'Microsoft365', 'Office365', 'MicrosoftGraph', 'GraphAPI','Graph'

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        ProjectUri = 'https://uclobby.com/uclobby-teams-powershell-module/'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        ReleaseNotes = '
1.0.0 - 2025/03/17
    Get-UcTeamsDeviceConfigurationProfile
        New cmdlet: Returns all Teams Device Configuration Profiles.
    Set-UcTeamsDeviceConfigurationProfile
        New cmdlet: Allows to set a configuration profile to Teams Device (Phone, MTRoA and Panels).
    Connect-UcTeamsDeviceTAC
        New cmdlet: Connects to TAC Api using the EntraAuth PowerShell module to manage Authentication tokens/request.
    Get-UcTeamsDevice
        Change: UseTAC switch requires EntraAuth PowerShell module instead of MSAL.PS.
        Feature: Added TACDeviceID parameter in case we want to get the details for a single Teams Device, works with Graph API and UseTAC.
        Fix: WhenCreated/WhenChanged was showing the wrong Date.
    Export-UcOneDriveWithMultiplePermissions
        Fix: Issue with access denied while writing csv file.0.7.1 - 2025/02/03
    Get-UcTeamsDevice
        Fix: Issue checking the MS Graph Permissions.
        Feature: Added switch UseTAC that allows to use Teams Admin Center API to get the Teams Devices information.
0.7.0 - 2024/11/01
    Get-UcObjectsOwnedByUser
        New cmdlet: Returns all Entra objects associated with a user.
    Get-UcTeamsVersion
        Fix: Teams Classic was include in the output if settings file was present after Teams Classic uninstallation.
        Fix: Running this in Windows 10 with PowerShell 7 an exception could be raised while importing the Appx PowerShell module. Thank you Steve Chupack for reporting this issue.
0.6.3 - 2024/10/25
    Get-UcTeamsVersion
        Fix: No output generated for New Teams if the tma_settings.json file was missing.
0.6.2 - 2024/10/23
    Export-UcM365LicenseAssignment
        Change: The SKU parameter can be use to search, if we use "copilot" then all licenses with copilot will be included in the export file.
        Fix: In some scenarios the license exists in the tenant but no information available in "Products names and Services Identifiers" file, for these cases the output will be the SKU Part Number.
    Update-UcTeamsDevice
        Change: For ReportOnly we can use TeamworkDevice.Read.All, since we dont require write permissions (TeamworkDevice.ReadWrite.All).
        Change: Now the user/application running the cmdlet can have User.ReadBasic.All or User.Read.All as graph permission.
        Fix: The current user display name was empty in the output file.'
        
    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
HelpInfoURI = 'https://uclobby.com/uclobby-teams-powershell-module/'

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

