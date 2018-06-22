#
# Module manifest for module 'Helper'
#
# Generated by: Jack Bennett
#
# Generated on: 21/09/2017
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'Helper.psm1'

# Version number of this module.
ModuleVersion = '1.3.1'

# Supported PSEditions
# CompatiblePSEditions = @()

# ID used to uniquely identify this module
GUID = 'ddb5316f-0dda-4927-af11-aa64e37e944c'

# Author of this module
Author = 'Jack Bennett <github@jackben.net>'

# Company or vendor of this module
CompanyName = 'Birkdale High'

# Copyright statement for this module
Copyright = '(c) 2017 Jack Bennett. All rights reserved.'

# Description of the functionality provided by this module
Description = 'Tasks done often or rarely can be useful to wrap up quirky or verbose commands into short custom functions that are easy to remember.'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '3.0'

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
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
# Get new list with; Import-module .\helper.psm1 -force; Get-Command -module helper | Select name | convertto-csv -NoTypeInformation | select -Skip 1 | Set-clipboard
# This bypasses the exported command listsed here in .psd1
FunctionsToExport = @(
    "Add-Access"
    "Disable-Access"
    "Enable-Access"
    "Get-BypassedSender"
    "Get-MailServer"
    "Get-MonitorInformation"
    "Get-RecentFailedMessage"
    "Import-MailServer"
    "Move-Work"
    "New-SharedMailbox"
    "Remove-Access"
    "Remove-BypassedSender"
    "Remove-MailServer"
    "Search-MailDate"
    "Search-MailFrom"
    "Search-MailSubject"
    "Show-Access"
)

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

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
        Tags = @('Exchange', 'NTFS')

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/BirkdaleHigh/ps-helper/blob/master/LICENSE'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/BirkdaleHigh/ps-helper'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

