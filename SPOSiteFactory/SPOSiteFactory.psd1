@{
    # Script module or binary module file associated with this manifest.
    RootModule = 'SPOSiteFactory.psm1'

    # Version number of this module.
    ModuleVersion = '0.1.0'

    # Supported PSEditions
    CompatiblePSEditions = @('Desktop', 'Core')

    # ID used to uniquely identify this module
    GUID = 'a1b2c3d4-e5f6-7890-1234-567890123456'

    # Author of this module
    Author = 'MSP PowerShell Team'

    # Company or vendor of this module
    CompanyName = 'Managed Service Provider'

    # Copyright statement for this module
    Copyright = '(c) 2025 MSP PowerShell Team. All rights reserved.'

    # Description of the functionality provided by this module
    Description = 'SharePoint Online Site Factory module designed for Managed Service Providers (MSPs) to manage multiple tenants with security auditing, provisioning automation, and compliance reporting capabilities.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion = '5.1'

    # Name of the PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    DotNetFrameworkVersion = '4.7.2'

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # ClrVersion = ''

    # Processor architecture (None, X86, Amd64) required by this module
    # ProcessorArchitecture = ''

    # Modules that must be imported into the global environment prior to importing this module
    RequiredModules = @(
        @{
            ModuleName = 'PnP.PowerShell'
            ModuleVersion = '2.0.0'
        },
        @{
            ModuleName = 'PSFramework'
            ModuleVersion = '1.7.0'
        },
        @{
            ModuleName = 'Microsoft.PowerShell.SecretManagement'
            RequiredVersion = '1.1.2'
        }
    )

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
    FunctionsToExport = @()

    # Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
    CmdletsToExport = @()

    # Variables to export from this module
    VariablesToExport = @()

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
            # MSP-specific configuration
            MSPConfiguration = @{
                SupportedTenantCount = 1000
                MaxConcurrentConnections = 50
                DefaultCredentialVault = 'SPOFactory'
                AuditLogRetentionDays = 90
                EnableComplianceReporting = $true
                SupportedRegions = @('Global', 'GCC', 'GCCH', 'DoD')
            }

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags = @('SharePoint', 'SPO', 'MSP', 'Automation', 'Security', 'Auditing', 'Provisioning', 'MultiTenant')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/MSPPowerShell/SPOSiteFactory/blob/main/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/MSPPowerShell/SPOSiteFactory'

            # A URL to an icon representing this module.
            IconUri = 'https://github.com/MSPPowerShell/SPOSiteFactory/blob/main/icon.png'

            # ReleaseNotes of this module
            ReleaseNotes = @'
Version 0.1.0 - Initial Release
- Module foundation and structure
- Multi-tenant connection management
- MSP-focused logging framework
- Error handling with tenant isolation
- Configuration management for multiple clients
- Basic security auditing capabilities
- Support for 1000+ tenant environments
- Integration with PSFramework and SecretManagement
'@

            # Prerelease string of this module
            # Prerelease = ''

            # Flag to indicate whether the module requires explicit user acceptance for install/update/save
            # RequireLicenseAcceptance = $false

            # External dependent modules of this module
            ExternalModuleDependencies = @('PnP.PowerShell', 'PSFramework', 'Microsoft.PowerShell.SecretManagement')
        }
    }

    # HelpInfo URI of this module
    HelpInfoURI = 'https://github.com/MSPPowerShell/SPOSiteFactory/blob/main/docs/help'

    # Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
    DefaultCommandPrefix = 'SPOFactory'
}