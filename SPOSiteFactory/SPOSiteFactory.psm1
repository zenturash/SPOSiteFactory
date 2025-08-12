#Requires -Version 5.1
#Requires -Modules PnP.PowerShell, PSFramework, Microsoft.PowerShell.SecretManagement

<#
.SYNOPSIS
    SPOSiteFactory PowerShell Module - MSP Edition

.DESCRIPTION
    SharePoint Online Site Factory module designed for Managed Service Providers (MSPs) 
    to manage multiple tenants with security auditing, provisioning automation, and 
    compliance reporting capabilities.

.NOTES
    Author: MSP PowerShell Team
    Version: 0.1.0
    Requires: PowerShell 5.1+, PnP.PowerShell 2.0+, PSFramework 1.7+
#>

#region Module Variables and Configuration

# Module scope variables for MSP operations
$script:SPOFactoryConnections = @{}
$script:SPOFactoryConfig = @{
    MSPTenantId = $null
    ClientTenants = @{}
    DefaultBaseline = 'MSPStandard'
    LogPath = "$env:ProgramData\SPOSiteFactory\Logs"
    CredentialVault = 'SPOFactory'
    ConfigPath = "$env:ProgramData\SPOSiteFactory\Config"
    EnableAuditLog = $true
    EnableDebugLogging = $false
    AlertEmail = $null
    MaxConcurrentConnections = 50
    ConnectionTimeout = 300
    RetryAttempts = 3
    BatchSize = 100
}

# MSP-specific constants
$script:SPOFactoryConstants = @{
    SupportedRegions = @('Global', 'GCC', 'GCCH', 'DoD')
    MaxTenantCount = 1000
    LogRetentionDays = 90
    ConfigVersion = '1.0'
    ModuleVersion = '0.1.0'
}

# Connection pool for MSP multi-tenant scenarios
$script:SPOFactoryConnectionPool = @{}
$script:SPOFactoryBaselines = @{}
$script:SPOFactoryTemplates = @{}

#endregion

#region Module Initialization

Write-Host "Initializing SPOSiteFactory MSP Module v$($script:SPOFactoryConstants.ModuleVersion)..." -ForegroundColor Green

# Ensure required directories exist
$requiredDirs = @(
    $script:SPOFactoryConfig.LogPath,
    $script:SPOFactoryConfig.ConfigPath,
    "$($script:SPOFactoryConfig.ConfigPath)\Tenants",
    "$($script:SPOFactoryConfig.ConfigPath)\Baselines",
    "$($script:SPOFactoryConfig.ConfigPath)\Templates"
)

foreach ($dir in $requiredDirs) {
    if (-not (Test-Path $dir)) {
        try {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
            Write-Verbose "Created directory: $dir"
        }
        catch {
            Write-Warning "Failed to create directory $dir`: $_"
        }
    }
}

#endregion

#region Function Loading

# Get public and private function definition files
$PublicFunctions = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue)
$PrivateFunctions = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

# Dot source the functions
foreach ($import in @($PublicFunctions + $PrivateFunctions)) {
    try {
        Write-Verbose "Importing function: $($import.Name)"
        . $import.FullName
    }
    catch {
        Write-Error "Failed to import function $($import.FullName): $_"
    }
}

# Export public functions
if ($PublicFunctions) {
    Export-ModuleMember -Function $PublicFunctions.BaseName
    Write-Verbose "Exported $($PublicFunctions.Count) public functions"
}

#endregion

#region PSFramework Logging Configuration

# Configure PSFramework logging for MSP environments
try {
    # Set up file logging provider
    $logFilePath = Join-Path $script:SPOFactoryConfig.LogPath "SPOFactory-$(Get-Date -Format 'yyyy-MM-dd').log"
    
    Set-PSFLoggingProvider -Name logfile -FilePath $logFilePath -Enabled $true -LogRotatePath $script:SPOFactoryConfig.LogPath -LogRetentionTime (New-TimeSpan -Days $script:SPOFactoryConstants.LogRetentionDays)
    
    # Configure log levels
    if ($script:SPOFactoryConfig.EnableDebugLogging) {
        Set-PSFConfig -FullName 'psframework.logging.maximummessagelevel' -Value 'Debug'
    }
    
    Write-PSFMessage -Level Host -Message "SPOSiteFactory logging initialized. Log path: $logFilePath"
}
catch {
    Write-Warning "Failed to initialize PSFramework logging: $_"
}

#endregion

#region Module Cleanup

# Register module removal event
$ExecutionContext.SessionState.Module.OnRemove = {
    Write-PSFMessage -Level Host -Message "Cleaning up SPOSiteFactory module..."
    
    # Disconnect all active connections
    if ($script:SPOFactoryConnections.Count -gt 0) {
        foreach ($connection in $script:SPOFactoryConnections.Keys) {
            try {
                if (Get-PnPConnection -ErrorAction SilentlyContinue) {
                    Disconnect-PnPOnline
                }
                Write-PSFMessage -Level Verbose -Message "Disconnected from tenant: $connection"
            }
            catch {
                Write-PSFMessage -Level Warning -Message "Failed to disconnect from $connection`: $_"
            }
        }
    }
    
    # Clear connection pool
    $script:SPOFactoryConnectionPool.Clear()
    $script:SPOFactoryConnections.Clear()
    
    Write-PSFMessage -Level Host -Message "SPOSiteFactory module cleanup completed"
}

#endregion

#region Module Validation

# Validate required modules are available
$requiredModules = @('PnP.PowerShell', 'PSFramework', 'Microsoft.PowerShell.SecretManagement')
$missingModules = @()

foreach ($module in $requiredModules) {
    if (-not (Get-Module -Name $module -ListAvailable)) {
        $missingModules += $module
    }
}

if ($missingModules.Count -gt 0) {
    Write-Warning "Missing required modules: $($missingModules -join ', ')"
    Write-Warning "Please install missing modules using: Install-Module -Name $($missingModules -join ', ')"
}

# Validate secret vault exists
try {
    if (-not (Get-SecretVault -Name $script:SPOFactoryConfig.CredentialVault -ErrorAction SilentlyContinue)) {
        Write-PSFMessage -Level Warning -Message "Secret vault '$($script:SPOFactoryConfig.CredentialVault)' not found. MSP credential management will be limited."
        Write-PSFMessage -Level Host -Message "To enable full MSP functionality, create a secret vault using: Register-SecretVault"
    }
}
catch {
    Write-PSFMessage -Level Warning -Message "Could not validate secret vault: $_"
}

#endregion

Write-Host "SPOSiteFactory MSP Module loaded successfully!" -ForegroundColor Green
Write-Host "Ready to manage $($script:SPOFactoryConstants.MaxTenantCount)+ SharePoint Online tenants" -ForegroundColor Cyan

# Display initialization summary
Write-PSFMessage -Level Host -Message @"
SPOSiteFactory Initialization Summary:
- Configuration Path: $($script:SPOFactoryConfig.ConfigPath)
- Log Path: $($script:SPOFactoryConfig.LogPath)
- Credential Vault: $($script:SPOFactoryConfig.CredentialVault)
- Max Concurrent Connections: $($script:SPOFactoryConfig.MaxConcurrentConnections)
- Batch Size: $($script:SPOFactoryConfig.BatchSize)
- Debug Logging: $($script:SPOFactoryConfig.EnableDebugLogging)
"@