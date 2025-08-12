#Requires -Version 5.1
# Temporarily commented for testing - uncomment in production
# #Requires -Modules PnP.PowerShell, PSFramework, Microsoft.PowerShell.SecretManagement

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

# Import Private Functions
Get-ChildItem -Path "$PSScriptRoot\Private" -Filter "*.ps1" | 
    ForEach-Object { . $_.FullName }

# Import Public Functions from each subfolder
# Connection functions
Get-ChildItem -Path "$PSScriptRoot\Public\Connection" -Filter "*.ps1" -ErrorAction SilentlyContinue | 
    ForEach-Object { . $_.FullName }

# Configuration functions  
Get-ChildItem -Path "$PSScriptRoot\Public\Configuration" -Filter "*.ps1" -ErrorAction SilentlyContinue | 
    ForEach-Object { . $_.FullName }

# Hub functions
Get-ChildItem -Path "$PSScriptRoot\Public\Hub" -Filter "*.ps1" -ErrorAction SilentlyContinue | 
    ForEach-Object { . $_.FullName }

# Provisioning functions
Get-ChildItem -Path "$PSScriptRoot\Public\Provisioning" -Filter "*.ps1" -ErrorAction SilentlyContinue | 
    ForEach-Object { . $_.FullName }

# Security functions (if any)
Get-ChildItem -Path "$PSScriptRoot\Public\Security" -Filter "*.ps1" -ErrorAction SilentlyContinue | 
    ForEach-Object { . $_.FullName }

#endregion

#region PSFramework Logging Configuration

# Configure PSFramework logging for MSP environments
try {
    # Set up file logging provider
    $logFilePath = Join-Path $script:SPOFactoryConfig.LogPath "SPOFactory-$(Get-Date -Format 'yyyy-MM-dd').log"
    
    if (Get-Command Set-PSFLoggingProvider -ErrorAction SilentlyContinue) {
        Set-PSFLoggingProvider -Name logfile -FilePath $logFilePath -Enabled $true -LogRotatePath $script:SPOFactoryConfig.LogPath -LogRetentionTime (New-TimeSpan -Days $script:SPOFactoryConstants.LogRetentionDays)
        
        # Configure log levels
        if ($script:SPOFactoryConfig.EnableDebugLogging) {
            Set-PSFConfig -FullName 'psframework.logging.maximummessagelevel' -Value 'Debug'
        }
        
        Write-PSFMessage -Level Host -Message "SPOSiteFactory logging initialized. Log path: $logFilePath"
    } else {
        Write-Host "SPOSiteFactory logging initialized. Log path: $logFilePath" -ForegroundColor Cyan
    }
}
catch {
    Write-Warning "Failed to initialize PSFramework logging: $_"
}

#endregion

#region Module Cleanup

# Register module removal event
$ExecutionContext.SessionState.Module.OnRemove = {
    if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
        Write-PSFMessage -Level Host -Message "Cleaning up SPOSiteFactory module..."
    } else {
        Write-Host "Cleaning up SPOSiteFactory module..." -ForegroundColor Yellow
    }
    
    # Disconnect all active connections
    if ($script:SPOFactoryConnections.Count -gt 0) {
        foreach ($connection in $script:SPOFactoryConnections.Keys) {
            try {
                if (Get-PnPConnection -ErrorAction SilentlyContinue) {
                    Disconnect-PnPOnline
                }
                if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
                    Write-PSFMessage -Level Verbose -Message "Disconnected from tenant: $connection"
                } else {
                    Write-Verbose "Disconnected from tenant: $connection"
                }
            }
            catch {
                if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
                    Write-PSFMessage -Level Warning -Message "Failed to disconnect from $connection`: $_"
                } else {
                    Write-Warning "Failed to disconnect from $connection`: $_"
                }
            }
        }
    }
    
    # Clear connection pool
    $script:SPOFactoryConnectionPool.Clear()
    $script:SPOFactoryConnections.Clear()
    
    if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
        Write-PSFMessage -Level Host -Message "SPOSiteFactory module cleanup completed"
    } else {
        Write-Host "SPOSiteFactory module cleanup completed" -ForegroundColor Green
    }
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
    if (Get-Command Get-SecretVault -ErrorAction SilentlyContinue) {
        if (-not (Get-SecretVault -Name $script:SPOFactoryConfig.CredentialVault -ErrorAction SilentlyContinue)) {
            if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
                Write-PSFMessage -Level Warning -Message "Secret vault '$($script:SPOFactoryConfig.CredentialVault)' not found. MSP credential management will be limited."
                Write-PSFMessage -Level Host -Message "To enable full MSP functionality, create a secret vault using: Register-SecretVault"
            } else {
                Write-Warning "Secret vault '$($script:SPOFactoryConfig.CredentialVault)' not found. MSP credential management will be limited."
                Write-Host "To enable full MSP functionality, create a secret vault using: Register-SecretVault" -ForegroundColor Yellow
            }
        }
    }
}
catch {
    if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
        Write-PSFMessage -Level Warning -Message "Could not validate secret vault: $_"
    } else {
        Write-Warning "Could not validate secret vault: $_"
    }
}

#endregion

Write-Host "SPOSiteFactory MSP Module loaded successfully!" -ForegroundColor Green
Write-Host "Ready to manage $($script:SPOFactoryConstants.MaxTenantCount)+ SharePoint Online tenants" -ForegroundColor Cyan

# Display initialization summary
if (Get-Command Write-PSFMessage -ErrorAction SilentlyContinue) {
    Write-PSFMessage -Level Host -Message @"
SPOSiteFactory Initialization Summary:
- Configuration Path: $($script:SPOFactoryConfig.ConfigPath)
- Log Path: $($script:SPOFactoryConfig.LogPath)
- Credential Vault: $($script:SPOFactoryConfig.CredentialVault)
- Max Concurrent Connections: $($script:SPOFactoryConfig.MaxConcurrentConnections)
- Batch Size: $($script:SPOFactoryConfig.BatchSize)
- Debug Logging: $($script:SPOFactoryConfig.EnableDebugLogging)
"@
} else {
    Write-Host @"
SPOSiteFactory Initialization Summary:
- Configuration Path: $($script:SPOFactoryConfig.ConfigPath)
- Log Path: $($script:SPOFactoryConfig.LogPath)
- Credential Vault: $($script:SPOFactoryConfig.CredentialVault)
- Max Concurrent Connections: $($script:SPOFactoryConfig.MaxConcurrentConnections)
- Batch Size: $($script:SPOFactoryConfig.BatchSize)
- Debug Logging: $($script:SPOFactoryConfig.EnableDebugLogging)
"@ -ForegroundColor Cyan
}

#endregion

#region Export Module Members

# Get all public function names by parsing the actual function definitions
$functionsToExport = @()

# Get all public .ps1 files
$publicFiles = Get-ChildItem -Path "$PSScriptRoot\Public" -Filter "*.ps1" -Recurse -ErrorAction SilentlyContinue

foreach ($file in $publicFiles) {
    # Parse the file content to find function definitions
    $content = Get-Content -Path $file.FullName -Raw
    $functionPattern = 'function\s+([a-zA-Z0-9\-_]+)\s*\{'
    $matches = [regex]::Matches($content, $functionPattern)
    
    foreach ($match in $matches) {
        $functionName = $match.Groups[1].Value
        if ($functionName -and $functionName -notlike '*Test*Internal*') {
            $functionsToExport += $functionName
        }
    }
}

# Remove duplicates
$functionsToExport = $functionsToExport | Select-Object -Unique

# Export all public functions
if ($functionsToExport) {
    Write-Host "Exporting $($functionsToExport.Count) functions" -ForegroundColor Green
    Export-ModuleMember -Function $functionsToExport
}