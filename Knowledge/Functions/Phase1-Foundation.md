# Phase 1: Module Foundation Functions

## Overview

Phase 1 functions provide the core foundation for the SPOSiteFactory module, including connection management, logging, error handling, and module configuration.

## Connection Management

### Connect-SPOFactory

Establishes a connection to SharePoint Online with support for multiple authentication methods.

**Syntax:**
```powershell
Connect-SPOFactory -TenantUrl <String> -ClientName <String> 
                   [-AuthenticationMethod <String>] 
                   [-CertificateThumbprint <String>]
                   [-ClientId <String>]
                   [-StoredCredential <String>]
                   [-Interactive]
                   [-DeviceCode]
                   [-RetryCount <Int32>]
                   [-RetryDelay <Int32>]
                   [-Force]
```

**Parameters:**
- **TenantUrl** (Required): SharePoint admin center URL (https://tenant-admin.sharepoint.com)
- **ClientName** (Required): Identifier for MSP client isolation
- **AuthenticationMethod**: Interactive, Certificate, AppOnly, StoredCredential, DeviceCode
- **CertificateThumbprint**: Certificate thumbprint for app-only auth
- **ClientId**: Azure AD application ID
- **StoredCredential**: Name of stored credential in SecretStore
- **Interactive**: Use interactive browser authentication
- **DeviceCode**: Use device code flow for authentication
- **RetryCount**: Number of connection retry attempts (default: 3)
- **RetryDelay**: Delay between retries in seconds (default: 5)
- **Force**: Force new connection even if one exists

**Examples:**
```powershell
# Interactive authentication
Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" `
                   -ClientName "Contoso" `
                   -Interactive

# Certificate-based authentication
Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" `
                   -ClientName "Contoso" `
                   -AuthenticationMethod "Certificate" `
                   -CertificateThumbprint "1234567890ABCDEF" `
                   -ClientId "client-id-guid"

# Stored credential
Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" `
                   -ClientName "Contoso" `
                   -StoredCredential "ContosoSPOAdmin"

# Device code flow (for restricted environments)
Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" `
                   -ClientName "Contoso" `
                   -DeviceCode
```

### Disconnect-SPOFactory

Closes the SharePoint connection and cleans up resources.

**Syntax:**
```powershell
Disconnect-SPOFactory -ClientName <String> [-Force]
```

**Examples:**
```powershell
# Disconnect specific client
Disconnect-SPOFactory -ClientName "Contoso"

# Force disconnect all clients
Disconnect-SPOFactory -ClientName "All" -Force
```

### Test-SPOFactoryConnection

Verifies the current connection status and validates permissions.

**Syntax:**
```powershell
Test-SPOFactoryConnection -ClientName <String> [-Detailed]
```

**Returns:**
```powershell
@{
    IsConnected = $true
    ClientName = "Contoso"
    TenantUrl = "https://contoso-admin.sharepoint.com"
    AuthMethod = "Interactive"
    ConnectedAt = [DateTime]
    Permissions = @("Sites.FullControl.All")
}
```

**Examples:**
```powershell
# Basic connection test
$connection = Test-SPOFactoryConnection -ClientName "Contoso"
if ($connection.IsConnected) {
    Write-Host "Connected to $($connection.TenantUrl)"
}

# Detailed connection info
Test-SPOFactoryConnection -ClientName "Contoso" -Detailed | Format-List
```

## Logging Framework

### Write-SPOFactoryLog

Writes structured logs using PSFramework with tenant isolation.

**Syntax:**
```powershell
Write-SPOFactoryLog -Message <String> 
                    -Level <String> 
                    [-ClientName <String>]
                    [-Category <String>]
                    [-Tag <String[]>]
                    [-Exception <Exception>]
                    [-Data <Hashtable>]
                    [-EnableAuditLog]
```

**Parameters:**
- **Message** (Required): Log message text
- **Level**: Critical, Error, Warning, Info, Debug, Verbose (default: Info)
- **ClientName**: Client identifier for log isolation
- **Category**: Log category (Connection, Provisioning, Security, etc.)
- **Tag**: Array of tags for log filtering
- **Exception**: Exception object to log
- **Data**: Additional structured data
- **EnableAuditLog**: Write to audit log for compliance

**Examples:**
```powershell
# Basic logging
Write-SPOFactoryLog -Message "Site creation started" -Level Info

# Error with exception
try {
    # Some operation
} catch {
    Write-SPOFactoryLog -Message "Operation failed" `
                        -Level Error `
                        -Exception $_.Exception `
                        -ClientName "Contoso"
}

# Audit logging
Write-SPOFactoryLog -Message "User created site: Marketing" `
                    -Level Info `
                    -ClientName "Contoso" `
                    -Category "Provisioning" `
                    -Tag @("SiteCreation", "Audit") `
                    -EnableAuditLog `
                    -Data @{
                        SiteUrl = "https://contoso.sharepoint.com/sites/marketing"
                        Owner = "admin@contoso.com"
                        Template = "TeamSite"
                    }
```

### Get-SPOFactoryLog

Retrieves logs with filtering capabilities.

**Syntax:**
```powershell
Get-SPOFactoryLog [-ClientName <String>]
                  [-Level <String>]
                  [-Category <String>]
                  [-StartTime <DateTime>]
                  [-EndTime <DateTime>]
                  [-Last <Int32>]
```

**Examples:**
```powershell
# Get last 10 error logs
Get-SPOFactoryLog -Level Error -Last 10

# Get logs for specific client today
Get-SPOFactoryLog -ClientName "Contoso" `
                  -StartTime (Get-Date).Date

# Export logs to CSV
Get-SPOFactoryLog -ClientName "Contoso" `
                  -Category "Security" `
                  -Last 100 | Export-Csv "security-audit.csv"
```

## Error Handling

### Invoke-SPOFactoryCommand

Executes commands with retry logic and comprehensive error handling.

**Syntax:**
```powershell
Invoke-SPOFactoryCommand -ScriptBlock <ScriptBlock>
                        [-ClientName <String>]
                        [-Category <String>]
                        [-RetryCount <Int32>]
                        [-RetryDelay <Int32>]
                        [-ErrorMessage <String>]
                        [-CriticalOperation]
                        [-SuppressErrors]
```

**Parameters:**
- **ScriptBlock** (Required): Code to execute with error handling
- **RetryCount**: Number of retry attempts (default: 3)
- **RetryDelay**: Delay between retries in seconds (default: 2)
- **ErrorMessage**: Custom error message
- **CriticalOperation**: Fail immediately without retries
- **SuppressErrors**: Suppress non-critical errors

**Examples:**
```powershell
# Basic command with retry
$result = Invoke-SPOFactoryCommand -ScriptBlock {
    Get-PnPSite -Identity "https://contoso.sharepoint.com/sites/test"
} -ClientName "Contoso" -Category "Query"

# Critical operation (no retry)
$site = Invoke-SPOFactoryCommand -ScriptBlock {
    New-PnPSite -Type TeamSite -Title "Finance" -Url "finance"
} -ClientName "Contoso" `
  -Category "Provisioning" `
  -CriticalOperation `
  -ErrorMessage "Failed to create Finance site"

# With custom retry logic
Invoke-SPOFactoryCommand -ScriptBlock {
    # Potentially throttled operation
    Get-PnPListItem -List "Large List" -PageSize 5000
} -RetryCount 5 -RetryDelay 10
```

## Configuration Management

### Get-SPOFactoryConfig

Retrieves module configuration settings.

**Syntax:**
```powershell
Get-SPOFactoryConfig [-ClientName <String>] [-Setting <String>]
```

**Returns:**
```powershell
@{
    DefaultClientName = "Contoso"
    DefaultSecurityBaseline = "MSPStandard"
    LogPath = "C:\Logs\SPOSiteFactory"
    LogLevel = "Info"
    LogRetentionDays = 90
    ConnectionPoolSize = 50
    DefaultTimeZone = 13
    MSPMode = $true
    Clients = @{
        "Contoso" = @{...}
        "Fabrikam" = @{...}
    }
}
```

**Examples:**
```powershell
# Get all configuration
$config = Get-SPOFactoryConfig
$config | Format-List

# Get specific client config
Get-SPOFactoryConfig -ClientName "Contoso"

# Get specific setting
$logPath = Get-SPOFactoryConfig -Setting "LogPath"
```

### Set-SPOFactoryConfig

Updates module configuration settings.

**Syntax:**
```powershell
Set-SPOFactoryConfig [-DefaultClientName <String>]
                     [-DefaultSecurityBaseline <String>]
                     [-LogPath <String>]
                     [-LogLevel <String>]
                     [-LogRetentionDays <Int32>]
                     [-ClientName <String>]
                     [-Settings <Hashtable>]
```

**Examples:**
```powershell
# Set global defaults
Set-SPOFactoryConfig -DefaultClientName "Contoso" `
                     -DefaultSecurityBaseline "MSPSecure" `
                     -LogPath "D:\Logs\SPO" `
                     -LogRetentionDays 180

# Set client-specific configuration
Set-SPOFactoryConfig -ClientName "Fabrikam" -Settings @{
    SecurityBaseline = "MSPStrict"
    DefaultTimeZone = 10
    RequireMFA = $true
    NamingPrefix = "FAB"
}
```

## Module Initialization

### Initialize-SPOFactory

Initializes the module environment and validates prerequisites.

**Syntax:**
```powershell
Initialize-SPOFactory [-CheckDependencies] [-CreateFolders] [-LoadSecrets]
```

**Examples:**
```powershell
# Full initialization
Initialize-SPOFactory -CheckDependencies -CreateFolders -LoadSecrets

# Check only
Initialize-SPOFactory -CheckDependencies
```

### Get-SPOFactoryInfo

Retrieves module information and statistics.

**Syntax:**
```powershell
Get-SPOFactoryInfo [-Detailed]
```

**Returns:**
```powershell
@{
    Version = "1.0.0"
    InstalledAt = [DateTime]
    TotalClients = 5
    ActiveConnections = 2
    TotalSitesCreated = 150
    ModulePath = "C:\Program Files\..."
    Dependencies = @{
        "PnP.PowerShell" = "3.0.0"
        "PSFramework" = "1.7.0"
    }
}
```

## Helper Functions

### Convert-SPOFactoryUrl

Converts and validates SharePoint URLs.

**Syntax:**
```powershell
Convert-SPOFactoryUrl -Url <String> 
                      [-ClientName <String>]
                      [-AddPrefix]
                      [-EnsureHttps]
```

**Examples:**
```powershell
# Add client prefix
$url = Convert-SPOFactoryUrl -Url "marketing" `
                             -ClientName "Contoso" `
                             -AddPrefix
# Result: "contoso-marketing"

# Ensure HTTPS
$fullUrl = Convert-SPOFactoryUrl -Url "http://sharepoint.com/sites/test" `
                                  -EnsureHttps
# Result: "https://sharepoint.com/sites/test"
```

### Test-SPOFactoryPrerequisites

Validates all module prerequisites.

**Syntax:**
```powershell
Test-SPOFactoryPrerequisites [-Detailed] [-FixIssues]
```

**Examples:**
```powershell
# Check prerequisites
$prereq = Test-SPOFactoryPrerequisites -Detailed
if (-not $prereq.AllMet) {
    Write-Warning "Missing prerequisites: $($prereq.Missing -join ', ')"
}

# Auto-fix issues
Test-SPOFactoryPrerequisites -FixIssues
```

## Best Practices

### Connection Management
1. Always use client names for multi-tenant isolation
2. Store credentials securely using SecretManagement
3. Implement connection pooling for large operations
4. Use certificate auth for unattended scenarios

### Logging
1. Use appropriate log levels (Error for failures, Info for operations)
2. Include structured data for analysis
3. Enable audit logging for compliance-required operations
4. Regularly rotate logs based on retention policy

### Error Handling
1. Use Invoke-SPOFactoryCommand for all SharePoint operations
2. Implement appropriate retry logic for transient failures
3. Log all errors with full context
4. Provide meaningful error messages to users

### Configuration
1. Set organization defaults in module config
2. Override with client-specific settings
3. Use environment variables for sensitive settings
4. Document all configuration changes

## Troubleshooting

### Connection Issues
```powershell
# Verbose connection diagnostics
$VerbosePreference = "Continue"
Connect-SPOFactory -TenantUrl "https://tenant-admin.sharepoint.com" `
                   -ClientName "Test" `
                   -Interactive `
                   -Force

# Check connection details
Test-SPOFactoryConnection -ClientName "Test" -Detailed
```

### Logging Issues
```powershell
# Check log configuration
Get-SPOFactoryConfig | Select-Object Log*

# Test log writing
Write-SPOFactoryLog -Message "Test log entry" -Level Info
Get-SPOFactoryLog -Last 1
```

### Module Loading Issues
```powershell
# Reimport module
Remove-Module SPOSiteFactory -Force -ErrorAction SilentlyContinue
Import-Module SPOSiteFactory -Force -Verbose

# Check loaded functions
Get-Command -Module SPOSiteFactory | Measure-Object
```

---

**Note**: Phase 1 functions form the foundation for all other module operations. Ensure these are working correctly before using Phase 2+ functions.