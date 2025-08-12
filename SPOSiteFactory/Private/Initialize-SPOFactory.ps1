function Initialize-SPOFactory {
    <#
    .SYNOPSIS
        Initializes the SPOSiteFactory module environment.
    
    .DESCRIPTION
        Sets up module-level variables, connection cache, and validates prerequisites.
    #>
    
    [CmdletBinding()]
    param()
    
    # Initialize connection cache
    if (-not $script:SPOConnections) {
        $script:SPOConnections = @{}
        if (Get-Command Write-SPOFactoryLog -ErrorAction SilentlyContinue) {
            Write-SPOFactoryLog -Message "Connection cache initialized" -Level Debug
        }
    }
    
    # Initialize current connection tracker
    if (-not $script:CurrentSPOConnection) {
        $script:CurrentSPOConnection = $null
    }
    
    # Check for existing environment variable
    if ($env:SPOFactoryCurrentClient) {
        $script:CurrentSPOConnection = $env:SPOFactoryCurrentClient
        if (Get-Command Write-SPOFactoryLog -ErrorAction SilentlyContinue) {
            Write-SPOFactoryLog -Message "Restored current client from environment: $($env:SPOFactoryCurrentClient)" -Level Debug
        }
    }
    
    # Initialize configuration if not exists
    if (-not $script:SPOFactoryConfig) {
        $script:SPOFactoryConfig = @{
            DefaultSecurityBaseline = "MSPStandard"
            DefaultTimeZone = 13
            MaxConcurrentConnections = 50
            RetryCount = 3
            RetryDelay = 2
            LogLevel = "Info"
        }
    }
    
    if (Get-Command Write-SPOFactoryLog -ErrorAction SilentlyContinue) {
        Write-SPOFactoryLog -Message "SPOFactory initialization complete" -Level Debug
    }
}

# Call initialization when module loads
Initialize-SPOFactory