function Write-SPOFactoryLog {
    <#
    .SYNOPSIS
        Writes log messages with MSP tenant isolation and compliance tracking.

    .DESCRIPTION
        Enterprise logging function designed for MSP environments managing multiple
        SharePoint Online tenants. Provides tenant isolation, audit trails, and
        compliance reporting capabilities using PSFramework.

    .PARAMETER Message
        The log message to write

    .PARAMETER Level
        The log level (Info, Warning, Error, Debug, Verbose, Critical, Host)

    .PARAMETER ClientName
        The client name for tenant isolation

    .PARAMETER Category
        Log category for filtering and reporting

    .PARAMETER FunctionName
        Name of the function generating the log entry

    .PARAMETER Tag
        Additional tags for log entry classification

    .PARAMETER Exception
        Exception object to include in error logs

    .PARAMETER EnableAuditLog
        Forces the message to be written to the audit log

    .PARAMETER Target
        Target object or identifier for the operation

    .EXAMPLE
        Write-SPOFactoryLog -Message "Site created successfully" -Level Info -ClientName "Contoso Corp" -Category "Provisioning"

    .EXAMPLE
        Write-SPOFactoryLog -Message "Failed to connect to tenant" -Level Error -ClientName "Contoso Corp" -Exception $_.Exception

    .EXAMPLE
        Write-SPOFactoryLog -Message "Security scan completed" -Level Info -ClientName "Contoso Corp" -Category "Security" -EnableAuditLog
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Verbose', 'Critical', 'Host')]
        [string]$Level = 'Info',
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'Provisioning', 'Security', 'Configuration', 'Hub', 'Audit', 'Performance', 'Error', 'System')]
        [string]$Category = 'System',
        
        [Parameter(Mandatory = $false)]
        [string]$FunctionName,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Tag,
        
        [Parameter(Mandatory = $false)]
        [System.Exception]$Exception,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnableAuditLog,
        
        [Parameter(Mandatory = $false)]
        [string]$Target
    )

    begin {
        # Get caller information if not provided
        if (-not $FunctionName) {
            $callerInfo = Get-PSCallStack | Select-Object -Skip 1 -First 1
            $FunctionName = $callerInfo.FunctionName
            if (-not $FunctionName -or $FunctionName -eq '<ScriptBlock>') {
                $FunctionName = 'Unknown'
            }
        }

        # Generate session ID for tracking related operations
        if (-not $script:SPOFactoryLogSessionId) {
            $script:SPOFactoryLogSessionId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        }

        # Prepare log data
        $timestamp = Get-Date
        $logData = @{
            Timestamp = $timestamp
            SessionId = $script:SPOFactoryLogSessionId
            Level = $Level
            Category = $Category
            FunctionName = $FunctionName
            ClientName = $ClientName
            Message = $Message
            Target = $Target
            Tag = $Tag
            Exception = $Exception
            ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            Username = $env:USERNAME
            ComputerName = $env:COMPUTERNAME
            ProcessId = $PID
        }
    }

    process {
        try {
            # Build comprehensive log message
            $logMessage = Build-SPOFactoryLogMessage -LogData $logData

            # Write to PSFramework with appropriate level
            $psfLevel = Convert-ToPSFLevel -Level $Level
            $psfMessage = @{
                Level = $psfLevel
                Message = $logMessage
                FunctionName = $FunctionName
                ModuleName = 'SPOSiteFactory'
            }

            # Add client-specific tags
            $psfTags = @()
            if ($ClientName) {
                $psfTags += "Client:$ClientName"
            }
            if ($Category) {
                $psfTags += "Category:$Category"
            }
            if ($Tag) {
                $psfTags += $Tag
            }
            
            if ($psfTags.Count -gt 0) {
                $psfMessage.Tag = $psfTags
            }

            # Add target if specified
            if ($Target) {
                $psfMessage.Target = $Target
            }

            # Write to PSFramework
            Write-PSFMessage @psfMessage

            # Write to client-specific log file if client specified
            if ($ClientName) {
                Write-SPOFactoryClientLog -LogData $logData
            }

            # Write to audit log if required or enabled globally
            if ($EnableAuditLog -or $script:SPOFactoryConfig.EnableAuditLog) {
                Write-SPOFactoryAuditLog -LogData $logData
            }

            # Write performance metrics if applicable
            if ($Category -eq 'Performance') {
                Write-SPOFactoryPerformanceLog -LogData $logData
            }

            # Handle critical errors
            if ($Level -eq 'Critical') {
                Send-SPOFactoryAlert -LogData $logData
            }

            # Exception handling
            if ($Exception) {
                Write-PSFMessage -Level Error -Message "Exception Details: $($Exception.ToString())" -FunctionName $FunctionName -ModuleName 'SPOSiteFactory'
                
                # Log exception to separate error log
                Write-SPOFactoryErrorLog -LogData $logData -Exception $Exception
            }
        }
        catch {
            # Fallback logging in case PSFramework fails
            $fallbackMessage = "[$timestamp] [$Level] [$FunctionName] $Message"
            if ($ClientName) {
                $fallbackMessage = "[$timestamp] [$Level] [$ClientName] [$FunctionName] $Message"
            }
            
            try {
                Write-Host $fallbackMessage -ForegroundColor Red
                
                # Try to write to a basic log file
                $fallbackLogPath = Join-Path $script:SPOFactoryConfig.LogPath "SPOFactory-Error-$(Get-Date -Format 'yyyy-MM-dd').log"
                $fallbackMessage | Out-File -FilePath $fallbackLogPath -Append -Encoding UTF8
            }
            catch {
                # Last resort - just write to console
                Write-Warning "Logging system failed: $_"
                Write-Warning "Original message: $Message"
            }
        }
    }
}

function Build-SPOFactoryLogMessage {
    <#
    .SYNOPSIS
        Builds a comprehensive log message from log data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData
    )

    $messageBuilder = [System.Text.StringBuilder]::new()
    
    # Add core message
    [void]$messageBuilder.Append($LogData.Message)
    
    # Add client context
    if ($LogData.ClientName) {
        [void]$messageBuilder.Append(" | Client: $($LogData.ClientName)")
    }
    
    # Add target context
    if ($LogData.Target) {
        [void]$messageBuilder.Append(" | Target: $($LogData.Target)")
    }
    
    # Add session tracking
    if ($LogData.SessionId) {
        [void]$messageBuilder.Append(" | Session: $($LogData.SessionId)")
    }
    
    # Add category
    if ($LogData.Category) {
        [void]$messageBuilder.Append(" | Category: $($LogData.Category)")
    }
    
    return $messageBuilder.ToString()
}

function Write-SPOFactoryClientLog {
    <#
    .SYNOPSIS
        Writes log entry to client-specific log file.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData
    )

    try {
        if (-not $LogData.ClientName) {
            return
        }

        # Create client-specific log directory
        $clientLogPath = Join-Path $script:SPOFactoryConfig.LogPath "Clients\$($LogData.ClientName)"
        if (-not (Test-Path $clientLogPath)) {
            New-Item -Path $clientLogPath -ItemType Directory -Force | Out-Null
        }

        # Client log file with date rotation
        $clientLogFile = Join-Path $clientLogPath "SPOFactory-$($LogData.ClientName)-$(Get-Date -Format 'yyyy-MM-dd').log"
        
        # Format client log entry
        $logEntry = @{
            Timestamp = $LogData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff')
            Level = $LogData.Level
            Category = $LogData.Category
            Function = $LogData.FunctionName
            Message = $LogData.Message
            Target = $LogData.Target
            SessionId = $LogData.SessionId
            User = $LogData.Username
            Computer = $LogData.ComputerName
        }

        # Convert to JSON and write
        $jsonEntry = $logEntry | ConvertTo-Json -Compress
        $jsonEntry | Out-File -FilePath $clientLogFile -Append -Encoding UTF8

        # Manage log file rotation
        Manage-SPOFactoryLogRotation -LogPath $clientLogPath -RetentionDays $script:SPOFactoryConstants.LogRetentionDays
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to write client log for $($LogData.ClientName): $_"
    }
}

function Write-SPOFactoryAuditLog {
    <#
    .SYNOPSIS
        Writes entry to compliance audit log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData
    )

    try {
        # Create audit log directory
        $auditLogPath = Join-Path $script:SPOFactoryConfig.LogPath "Audit"
        if (-not (Test-Path $auditLogPath)) {
            New-Item -Path $auditLogPath -ItemType Directory -Force | Out-Null
        }

        # Audit log file with monthly rotation
        $auditLogFile = Join-Path $auditLogPath "SPOFactory-Audit-$(Get-Date -Format 'yyyy-MM').log"
        
        # Format audit log entry
        $auditEntry = @{
            Timestamp = $LogData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff')
            EventId = [System.Guid]::NewGuid().ToString()
            Level = $LogData.Level
            Category = $LogData.Category
            Function = $LogData.FunctionName
            ClientName = $LogData.ClientName
            Message = $LogData.Message
            Target = $LogData.Target
            SessionId = $LogData.SessionId
            User = $LogData.Username
            Computer = $LogData.ComputerName
            ModuleVersion = $LogData.ModuleVersion
            PowerShellVersion = $LogData.PowerShellVersion
            ProcessId = $LogData.ProcessId
        }

        # Convert to JSON and write
        $jsonEntry = $auditEntry | ConvertTo-Json -Compress
        $jsonEntry | Out-File -FilePath $auditLogFile -Append -Encoding UTF8
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to write audit log: $_"
    }
}

function Write-SPOFactoryPerformanceLog {
    <#
    .SYNOPSIS
        Writes performance metrics to specialized log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData
    )

    try {
        # Create performance log directory
        $perfLogPath = Join-Path $script:SPOFactoryConfig.LogPath "Performance"
        if (-not (Test-Path $perfLogPath)) {
            New-Item -Path $perfLogPath -ItemType Directory -Force | Out-Null
        }

        # Performance log file with daily rotation
        $perfLogFile = Join-Path $perfLogPath "SPOFactory-Performance-$(Get-Date -Format 'yyyy-MM-dd').log"
        
        # Format performance log entry
        $perfEntry = @{
            Timestamp = $LogData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ClientName = $LogData.ClientName
            Function = $LogData.FunctionName
            Message = $LogData.Message
            Target = $LogData.Target
            SessionId = $LogData.SessionId
            Computer = $LogData.ComputerName
        }

        # Convert to CSV format for easier analysis
        $csvEntry = "$($perfEntry.Timestamp),$($perfEntry.ClientName),$($perfEntry.Function),$($perfEntry.Message),$($perfEntry.Target),$($perfEntry.SessionId),$($perfEntry.Computer)"
        $csvEntry | Out-File -FilePath $perfLogFile -Append -Encoding UTF8
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to write performance log: $_"
    }
}

function Write-SPOFactoryErrorLog {
    <#
    .SYNOPSIS
        Writes detailed error information to error log.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData,
        
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception
    )

    try {
        # Create error log directory
        $errorLogPath = Join-Path $script:SPOFactoryConfig.LogPath "Errors"
        if (-not (Test-Path $errorLogPath)) {
            New-Item -Path $errorLogPath -ItemType Directory -Force | Out-Null
        }

        # Error log file with daily rotation
        $errorLogFile = Join-Path $errorLogPath "SPOFactory-Errors-$(Get-Date -Format 'yyyy-MM-dd').log"
        
        # Format error log entry with full exception details
        $errorEntry = @{
            Timestamp = $LogData.Timestamp.ToString('yyyy-MM-dd HH:mm:ss.fff')
            ErrorId = [System.Guid]::NewGuid().ToString()
            ClientName = $LogData.ClientName
            Function = $LogData.FunctionName
            Message = $LogData.Message
            ExceptionType = $Exception.GetType().FullName
            ExceptionMessage = $Exception.Message
            StackTrace = $Exception.StackTrace
            InnerException = if ($Exception.InnerException) { $Exception.InnerException.Message } else { $null }
            Target = $LogData.Target
            SessionId = $LogData.SessionId
            User = $LogData.Username
            Computer = $LogData.ComputerName
            ModuleVersion = $LogData.ModuleVersion
            PowerShellVersion = $LogData.PowerShellVersion
        }

        # Convert to JSON and write
        $jsonEntry = $errorEntry | ConvertTo-Json -Depth 5
        $jsonEntry | Out-File -FilePath $errorLogFile -Append -Encoding UTF8
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to write error log: $_"
    }
}

function Send-SPOFactoryAlert {
    <#
    .SYNOPSIS
        Sends alert for critical events.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$LogData
    )

    try {
        if (-not $script:SPOFactoryConfig.AlertEmail) {
            Write-PSFMessage -Level Debug -Message "No alert email configured, skipping alert notification"
            return
        }

        # This would integrate with MSP's alerting system
        # For now, just log the critical event
        Write-PSFMessage -Level Critical -Message "CRITICAL ALERT: $($LogData.Message)" -FunctionName $LogData.FunctionName
        
        # Future implementation could include:
        # - Email notifications
        # - SNMP traps
        # - Webhook notifications
        # - Integration with MSP monitoring tools
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to send alert: $_"
    }
}

function Convert-ToPSFLevel {
    <#
    .SYNOPSIS
        Converts log level to PSFramework level.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Level
    )

    switch ($Level) {
        'Debug' { return 'Debug' }
        'Verbose' { return 'Verbose' }
        'Info' { return 'Host' }
        'Warning' { return 'Warning' }
        'Error' { return 'Warning' }
        'Critical' { return 'Critical' }
        'Host' { return 'Host' }
        default { return 'Host' }
    }
}

function Manage-SPOFactoryLogRotation {
    <#
    .SYNOPSIS
        Manages log file rotation and cleanup.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogPath,
        
        [Parameter(Mandatory = $false)]
        [int]$RetentionDays = 90
    )

    try {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $oldLogFiles = Get-ChildItem -Path $LogPath -Filter "*.log" | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        
        foreach ($logFile in $oldLogFiles) {
            try {
                Remove-Item -Path $logFile.FullName -Force
                Write-PSFMessage -Level Debug -Message "Removed old log file: $($logFile.Name)"
            }
            catch {
                Write-PSFMessage -Level Warning -Message "Failed to remove old log file $($logFile.Name): $_"
            }
        }
    }
    catch {
        Write-PSFMessage -Level Warning -Message "Failed to manage log rotation for $LogPath`: $_"
    }
}