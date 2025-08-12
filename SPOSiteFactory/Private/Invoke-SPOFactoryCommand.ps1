function Invoke-SPOFactoryCommand {
    <#
    .SYNOPSIS
        Executes commands with MSP-specific error handling, retry logic, and tenant isolation.

    .DESCRIPTION
        Enterprise-grade command wrapper designed for MSP environments managing multiple
        SharePoint Online tenants. Provides error classification, retry logic with
        exponential backoff, detailed error reporting, and tenant-specific error isolation.

    .PARAMETER ScriptBlock
        The script block to execute

    .PARAMETER ErrorMessage
        Custom error message prefix for failed operations

    .PARAMETER MaxRetries
        Maximum number of retry attempts (default: configured value)

    .PARAMETER RetryDelay
        Initial retry delay in seconds (default: 1 second)

    .PARAMETER ClientName
        Client name for tenant-specific error isolation

    .PARAMETER Category
        Operation category for error classification

    .PARAMETER ThrottleRetry
        Enable special handling for throttling scenarios

    .PARAMETER CriticalOperation
        Mark as critical operation for enhanced error reporting

    .PARAMETER SuppressErrors
        Suppress error output (still logs errors)

    .PARAMETER PassThru
        Return the result of the script block execution

    .EXAMPLE
        Invoke-SPOFactoryCommand -ScriptBlock { Get-PnPWeb } -ErrorMessage "Failed to get web" -ClientName "Contoso Corp"

    .EXAMPLE
        $result = Invoke-SPOFactoryCommand -ScriptBlock { Get-PnPSite } -ClientName "Contoso Corp" -MaxRetries 5 -PassThru

    .EXAMPLE
        Invoke-SPOFactoryCommand -ScriptBlock { Set-PnPWeb -Title "New Title" } -ClientName "Contoso Corp" -Category "Configuration" -CriticalOperation
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorMessage = "Operation failed",
        
        [Parameter(Mandatory = $false)]
        [int]$MaxRetries = $script:SPOFactoryConfig.RetryAttempts,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryDelay = 1,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Connection', 'Provisioning', 'Security', 'Configuration', 'Hub', 'Audit', 'Performance', 'System')]
        [string]$Category = 'System',
        
        [Parameter(Mandatory = $false)]
        [switch]$ThrottleRetry,
        
        [Parameter(Mandatory = $false)]
        [switch]$CriticalOperation,
        
        [Parameter(Mandatory = $false)]
        [switch]$SuppressErrors,
        
        [Parameter(Mandatory = $false)]
        [switch]$PassThru
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        $attempt = 0
        $result = $null
        $lastError = $null
        $startTime = Get-Date

        Write-SPOFactoryLog -Message "Starting operation: $ErrorMessage" -Level Debug -ClientName $ClientName -Category $Category -Tag @('OperationStart', $operationId)
    }

    process {
        while ($attempt -le $MaxRetries) {
            $attempt++
            
            try {
                Write-SPOFactoryLog -Message "Executing operation (Attempt $attempt/$($MaxRetries + 1))" -Level Debug -ClientName $ClientName -Category $Category -Tag @('OperationAttempt', $operationId)
                
                # Execute the script block
                $result = & $ScriptBlock
                
                # Calculate execution time
                $executionTime = (Get-Date) - $startTime
                Write-SPOFactoryLog -Message "Operation completed successfully in $($executionTime.TotalMilliseconds)ms" -Level Info -ClientName $ClientName -Category 'Performance' -Tag @('OperationSuccess', $operationId)
                
                # Return result if requested
                if ($PassThru) {
                    return $result
                }
                return
            }
            catch {
                $lastError = $_
                $errorInfo = Get-SPOFactoryErrorClassification -Exception $_.Exception -ClientName $ClientName
                
                # Log the error with classification
                Write-SPOFactoryLog -Message "Operation failed (Attempt $attempt/$($MaxRetries + 1)): $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category $Category -Exception $_.Exception -Tag @('OperationError', $operationId, $errorInfo.Classification)
                
                # Check if we should retry
                if ($attempt -le $MaxRetries -and $errorInfo.ShouldRetry) {
                    $delay = Get-SPOFactoryRetryDelay -Attempt $attempt -BaseDelay $RetryDelay -ErrorType $errorInfo.Classification -ThrottleRetry $ThrottleRetry
                    
                    Write-SPOFactoryLog -Message "Retrying in $delay seconds (Classification: $($errorInfo.Classification))" -Level Info -ClientName $ClientName -Category $Category -Tag @('OperationRetry', $operationId)
                    
                    Start-Sleep -Seconds $delay
                    continue
                }
                else {
                    # Final failure
                    $executionTime = (Get-Date) - $startTime
                    $finalErrorMessage = "$ErrorMessage after $attempt attempts in $($executionTime.TotalSeconds) seconds"
                    
                    # Enhanced error reporting for MSP scenarios
                    $errorReport = New-SPOFactoryErrorReport -Exception $_.Exception -ClientName $ClientName -Category $Category -OperationId $operationId -Attempts $attempt -ExecutionTime $executionTime -ErrorInfo $errorInfo
                    
                    # Log final failure
                    $logLevel = if ($CriticalOperation) { 'Critical' } else { 'Error' }
                    Write-SPOFactoryLog -Message $finalErrorMessage -Level $logLevel -ClientName $ClientName -Category $Category -Exception $_.Exception -Tag @('OperationFailed', $operationId, $errorInfo.Classification)
                    
                    # Store error for MSP reporting
                    Add-SPOFactoryErrorToRegistry -ErrorReport $errorReport
                    
                    # Send alert for critical operations
                    if ($CriticalOperation) {
                        Send-SPOFactoryCriticalAlert -ErrorReport $errorReport
                    }
                    
                    # Re-throw error unless suppressed
                    if (-not $SuppressErrors) {
                        throw [SPOFactoryException]::new($finalErrorMessage, $_.Exception, $errorInfo.Classification, $ClientName, $errorReport)
                    }
                    
                    return $null
                }
            }
        }
    }

    end {
        if ($result -ne $null -and $PassThru) {
            return $result
        }
    }
}

function Get-SPOFactoryErrorClassification {
    <#
    .SYNOPSIS
        Classifies errors for MSP-specific handling and retry logic.

    .DESCRIPTION
        Analyzes exceptions to determine the appropriate error classification,
        retry strategy, and MSP-specific handling requirements.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $classification = @{
        Classification = 'Unknown'
        ShouldRetry = $false
        RetryMultiplier = 1.0
        Severity = 'Medium'
        MSPAction = 'Review'
        TenantSpecific = $true
    }

    $exceptionMessage = $Exception.Message.ToLower()
    $exceptionType = $Exception.GetType().Name

    # SharePoint-specific error classifications
    switch -Regex ($exceptionMessage) {
        'throttl|rate limit|429' {
            $classification.Classification = 'Throttling'
            $classification.ShouldRetry = $true
            $classification.RetryMultiplier = 2.0
            $classification.Severity = 'Low'
            $classification.MSPAction = 'Monitor'
            break
        }

        'timeout|timed out' {
            $classification.Classification = 'Timeout'
            $classification.ShouldRetry = $true
            $classification.RetryMultiplier = 1.5
            $classification.Severity = 'Medium'
            $classification.MSPAction = 'Monitor'
            break
        }

        'unauthorized|401|access denied|forbidden|403' {
            $classification.Classification = 'Authorization'
            $classification.ShouldRetry = $false
            $classification.Severity = 'High'
            $classification.MSPAction = 'CheckCredentials'
            break
        }

        'not found|404|does not exist' {
            $classification.Classification = 'NotFound'
            $classification.ShouldRetry = $false
            $classification.Severity = 'Medium'
            $classification.MSPAction = 'Verify'
            break
        }

        'network|connection|dns|resolve' {
            $classification.Classification = 'Network'
            $classification.ShouldRetry = $true
            $classification.RetryMultiplier = 1.2
            $classification.Severity = 'Medium'
            $classification.MSPAction = 'CheckConnectivity'
            break
        }

        'tenant|subscription|license' {
            $classification.Classification = 'TenantConfig'
            $classification.ShouldRetry = $false
            $classification.Severity = 'High'
            $classification.MSPAction = 'ReviewTenantConfig'
            break
        }

        'quota|storage|limit exceeded' {
            $classification.Classification = 'QuotaExceeded'
            $classification.ShouldRetry = $false
            $classification.Severity = 'High'
            $classification.MSPAction = 'ReviewQuota'
            break
        }

        'sharepoint is not accessible|service unavailable|503|502|500' {
            $classification.Classification = 'ServiceUnavailable'
            $classification.ShouldRetry = $true
            $classification.RetryMultiplier = 3.0
            $classification.Severity = 'High'
            $classification.MSPAction = 'CheckServiceHealth'
            break
        }

        'invalid|bad request|400' {
            $classification.Classification = 'InvalidRequest'
            $classification.ShouldRetry = $false
            $classification.Severity = 'Medium'
            $classification.MSPAction = 'ReviewParameters'
            break
        }

        'conflict|409|already exists' {
            $classification.Classification = 'Conflict'
            $classification.ShouldRetry = $false
            $classification.Severity = 'Low'
            $classification.MSPAction = 'ReviewExisting'
            break
        }

        'certificate|ssl|tls' {
            $classification.Classification = 'Certificate'
            $classification.ShouldRetry = $true
            $classification.RetryMultiplier = 1.0
            $classification.Severity = 'High'
            $classification.MSPAction = 'CheckCertificate'
            break
        }

        default {
            # Analyze by exception type
            switch ($exceptionType) {
                'ArgumentException' {
                    $classification.Classification = 'InvalidArgument'
                    $classification.ShouldRetry = $false
                    $classification.Severity = 'Medium'
                    $classification.MSPAction = 'ReviewParameters'
                }
                
                'UnauthorizedAccessException' {
                    $classification.Classification = 'Authorization'
                    $classification.ShouldRetry = $false
                    $classification.Severity = 'High'
                    $classification.MSPAction = 'CheckPermissions'
                }
                
                'TimeoutException' {
                    $classification.Classification = 'Timeout'
                    $classification.ShouldRetry = $true
                    $classification.RetryMultiplier = 1.5
                    $classification.Severity = 'Medium'
                    $classification.MSPAction = 'Monitor'
                }
                
                default {
                    $classification.Classification = 'General'
                    $classification.ShouldRetry = $true
                    $classification.RetryMultiplier = 1.0
                    $classification.Severity = 'Medium'
                    $classification.MSPAction = 'Investigate'
                }
            }
        }
    }

    return $classification
}

function Get-SPOFactoryRetryDelay {
    <#
    .SYNOPSIS
        Calculates retry delay with exponential backoff and jitter.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$Attempt,
        
        [Parameter(Mandatory = $false)]
        [int]$BaseDelay = 1,
        
        [Parameter(Mandatory = $false)]
        [string]$ErrorType = 'General',
        
        [Parameter(Mandatory = $false)]
        [switch]$ThrottleRetry
    )

    # Base exponential backoff
    $delay = $BaseDelay * [Math]::Pow(2, $Attempt - 1)

    # Apply error-type specific multipliers
    switch ($ErrorType) {
        'Throttling' {
            $delay *= 3  # Longer delays for throttling
            if ($ThrottleRetry) {
                $delay *= 2  # Even longer if explicitly handling throttling
            }
        }
        'ServiceUnavailable' {
            $delay *= 2
        }
        'Network' {
            $delay *= 1.5
        }
    }

    # Add jitter to prevent thundering herd
    $jitter = Get-Random -Minimum 0.8 -Maximum 1.2
    $delay = [int]($delay * $jitter)

    # Cap maximum delay at 5 minutes
    $maxDelay = 300
    if ($delay -gt $maxDelay) {
        $delay = $maxDelay
    }

    return $delay
}

function New-SPOFactoryErrorReport {
    <#
    .SYNOPSIS
        Creates comprehensive error report for MSP scenarios.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$Category,
        
        [Parameter(Mandatory = $false)]
        [string]$OperationId,
        
        [Parameter(Mandatory = $false)]
        [int]$Attempts,
        
        [Parameter(Mandatory = $false)]
        [timespan]$ExecutionTime,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$ErrorInfo
    )

    return @{
        ErrorId = [System.Guid]::NewGuid().ToString()
        Timestamp = Get-Date
        ClientName = $ClientName
        Category = $Category
        OperationId = $OperationId
        Attempts = $Attempts
        ExecutionTime = $ExecutionTime
        Exception = @{
            Type = $Exception.GetType().FullName
            Message = $Exception.Message
            StackTrace = $Exception.StackTrace
            InnerException = if ($Exception.InnerException) { $Exception.InnerException.Message } else { $null }
        }
        Classification = $ErrorInfo.Classification
        Severity = $ErrorInfo.Severity
        MSPAction = $ErrorInfo.MSPAction
        TenantSpecific = $ErrorInfo.TenantSpecific
        Environment = @{
            ComputerName = $env:COMPUTERNAME
            UserName = $env:USERNAME
            PowerShellVersion = $PSVersionTable.PSVersion.ToString()
            ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
            ProcessId = $PID
        }
        Resolution = @{
            Suggested = Get-SPOFactoryErrorResolution -Classification $ErrorInfo.Classification
            DocumentationUrl = Get-SPOFactoryErrorDocumentationUrl -Classification $ErrorInfo.Classification
        }
    }
}

function Add-SPOFactoryErrorToRegistry {
    <#
    .SYNOPSIS
        Adds error to MSP error registry for tracking and reporting.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ErrorReport
    )

    try {
        # Create error registry directory
        $errorRegistryPath = Join-Path $script:SPOFactoryConfig.ConfigPath "ErrorRegistry"
        if (-not (Test-Path $errorRegistryPath)) {
            New-Item -Path $errorRegistryPath -ItemType Directory -Force | Out-Null
        }

        # Client-specific error tracking
        if ($ErrorReport.ClientName) {
            $clientErrorPath = Join-Path $errorRegistryPath "$($ErrorReport.ClientName)-$(Get-Date -Format 'yyyy-MM').json"
            
            # Load existing errors or create new array
            $clientErrors = @()
            if (Test-Path $clientErrorPath) {
                $clientErrors = Get-Content $clientErrorPath -Raw | ConvertFrom-Json
            }
            
            # Add new error
            $clientErrors += $ErrorReport
            
            # Save updated errors
            $clientErrors | ConvertTo-Json -Depth 5 | Out-File -FilePath $clientErrorPath -Encoding UTF8
        }

        # Global error tracking
        $globalErrorPath = Join-Path $errorRegistryPath "Global-$(Get-Date -Format 'yyyy-MM').json"
        
        $globalErrors = @()
        if (Test-Path $globalErrorPath) {
            $globalErrors = Get-Content $globalErrorPath -Raw | ConvertFrom-Json
        }
        
        $globalErrors += $ErrorReport
        $globalErrors | ConvertTo-Json -Depth 5 | Out-File -FilePath $globalErrorPath -Encoding UTF8
    }
    catch {
        Write-SPOFactoryLog -Message "Failed to add error to registry: $_" -Level Warning -Category 'System'
    }
}

function Send-SPOFactoryCriticalAlert {
    <#
    .SYNOPSIS
        Sends critical alert for high-priority errors in MSP environment.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$ErrorReport
    )

    try {
        Write-SPOFactoryLog -Message "CRITICAL ERROR: $($ErrorReport.Exception.Message)" -Level Critical -ClientName $ErrorReport.ClientName -Category $ErrorReport.Category
        
        # Future implementation could include:
        # - Email notifications to MSP team
        # - SMS alerts for critical clients
        # - Integration with MSP ticketing system
        # - Webhook notifications to monitoring platforms
        # - SNMP traps to network management systems
        
        # For now, ensure critical errors are prominently logged
        Write-Host "CRITICAL ERROR DETECTED - Check logs for details: $($ErrorReport.ErrorId)" -ForegroundColor Red -BackgroundColor Yellow
    }
    catch {
        Write-SPOFactoryLog -Message "Failed to send critical alert: $_" -Level Warning -Category 'System'
    }
}

function Get-SPOFactoryErrorResolution {
    <#
    .SYNOPSIS
        Provides suggested resolution steps for common error classifications.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Classification
    )

    $resolutions = @{
        'Throttling' = @(
            'Reduce request frequency',
            'Implement exponential backoff',
            'Consider batching operations',
            'Review API usage patterns'
        )
        'Authorization' = @(
            'Verify application permissions',
            'Check user credentials',
            'Confirm tenant access rights',
            'Review authentication configuration'
        )
        'Network' = @(
            'Check internet connectivity',
            'Verify DNS resolution',
            'Test firewall/proxy settings',
            'Confirm SharePoint Online accessibility'
        )
        'TenantConfig' = @(
            'Review tenant subscription status',
            'Verify feature availability',
            'Check licensing requirements',
            'Confirm tenant configuration'
        )
        'QuotaExceeded' = @(
            'Review storage usage',
            'Consider upgrading plan',
            'Clean up unnecessary content',
            'Implement retention policies'
        )
        'ServiceUnavailable' = @(
            'Check Microsoft 365 service health',
            'Wait for service restoration',
            'Monitor service status',
            'Consider postponing operation'
        )
    }

    return $resolutions[$Classification] -join '; '
}

function Get-SPOFactoryErrorDocumentationUrl {
    <#
    .SYNOPSIS
        Returns documentation URL for error classification.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Classification
    )

    $baseUrl = "https://docs.microsoft.com/en-us/sharepoint/troubleshoot"
    
    $urls = @{
        'Throttling' = "$baseUrl/sharepoint-online-throttling"
        'Authorization' = "$baseUrl/access-denied-errors"
        'Network' = "$baseUrl/connectivity-issues"
        'TenantConfig' = "$baseUrl/administration"
        'QuotaExceeded' = "$baseUrl/storage-management"
        'ServiceUnavailable' = "https://status.office365.com/"
    }

    return $urls[$Classification] ?? $baseUrl
}

# Define custom exception class for SPOFactory
class SPOFactoryException : System.Exception {
    [string]$Classification
    [string]$ClientName
    [hashtable]$ErrorReport

    SPOFactoryException([string]$message, [System.Exception]$innerException, [string]$classification, [string]$clientName, [hashtable]$errorReport) : base($message, $innerException) {
        $this.Classification = $classification
        $this.ClientName = $clientName
        $this.ErrorReport = $errorReport
    }
}