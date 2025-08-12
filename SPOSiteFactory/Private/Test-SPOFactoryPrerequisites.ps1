function Test-SPOFactoryPrerequisites {
    <#
    .SYNOPSIS
        Tests and validates all prerequisites for SPOSiteFactory module in MSP environments.

    .DESCRIPTION
        Comprehensive prerequisite validation for MSP SharePoint Online operations including
        PowerShell version, required modules, network connectivity, credentials, and
        tenant-specific requirements.

    .PARAMETER ClientName
        Specific client name to test prerequisites for

    .PARAMETER SkipConnectivity
        Skip network connectivity tests

    .PARAMETER SkipCredentials
        Skip credential validation tests

    .PARAMETER SkipModules
        Skip module availability tests

    .PARAMETER Detailed
        Return detailed test results for each check

    .PARAMETER Fix
        Attempt to fix issues automatically where possible

    .EXAMPLE
        Test-SPOFactoryPrerequisites

    .EXAMPLE
        Test-SPOFactoryPrerequisites -ClientName "Contoso Corp" -Detailed

    .EXAMPLE
        Test-SPOFactoryPrerequisites -Fix
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipConnectivity,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipCredentials,
        
        [Parameter(Mandatory = $false)]
        [switch]$SkipModules,
        
        [Parameter(Mandatory = $false)]
        [switch]$Detailed,
        
        [Parameter(Mandatory = $false)]
        [switch]$Fix
    )

    begin {
        Write-SPOFactoryLog -Message "Starting prerequisite validation" -Level Info -ClientName $ClientName -Category 'System'
        
        $results = @{
            Overall = @{
                Passed = $true
                Score = 0
                MaxScore = 0
                Issues = @()
                Warnings = @()
            }
            Tests = @{}
        }
    }

    process {
        try {
            # Test PowerShell Version
            $results.Tests['PowerShellVersion'] = Test-SPOFactoryPowerShellVersion -Fix:$Fix
            Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'PowerShellVersion'

            # Test Required Modules
            if (-not $SkipModules) {
                $results.Tests['RequiredModules'] = Test-SPOFactoryRequiredModules -Fix:$Fix
                Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'RequiredModules'
            }

            # Test Execution Policy
            $results.Tests['ExecutionPolicy'] = Test-SPOFactoryExecutionPolicy -Fix:$Fix
            Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'ExecutionPolicy'

            # Test Directory Structure
            $results.Tests['DirectoryStructure'] = Test-SPOFactoryDirectoryStructure -Fix:$Fix
            Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'DirectoryStructure'

            # Test Secret Management
            if (-not $SkipCredentials) {
                $results.Tests['SecretManagement'] = Test-SPOFactorySecretManagement -Fix:$Fix
                Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'SecretManagement'
            }

            # Test Network Connectivity
            if (-not $SkipConnectivity) {
                $results.Tests['NetworkConnectivity'] = Test-SPOFactoryNetworkConnectivity
                Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'NetworkConnectivity'
            }

            # Test Client-Specific Prerequisites
            if ($ClientName) {
                $results.Tests['ClientPrerequisites'] = Test-SPOFactoryClientPrerequisites -ClientName $ClientName -Fix:$Fix
                Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'ClientPrerequisites'
            }

            # Test Security Configuration
            $results.Tests['SecurityConfig'] = Test-SPOFactorySecurityConfiguration -Fix:$Fix
            Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'SecurityConfig'

            # Test Performance Settings
            $results.Tests['PerformanceSettings'] = Test-SPOFactoryPerformanceSettings -Fix:$Fix
            Update-SPOFactoryPrerequisiteScore -Results $results -TestName 'PerformanceSettings'

            # Generate overall assessment
            $results.Overall.Score = ($results.Tests.Values | Measure-Object -Property Score -Sum).Sum
            $results.Overall.MaxScore = ($results.Tests.Values | Measure-Object -Property MaxScore -Sum).Sum
            $results.Overall.Passed = $results.Overall.Score -eq $results.Overall.MaxScore

            # Collect all issues
            $results.Overall.Issues = $results.Tests.Values | ForEach-Object { $_.Issues } | Where-Object { $_ }
            $results.Overall.Warnings = $results.Tests.Values | ForEach-Object { $_.Warnings } | Where-Object { $_ }

            # Log results
            $overallPercentage = if ($results.Overall.MaxScore -gt 0) { 
                [math]::Round(($results.Overall.Score / $results.Overall.MaxScore) * 100, 1) 
            } else { 
                0 
            }

            $logLevel = if ($results.Overall.Passed) { 'Info' } elseif ($overallPercentage -ge 80) { 'Warning' } else { 'Error' }
            Write-SPOFactoryLog -Message "Prerequisite validation completed: $($results.Overall.Score)/$($results.Overall.MaxScore) ($overallPercentage%)" -Level $logLevel -ClientName $ClientName -Category 'System'

            if ($results.Overall.Issues.Count -gt 0) {
                Write-SPOFactoryLog -Message "Issues found: $($results.Overall.Issues -join '; ')" -Level Warning -ClientName $ClientName -Category 'System'
            }

            if ($Detailed) {
                return $results
            } else {
                return @{
                    Passed = $results.Overall.Passed
                    Score = $results.Overall.Score
                    MaxScore = $results.Overall.MaxScore
                    Percentage = $overallPercentage
                    Issues = $results.Overall.Issues
                    Warnings = $results.Overall.Warnings
                }
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Prerequisite validation failed: $_" -Level Error -ClientName $ClientName -Category 'System' -Exception $_.Exception
            throw
        }
    }
}

function Test-SPOFactoryPowerShellVersion {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'PowerShell Version'
        Passed = $false
        Score = 0
        MaxScore = 10
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        $psVersion = $PSVersionTable.PSVersion
        $result.Details.CurrentVersion = $psVersion.ToString()
        $result.Details.RequiredVersion = '5.1'

        if ($psVersion.Major -ge 5 -and ($psVersion.Major -gt 5 -or $psVersion.Minor -ge 1)) {
            $result.Passed = $true
            $result.Score = 10
            
            if ($psVersion.Major -eq 5) {
                $result.Warnings += "PowerShell 5.1 detected. Consider upgrading to PowerShell 7+ for better performance"
            }
        } else {
            $result.Issues += "PowerShell version $($psVersion) is not supported. Minimum required: 5.1"
            
            if ($Fix) {
                $result.Issues += "Automatic PowerShell upgrade not supported. Please upgrade manually"
            }
        }
    }
    catch {
        $result.Issues += "Failed to check PowerShell version: $_"
    }

    return $result
}

function Test-SPOFactoryRequiredModules {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Required Modules'
        Passed = $false
        Score = 0
        MaxScore = 30
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    $requiredModules = @{
        'PnP.PowerShell' = '2.0.0'
        'PSFramework' = '1.7.0'
        'Microsoft.PowerShell.SecretManagement' = '1.1.2'
    }

    $moduleResults = @{}
    $totalScore = 0

    foreach ($moduleName in $requiredModules.Keys) {
        $requiredVersion = $requiredModules[$moduleName]
        $moduleInfo = @{
            Required = $requiredVersion
            Installed = $null
            Available = $false
            VersionMatch = $false
        }

        try {
            $installedModule = Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            
            if ($installedModule) {
                $moduleInfo.Installed = $installedModule.Version.ToString()
                $moduleInfo.Available = $true
                
                if ($installedModule.Version -ge [Version]$requiredVersion) {
                    $moduleInfo.VersionMatch = $true
                    $totalScore += 10
                } else {
                    $result.Issues += "$moduleName version $($installedModule.Version) is below required version $requiredVersion"
                    
                    if ($Fix) {
                        try {
                            Update-Module -Name $moduleName -RequiredVersion $requiredVersion -Force
                            $result.Warnings += "Updated $moduleName to version $requiredVersion"
                            $totalScore += 10
                            $moduleInfo.VersionMatch = $true
                        }
                        catch {
                            $result.Issues += "Failed to update $moduleName`: $_"
                        }
                    }
                }
            } else {
                $result.Issues += "$moduleName is not installed"
                
                if ($Fix) {
                    try {
                        Install-Module -Name $moduleName -RequiredVersion $requiredVersion -Scope CurrentUser -Force -AllowClobber
                        $result.Warnings += "Installed $moduleName version $requiredVersion"
                        $totalScore += 10
                        $moduleInfo.Available = $true
                        $moduleInfo.VersionMatch = $true
                        $moduleInfo.Installed = $requiredVersion
                    }
                    catch {
                        $result.Issues += "Failed to install $moduleName`: $_"
                    }
                }
            }
        }
        catch {
            $result.Issues += "Error checking $moduleName`: $_"
        }

        $moduleResults[$moduleName] = $moduleInfo
    }

    $result.Details.Modules = $moduleResults
    $result.Score = $totalScore
    $result.Passed = $totalScore -eq 30

    return $result
}

function Test-SPOFactoryExecutionPolicy {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Execution Policy'
        Passed = $false
        Score = 0
        MaxScore = 5
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
        $result.Details.CurrentPolicy = $currentPolicy
        $result.Details.RequiredPolicies = @('RemoteSigned', 'Unrestricted', 'Bypass')

        if ($currentPolicy -in @('RemoteSigned', 'Unrestricted', 'Bypass')) {
            $result.Passed = $true
            $result.Score = 5
        } else {
            $result.Issues += "Execution policy '$currentPolicy' may prevent module loading"
            
            if ($Fix) {
                try {
                    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
                    $result.Warnings += "Changed execution policy to RemoteSigned for current user"
                    $result.Passed = $true
                    $result.Score = 5
                }
                catch {
                    $result.Issues += "Failed to set execution policy: $_"
                }
            }
        }
    }
    catch {
        $result.Issues += "Failed to check execution policy: $_"
    }

    return $result
}

function Test-SPOFactoryDirectoryStructure {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Directory Structure'
        Passed = $false
        Score = 0
        MaxScore = 10
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    $requiredDirs = @(
        $script:SPOFactoryConfig.LogPath,
        $script:SPOFactoryConfig.ConfigPath,
        (Join-Path $script:SPOFactoryConfig.ConfigPath "Tenants"),
        (Join-Path $script:SPOFactoryConfig.ConfigPath "Baselines"),
        (Join-Path $script:SPOFactoryConfig.ConfigPath "Templates"),
        (Join-Path $script:SPOFactoryConfig.LogPath "Clients"),
        (Join-Path $script:SPOFactoryConfig.LogPath "Audit"),
        (Join-Path $script:SPOFactoryConfig.LogPath "Performance")
    )

    $createdDirs = 0
    $dirResults = @{}

    foreach ($dir in $requiredDirs) {
        $dirInfo = @{
            Path = $dir
            Exists = Test-Path $dir
            Created = $false
            Error = $null
        }

        if ($dirInfo.Exists) {
            $createdDirs++
        } elseif ($Fix) {
            try {
                New-Item -Path $dir -ItemType Directory -Force | Out-Null
                $dirInfo.Created = $true
                $dirInfo.Exists = $true
                $createdDirs++
                $result.Warnings += "Created directory: $dir"
            }
            catch {
                $dirInfo.Error = $_.Exception.Message
                $result.Issues += "Failed to create directory $dir`: $_"
            }
        } else {
            $result.Issues += "Required directory does not exist: $dir"
        }

        $dirResults[$dir] = $dirInfo
    }

    $result.Details.Directories = $dirResults
    $result.Score = [math]::Round(($createdDirs / $requiredDirs.Count) * 10)
    $result.Passed = $createdDirs -eq $requiredDirs.Count

    return $result
}

function Test-SPOFactorySecretManagement {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Secret Management'
        Passed = $false
        Score = 0
        MaxScore = 15
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        # Check if SecretManagement module is available
        $secretMgmtModule = Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable
        if (-not $secretMgmtModule) {
            $result.Issues += "Microsoft.PowerShell.SecretManagement module not found"
            return $result
        }

        $result.Score += 5
        $result.Details.ModuleAvailable = $true

        # Check for secret vault
        $vaultName = $script:SPOFactoryConfig.CredentialVault
        $vault = Get-SecretVault -Name $vaultName -ErrorAction SilentlyContinue
        
        if ($vault) {
            $result.Score += 10
            $result.Details.VaultExists = $true
            $result.Details.VaultName = $vaultName
            $result.Passed = $true
        } else {
            $result.Issues += "Secret vault '$vaultName' not found"
            $result.Details.VaultExists = $false
            
            if ($Fix) {
                # Note: This is a placeholder - actual vault creation requires specific provider
                $result.Warnings += "Secret vault creation requires manual setup with specific provider"
            }
        }
    }
    catch {
        $result.Issues += "Failed to check secret management: $_"
    }

    return $result
}

function Test-SPOFactoryNetworkConnectivity {
    [CmdletBinding()]
    param()

    $result = @{
        Name = 'Network Connectivity'
        Passed = $false
        Score = 0
        MaxScore = 15
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    $endpoints = @{
        'SharePoint Online' = 'https://graph.microsoft.com'
        'Microsoft Graph' = 'https://graph.microsoft.com/v1.0'
        'Office 365' = 'https://login.microsoftonline.com'
        'Azure AD' = 'https://login.microsoftonline.com/common/oauth2/authorize'
    }

    $successfulConnections = 0
    $connectionResults = @{}

    foreach ($endpointName in $endpoints.Keys) {
        $url = $endpoints[$endpointName]
        $connectionInfo = @{
            Url = $url
            Accessible = $false
            ResponseTime = $null
            Error = $null
        }

        try {
            $startTime = Get-Date
            $response = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 10 -UseBasicParsing
            $connectionInfo.ResponseTime = ((Get-Date) - $startTime).TotalMilliseconds
            $connectionInfo.Accessible = $response.StatusCode -eq 200
            
            if ($connectionInfo.Accessible) {
                $successfulConnections++
            }
        }
        catch {
            $connectionInfo.Error = $_.Exception.Message
            $result.Issues += "Cannot reach $endpointName ($url): $_"
        }

        $connectionResults[$endpointName] = $connectionInfo
    }

    $result.Details.Connections = $connectionResults
    $result.Score = [math]::Round(($successfulConnections / $endpoints.Count) * 15)
    $result.Passed = $successfulConnections -eq $endpoints.Count

    if ($result.Passed) {
        $avgResponseTime = ($connectionResults.Values | Where-Object { $_.ResponseTime } | Measure-Object -Property ResponseTime -Average).Average
        if ($avgResponseTime -gt 5000) {
            $result.Warnings += "High network latency detected (avg: $([math]::Round($avgResponseTime))ms)"
        }
    }

    return $result
}

function Test-SPOFactoryClientPrerequisites {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        [switch]$Fix
    )

    $result = @{
        Name = 'Client Prerequisites'
        Passed = $false
        Score = 0
        MaxScore = 20
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        # Check client configuration
        $clientConfig = Get-SPOFactoryClientConfig -ClientName $ClientName
        if ($clientConfig.Count -gt 0) {
            $result.Score += 5
            $result.Details.ConfigExists = $true
        } else {
            $result.Issues += "No configuration found for client: $ClientName"
            $result.Details.ConfigExists = $false
        }

        # Check client credentials
        $credential = Get-SPOFactoryCredential -ClientName $ClientName
        if ($credential) {
            $result.Score += 10
            $result.Details.CredentialsExist = $true
        } else {
            $result.Issues += "No stored credentials found for client: $ClientName"
            $result.Details.CredentialsExist = $false
        }

        # Check client-specific directories
        $clientLogDir = Join-Path $script:SPOFactoryConfig.LogPath "Clients\$ClientName"
        if (Test-Path $clientLogDir) {
            $result.Score += 5
            $result.Details.LogDirectoryExists = $true
        } elseif ($Fix) {
            try {
                New-Item -Path $clientLogDir -ItemType Directory -Force | Out-Null
                $result.Score += 5
                $result.Details.LogDirectoryExists = $true
                $result.Warnings += "Created client log directory: $clientLogDir"
            }
            catch {
                $result.Issues += "Failed to create client log directory: $_"
            }
        } else {
            $result.Issues += "Client log directory does not exist: $clientLogDir"
        }

        $result.Passed = $result.Score -eq $result.MaxScore
    }
    catch {
        $result.Issues += "Failed to check client prerequisites: $_"
    }

    return $result
}

function Test-SPOFactorySecurityConfiguration {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Security Configuration'
        Passed = $false
        Score = 0
        MaxScore = 10
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        $securityConfig = Get-SPOFactorySecurityConfig
        $result.Details.SecurityConfig = $securityConfig

        # Check encryption settings
        if ($securityConfig.EncryptCredentials) {
            $result.Score += 3
        } else {
            $result.Issues += "Credential encryption is not enabled"
        }

        # Check audit settings
        if ($securityConfig.AuditAllOperations) {
            $result.Score += 3
        } else {
            $result.Warnings += "Audit logging is not enabled"
        }

        # Check secure connection requirement
        if ($securityConfig.RequireSecureConnection) {
            $result.Score += 2
        } else {
            $result.Issues += "Secure connection requirement is not enabled"
        }

        # Check certificate validation
        if ($securityConfig.CertificateValidation) {
            $result.Score += 2
        } else {
            $result.Warnings += "Certificate validation is not enabled"
        }

        $result.Passed = $result.Score -eq $result.MaxScore
    }
    catch {
        $result.Issues += "Failed to check security configuration: $_"
    }

    return $result
}

function Test-SPOFactoryPerformanceSettings {
    [CmdletBinding()]
    param([switch]$Fix)

    $result = @{
        Name = 'Performance Settings'
        Passed = $false
        Score = 0
        MaxScore = 10
        Issues = @()
        Warnings = @()
        Details = @{}
    }

    try {
        $config = Get-SPOFactoryGlobalConfig
        
        # Check concurrent connections
        $maxConnections = $config.MaxConcurrentConnections
        if ($maxConnections -gt 0 -and $maxConnections -le 100) {
            $result.Score += 3
            if ($maxConnections -gt 50) {
                $result.Warnings += "High concurrent connection limit may impact performance"
            }
        } else {
            $result.Issues += "Invalid MaxConcurrentConnections setting: $maxConnections"
        }

        # Check batch size
        $batchSize = $config.BatchSize
        if ($batchSize -gt 0 -and $batchSize -le 1000) {
            $result.Score += 2
            if ($batchSize -gt 500) {
                $result.Warnings += "Large batch size may cause timeouts"
            }
        } else {
            $result.Issues += "Invalid BatchSize setting: $batchSize"
        }

        # Check timeout settings
        $timeout = $config.ConnectionTimeout
        if ($timeout -ge 30 -and $timeout -le 600) {
            $result.Score += 2
        } else {
            $result.Issues += "Invalid ConnectionTimeout setting: $timeout"
        }

        # Check retry settings
        $retries = $config.RetryAttempts
        if ($retries -ge 1 -and $retries -le 10) {
            $result.Score += 3
        } else {
            $result.Issues += "Invalid RetryAttempts setting: $retries"
        }

        $result.Details.MaxConnections = $maxConnections
        $result.Details.BatchSize = $batchSize
        $result.Details.ConnectionTimeout = $timeout
        $result.Details.RetryAttempts = $retries

        $result.Passed = $result.Score -eq $result.MaxScore
    }
    catch {
        $result.Issues += "Failed to check performance settings: $_"
    }

    return $result
}

function Update-SPOFactoryPrerequisiteScore {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Results,
        
        [Parameter(Mandatory = $true)]
        [string]$TestName
    )

    $testResult = $Results.Tests[$TestName]
    if (-not $testResult.Passed -and $testResult.Issues.Count -gt 0) {
        $Results.Overall.Issues += $testResult.Issues
    }
    
    if ($testResult.Warnings.Count -gt 0) {
        $Results.Overall.Warnings += $testResult.Warnings
    }
}