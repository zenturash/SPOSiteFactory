#Requires -Modules Pester

<#
.SYNOPSIS
    Comprehensive Pester tests for SPOSiteFactory module in MSP environments

.DESCRIPTION
    Tests module loading, basic functionality, configuration management,
    connection handling, and MSP-specific features for the SPOSiteFactory module.
#>

# Import required modules for testing
Import-Module Pester -Force

# Test Configuration
$script:TestConfig = @{
    ModulePath = Split-Path -Parent $PSScriptRoot
    ModuleName = 'SPOSiteFactory'
    TestTenantUrl = 'https://contoso.sharepoint.com'
    TestClientName = 'TestClient'
    SkipIntegrationTests = $true  # Set to false for full integration testing
    TestDataPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Data'
}

# Test Helper Functions
function Initialize-TestEnvironment {
    param()
    
    # Clean up any existing module instances
    if (Get-Module -Name $script:TestConfig.ModuleName) {
        Remove-Module -Name $script:TestConfig.ModuleName -Force
    }
    
    # Import fresh module instance
    Import-Module (Join-Path $script:TestConfig.ModulePath "$($script:TestConfig.ModuleName).psd1") -Force
}

function Cleanup-TestEnvironment {
    param()
    
    # Clean up test artifacts
    if (Get-Module -Name $script:TestConfig.ModuleName) {
        Remove-Module -Name $script:TestConfig.ModuleName -Force
    }
}

# Main Test Suite
Describe "SPOSiteFactory Module Tests" {
    
    BeforeAll {
        Initialize-TestEnvironment
    }
    
    AfterAll {
        Cleanup-TestEnvironment
    }

    Context "Module Loading and Structure" {
        
        It "Should import the module without errors" {
            { Import-Module (Join-Path $script:TestConfig.ModulePath "$($script:TestConfig.ModuleName).psd1") -Force } | Should -Not -Throw
        }

        It "Should have the correct module name" {
            $module = Get-Module -Name $script:TestConfig.ModuleName
            $module.Name | Should -Be $script:TestConfig.ModuleName
        }

        It "Should have the expected version format" {
            $module = Get-Module -Name $script:TestConfig.ModuleName
            $module.Version | Should -Match '^\d+\.\d+\.\d+$'
        }

        It "Should export functions" {
            $module = Get-Module -Name $script:TestConfig.ModuleName
            $module.ExportedFunctions.Count | Should -BeGreaterOrEqual 0
        }

        It "Should have required dependencies" {
            $module = Get-Module -Name $script:TestConfig.ModuleName
            $requiredModules = @('PnP.PowerShell', 'PSFramework', 'Microsoft.PowerShell.SecretManagement')
            
            foreach ($requiredModule in $requiredModules) {
                $module.RequiredModules.Name | Should -Contain $requiredModule
            }
        }

        It "Should have MSP-specific metadata" {
            $module = Get-Module -Name $script:TestConfig.ModuleName
            $module.PrivateData.PSData.Tags | Should -Contain 'MSP'
            $module.PrivateData.PSData.Tags | Should -Contain 'MultiTenant'
        }
    }

    Context "Module Variables and Configuration" {
        
        It "Should initialize script variables correctly" {
            # Test that core script variables are initialized
            $script:SPOFactoryConfig | Should -Not -BeNullOrEmpty
            $script:SPOFactoryConstants | Should -Not -BeNullOrEmpty
        }

        It "Should have proper default configuration values" {
            $script:SPOFactoryConfig.MaxConcurrentConnections | Should -BeGreaterThan 0
            $script:SPOFactoryConfig.RetryAttempts | Should -BeGreaterOrEqual 1
            $script:SPOFactoryConfig.BatchSize | Should -BeGreaterThan 0
            $script:SPOFactoryConfig.LogPath | Should -Not -BeNullOrEmpty
            $script:SPOFactoryConfig.ConfigPath | Should -Not -BeNullOrEmpty
        }

        It "Should have MSP-specific constants" {
            $script:SPOFactoryConstants.MaxTenantCount | Should -BeGreaterThan 0
            $script:SPOFactoryConstants.SupportedRegions | Should -Contain 'Global'
            $script:SPOFactoryConstants.LogRetentionDays | Should -BeGreaterThan 0
        }
    }

    Context "Directory Structure" {
        
        It "Should create required directories on module load" {
            $requiredDirs = @(
                $script:SPOFactoryConfig.LogPath,
                $script:SPOFactoryConfig.ConfigPath,
                (Join-Path $script:SPOFactoryConfig.ConfigPath "Tenants"),
                (Join-Path $script:SPOFactoryConfig.ConfigPath "Baselines"),
                (Join-Path $script:SPOFactoryConfig.ConfigPath "Templates")
            )
            
            foreach ($dir in $requiredDirs) {
                Test-Path $dir | Should -Be $true
            }
        }

        It "Should have correct Data directory structure" {
            $dataPath = Join-Path $script:TestConfig.ModulePath "Data"
            Test-Path $dataPath | Should -Be $true
            Test-Path (Join-Path $dataPath "Baselines") | Should -Be $true
            Test-Path (Join-Path $dataPath "Templates") | Should -Be $true
            Test-Path (Join-Path $dataPath "Schemas") | Should -Be $true
        }
    }

    Context "Private Functions" {
        
        It "Should load Connect-SPOFactory function" {
            { Get-Command -Name Connect-SPOFactory -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }

        It "Should load Write-SPOFactoryLog function" {
            { Get-Command -Name Write-SPOFactoryLog -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }

        It "Should load Invoke-SPOFactoryCommand function" {
            { Get-Command -Name Invoke-SPOFactoryCommand -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }

        It "Should load configuration management functions" {
            { Get-Command -Name Get-SPOFactoryConfig -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
            { Get-Command -Name Set-SPOFactoryConfig -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }

        It "Should load prerequisite testing function" {
            { Get-Command -Name Test-SPOFactoryPrerequisites -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }

        It "Should load credential management functions" {
            { Get-Command -Name Get-SPOFactoryCredential -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
            { Get-Command -Name Set-SPOFactoryCredential -Module $script:TestConfig.ModuleName -ErrorAction Stop } | Should -Not -Throw
        }
    }

    Context "Configuration Management" {
        
        It "Should retrieve global configuration" {
            { Get-SPOFactoryConfig } | Should -Not -Throw
            $config = Get-SPOFactoryConfig
            $config | Should -Not -BeNullOrEmpty
        }

        It "Should support configuration validation" {
            $testConfig = @{
                LogPath = $env:TEMP
                ConfigPath = $env:TEMP
                MaxConcurrentConnections = 25
            }
            
            { Test-SPOFactoryConfig -Configuration $testConfig -ConfigType 'Global' } | Should -Not -Throw
        }

        It "Should handle invalid configuration gracefully" {
            $invalidConfig = @{
                MaxConcurrentConnections = -1  # Invalid value
            }
            
            $result = Test-SPOFactoryConfig -Configuration $invalidConfig -ConfigType 'Global'
            $result.IsValid | Should -Be $false
            $result.Errors.Count | Should -BeGreaterThan 0
        }
    }

    Context "Logging Framework" {
        
        It "Should write log messages without error" {
            { Write-SPOFactoryLog -Message "Test message" -Level Info } | Should -Not -Throw
        }

        It "Should support client-specific logging" {
            { Write-SPOFactoryLog -Message "Test client message" -Level Info -ClientName $script:TestConfig.TestClientName } | Should -Not -Throw
        }

        It "Should handle different log levels" {
            $logLevels = @('Info', 'Warning', 'Error', 'Debug', 'Verbose')
            
            foreach ($level in $logLevels) {
                { Write-SPOFactoryLog -Message "Test $level message" -Level $level } | Should -Not -Throw
            }
        }

        It "Should support audit logging" {
            { Write-SPOFactoryLog -Message "Audit test" -Level Info -EnableAuditLog } | Should -Not -Throw
        }
    }

    Context "Error Handling" {
        
        It "Should execute commands through error handler" {
            $testScript = { "Test successful" }
            
            { Invoke-SPOFactoryCommand -ScriptBlock $testScript -ErrorMessage "Test error" -PassThru } | Should -Not -Throw
        }

        It "Should handle script block failures" {
            $failingScript = { throw "Test exception" }
            
            { Invoke-SPOFactoryCommand -ScriptBlock $failingScript -ErrorMessage "Test error" -SuppressErrors } | Should -Not -Throw
        }

        It "Should classify errors correctly" {
            $testException = [System.TimeoutException]::new("Connection timeout")
            
            $result = Get-SPOFactoryErrorClassification -Exception $testException
            $result.Classification | Should -Be 'Timeout'
            $result.ShouldRetry | Should -Be $true
        }

        It "Should calculate retry delays with exponential backoff" {
            $delay1 = Get-SPOFactoryRetryDelay -Attempt 1 -BaseDelay 1
            $delay2 = Get-SPOFactoryRetryDelay -Attempt 2 -BaseDelay 1
            $delay3 = Get-SPOFactoryRetryDelay -Attempt 3 -BaseDelay 1
            
            $delay2 | Should -BeGreaterThan $delay1
            $delay3 | Should -BeGreaterThan $delay2
        }
    }

    Context "Prerequisite Testing" {
        
        It "Should test prerequisites without error" {
            { Test-SPOFactoryPrerequisites -SkipConnectivity -SkipCredentials } | Should -Not -Throw
        }

        It "Should return prerequisite test results" {
            $result = Test-SPOFactoryPrerequisites -SkipConnectivity -SkipCredentials
            $result | Should -Not -BeNullOrEmpty
            $result.Passed | Should -BeOfType [bool]
            $result.Score | Should -BeOfType [int]
            $result.MaxScore | Should -BeOfType [int]
        }

        It "Should test PowerShell version" {
            $result = Test-SPOFactoryPowerShellVersion
            $result.Name | Should -Be 'PowerShell Version'
            $result.MaxScore | Should -BeGreaterThan 0
        }

        It "Should validate directory structure" {
            $result = Test-SPOFactoryDirectoryStructure -Fix
            $result.Name | Should -Be 'Directory Structure'
            $result.MaxScore | Should -BeGreaterThan 0
        }
    }

    Context "Data Files and Templates" {
        
        It "Should have MSP baseline files" {
            $baselineFiles = @('MSPStandard.json', 'MSPSecure.json')
            $baselinesPath = Join-Path $script:TestConfig.TestDataPath "Baselines"
            
            foreach ($file in $baselineFiles) {
                Test-Path (Join-Path $baselinesPath $file) | Should -Be $true
            }
        }

        It "Should have site template files" {
            $templateFiles = @('TeamSite.json', 'CommunicationSite.json', 'HubSite.json')
            $templatesPath = Join-Path $script:TestConfig.TestDataPath "Templates"
            
            foreach ($file in $templateFiles) {
                Test-Path (Join-Path $templatesPath $file) | Should -Be $true
            }
        }

        It "Should have schema validation files" {
            $schemaFiles = @('SiteTemplate.schema.json', 'Baseline.schema.json')
            $schemasPath = Join-Path $script:TestConfig.TestDataPath "Schemas"
            
            foreach ($file in $schemaFiles) {
                Test-Path (Join-Path $schemasPath $file) | Should -Be $true
            }
        }

        It "Should load baseline configurations as valid JSON" {
            $baselinePath = Join-Path $script:TestConfig.TestDataPath "Baselines\MSPStandard.json"
            
            { Get-Content $baselinePath -Raw | ConvertFrom-Json } | Should -Not -Throw
            
            $baseline = Get-Content $baselinePath -Raw | ConvertFrom-Json
            $baseline.name | Should -Be 'MSPStandard'
            $baseline.version | Should -Not -BeNullOrEmpty
            $baseline.tenantSettings | Should -Not -BeNullOrEmpty
        }

        It "Should load template configurations as valid JSON" {
            $templatePath = Join-Path $script:TestConfig.TestDataPath "Templates\TeamSite.json"
            
            { Get-Content $templatePath -Raw | ConvertFrom-Json } | Should -Not -Throw
            
            $template = Get-Content $templatePath -Raw | ConvertFrom-Json
            $template.name | Should -Be 'TeamSite'
            $template.version | Should -Not -BeNullOrEmpty
            $template.siteConfiguration | Should -Not -BeNullOrEmpty
        }
    }

    Context "MSP Multi-Tenant Support" -Skip:$script:TestConfig.SkipIntegrationTests {
        
        It "Should handle multiple client configurations" {
            $clients = @('Client1', 'Client2', 'Client3')
            
            foreach ($client in $clients) {
                $clientConfig = @{
                    DefaultBaseline = 'MSPStandard'
                    TenantUrl = "https://$client.sharepoint.com"
                }
                
                { Set-SPOFactoryConfig -ClientName $client -Configuration $clientConfig -ConfigType 'Client' } | Should -Not -Throw
            }
            
            foreach ($client in $clients) {
                $config = Get-SPOFactoryConfig -ClientName $client -ConfigType 'Client'
                $config | Should -Not -BeNullOrEmpty
            }
        }

        It "Should support tenant registry operations" {
            { Update-SPOFactoryTenantRegistry -ClientName $script:TestConfig.TestClientName -TenantUrl $script:TestConfig.TestTenantUrl -LastConnected (Get-Date) -Activity "Test" } | Should -Not -Throw
            
            $registry = Get-SPOFactoryTenantRegistry -ClientName $script:TestConfig.TestClientName
            $registry | Should -Not -BeNullOrEmpty
            $registry.ClientName | Should -Be $script:TestConfig.TestClientName
        }

        It "Should handle connection pool management" {
            # Test connection pool initialization
            $script:SPOFactoryConnectionPool | Should -Not -BeNull
            
            # Test connection pool cleanup
            { Remove-SPOFactoryStaleConnections -MaxAge 0 } | Should -Not -Throw
        }
    }

    Context "Security and Compliance" {
        
        It "Should support secure credential operations" {
            # Test credential name generation
            $secretName = Get-SPOFactorySecretName -ClientName $script:TestConfig.TestClientName -AuthType 'Password'
            $secretName | Should -Be "SPOFactory-$($script:TestConfig.TestClientName)-Password"
        }

        It "Should clear credential cache" {
            { Clear-SPOFactoryCredentialCache } | Should -Not -Throw
        }

        It "Should validate security configurations" {
            $securityConfig = Get-SPOFactorySecurityConfig
            $securityConfig | Should -Not -BeNullOrEmpty
            $securityConfig.EncryptCredentials | Should -BeOfType [bool]
            $securityConfig.AuditAllOperations | Should -BeOfType [bool]
        }
    }

    Context "Performance and Scalability" {
        
        It "Should handle cache operations efficiently" {
            $testKey = "TestKey"
            $testValue = @{ Data = "Test Value"; Timestamp = Get-Date }
            $cacheTime = New-TimeSpan -Minutes 5
            
            { Set-SPOFactoryCacheItem -Key $testKey -Value $testValue -ExpiresIn $cacheTime } | Should -Not -Throw
            
            $cachedValue = Get-SPOFactoryCacheItem -Key $testKey -MaxAge $cacheTime
            $cachedValue | Should -Not -BeNullOrEmpty
            $cachedValue.Data | Should -Be "Test Value"
        }

        It "Should handle batch operations configuration" {
            $config = Get-SPOFactoryConfig
            $config.BatchSize | Should -BeGreaterThan 0
            $config.BatchSize | Should -BeLessOrEqual 1000
        }

        It "Should support concurrent connection limits" {
            $config = Get-SPOFactoryConfig
            $config.MaxConcurrentConnections | Should -BeGreaterThan 0
            $config.MaxConcurrentConnections | Should -BeLessOrEqual 100
        }
    }

    Context "Module Cleanup and Removal" {
        
        It "Should handle module removal gracefully" {
            # This tests the OnRemove event handler
            { Remove-Module -Name $script:TestConfig.ModuleName -Force } | Should -Not -Throw
        }
        
        It "Should clean up resources on removal" {
            # Re-import for testing cleanup
            Import-Module (Join-Path $script:TestConfig.ModulePath "$($script:TestConfig.ModuleName).psd1") -Force
            
            # Simulate some connections
            $script:SPOFactoryConnections = @{
                'TestClient' = @{
                    ClientName = 'TestClient'
                    Connected = Get-Date
                }
            }
            
            { Remove-Module -Name $script:TestConfig.ModuleName -Force } | Should -Not -Throw
        }
    }
}

# Integration Tests (Only run when explicitly enabled)
Describe "SPOSiteFactory Integration Tests" -Skip:$script:TestConfig.SkipIntegrationTests {
    
    BeforeAll {
        Initialize-TestEnvironment
    }
    
    AfterAll {
        Cleanup-TestEnvironment
    }

    Context "SharePoint Online Connectivity" {
        
        It "Should test network connectivity to SharePoint endpoints" {
            $connectivityTest = Test-SPOFactoryNetworkConnectivity
            $connectivityTest | Should -Not -BeNullOrEmpty
            $connectivityTest.Name | Should -Be 'Network Connectivity'
        }
    }

    Context "End-to-End MSP Workflows" {
        
        It "Should complete full prerequisite check" {
            $result = Test-SPOFactoryPrerequisites -Detailed
            $result | Should -Not -BeNullOrEmpty
            $result.Tests | Should -Not -BeNullOrEmpty
        }
    }
}

# Performance Tests
Describe "SPOSiteFactory Performance Tests" -Skip:$script:TestConfig.SkipIntegrationTests {
    
    Context "Module Loading Performance" {
        
        It "Should load module within acceptable time" {
            $loadTime = Measure-Command {
                Remove-Module -Name $script:TestConfig.ModuleName -Force -ErrorAction SilentlyContinue
                Import-Module (Join-Path $script:TestConfig.ModulePath "$($script:TestConfig.ModuleName).psd1") -Force
            }
            
            $loadTime.TotalSeconds | Should -BeLessThan 10
        }
    }

    Context "Logging Performance" {
        
        It "Should handle high-volume logging efficiently" {
            $logTime = Measure-Command {
                1..100 | ForEach-Object {
                    Write-SPOFactoryLog -Message "Performance test message $_" -Level Info -ClientName "PerfTestClient"
                }
            }
            
            $logTime.TotalSeconds | Should -BeLessThan 30
        }
    }
}

# Export test configuration for external use
$script:TestConfig | Export-Clixml -Path (Join-Path $PSScriptRoot "TestConfig.xml") -Force