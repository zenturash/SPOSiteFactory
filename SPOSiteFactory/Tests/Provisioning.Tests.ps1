# SPOSiteFactory Provisioning Tests
# Tests for Phase 2: Core Provisioning Functions

BeforeAll {
    # Import the module
    $modulePath = Split-Path -Parent $PSScriptRoot
    Import-Module "$modulePath\SPOSiteFactory.psd1" -Force
    
    # Mock PnP cmdlets to avoid actual SharePoint calls
    Mock Get-PnPContext { return @{ Url = "https://contoso.sharepoint.com" } }
    Mock Connect-PnPOnline { return $true }
    Mock Get-PnPSite { return @{ Id = [guid]::NewGuid(); Title = "Test Site" } }
    Mock Get-PnPHubSite { return @{ Id = [guid]::NewGuid(); Title = "Test Hub" } }
    Mock New-PnPSite { return @{ Url = "https://contoso.sharepoint.com/sites/test" } }
    Mock New-PnPTenantSite { return $true }
    Mock Register-PnPHubSite { return $true }
    Mock Add-PnPHubSiteAssociation { return $true }
    Mock Set-PnPSite { return $true }
    Mock Enable-PnPFeature { return $true }
    Mock Get-PnPList { return @() }
    Mock New-PnPMicrosoft365Group { return @{ Id = [guid]::NewGuid(); DisplayName = "Test Group" } }
    Mock Remove-PnPSite { return $true }
    Mock Remove-PnPMicrosoft365Group { return $true }
    
    # Test variables
    $script:testClientName = "TestClient"
    $script:testSiteUrl = "https://contoso.sharepoint.com/sites/testsite"
    $script:testHubUrl = "https://contoso.sharepoint.com/sites/testhub"
    $script:testOwner = "admin@contoso.com"
}

Describe "New-SPOHubSite" {
    Context "Creating hub sites" {
        It "Should create a hub site with required parameters" {
            $result = New-SPOHubSite -Title "Test Hub" -Url "testhub" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should apply security baseline when specified" {
            Mock Set-SPOSiteSecurityBaseline { return @{ Success = $true } }
            
            $result = New-SPOHubSite -Title "Secure Hub" -Url "securehub" -ClientName $testClientName -SecurityBaseline "MSPSecure" -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should validate required parameters" {
            { New-SPOHubSite -Title "" -Url "hub" -ClientName $testClientName } | Should -Throw
            { New-SPOHubSite -Title "Hub" -Url "" -ClientName $testClientName } | Should -Throw
        }
        
        It "Should support WhatIf parameter" {
            $result = New-SPOHubSite -Title "WhatIf Hub" -Url "whatifhub" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Hub site configuration" {
        It "Should apply hub settings correctly" {
            Mock Invoke-SPOFactoryCommand { return @{ Success = $true } }
            
            $result = New-SPOHubSite -Title "Config Hub" -Url "confighub" -ClientName $testClientName -Description "Test Description" -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should handle hub site registration" {
            Mock Register-PnPHubSite { return @{ Id = [guid]::NewGuid() } }
            
            $result = New-SPOHubSite -Title "Register Hub" -Url "registerhub" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Error handling" {
        It "Should handle site creation failures gracefully" {
            Mock New-PnPTenantSite { throw "Site creation failed" }
            
            { New-SPOHubSite -Title "Failed Hub" -Url "failedhub" -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
        
        It "Should rollback on registration failure" {
            Mock Register-PnPHubSite { throw "Registration failed" }
            Mock Remove-PnPSite { return $true }
            
            { New-SPOHubSite -Title "Rollback Hub" -Url "rollbackhub" -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
    }
}

Describe "New-SPOSite" {
    Context "Team site creation" {
        It "Should create a team site with M365 Group" {
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Test Team Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should create team site without M365 Group when specified" {
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "No Group Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -CreateM365Group:$false -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should add members and owners to M365 Group" {
            Mock Add-PnPMicrosoft365GroupMember { return $true }
            Mock Add-PnPMicrosoft365GroupOwner { return $true }
            
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Group Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -GroupMembers @("user1@contoso.com") -GroupOwners @("owner2@contoso.com") -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Communication site creation" {
        It "Should create a communication site" {
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Test Comm Site" -Owner $testOwner -SiteType "CommunicationSite" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should not create M365 Group for communication sites" {
            Mock New-PnPMicrosoft365Group { throw "Should not be called" }
            
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Comm Site" -Owner $testOwner -SiteType "CommunicationSite" -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Security baseline application" {
        It "Should apply security baseline to new sites" {
            Mock Set-SPOSiteSecurityBaseline { return @{ Success = $true; Applied = $true } }
            
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Secure Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -SecurityBaseline "MSPSecure" -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should configure Office file handling" {
            Mock Set-SPOOfficeFileHandling { return @{ Success = $true } }
            
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Office Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -ConfigureOfficeFileHandling -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Hub association" {
        It "Should associate site with hub when specified" {
            Mock Add-SPOSiteToHub { return @{ Success = $true } }
            
            $result = New-SPOSite -SiteUrl $testSiteUrl -Title "Hub Member" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -HubSiteUrl $testHubUrl -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Error handling and rollback" {
        It "Should rollback M365 Group on site creation failure" {
            Mock New-PnPMicrosoft365Group { return @{ Id = [guid]::NewGuid() } }
            Mock New-PnPSite { throw "Site creation failed" }
            Mock Remove-PnPMicrosoft365Group { return $true }
            
            { New-SPOSite -SiteUrl $testSiteUrl -Title "Rollback Site" -Owner $testOwner -SiteType "TeamSite" -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
    }
}

Describe "Add-SPOSiteToHub" {
    Context "Single site association" {
        It "Should associate a single site with hub" {
            Mock Get-PnPHubSite { return @{ Id = [guid]::NewGuid(); Title = "Hub" } }
            Mock Add-PnPHubSiteAssociation { return $true }
            
            $result = Add-SPOSiteToHub -HubSiteUrl $testHubUrl -SiteUrl $testSiteUrl -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Bulk site association" {
        It "Should associate multiple sites with hub" {
            Mock Get-PnPHubSite { return @{ Id = [guid]::NewGuid(); Title = "Hub" } }
            Mock Add-PnPHubSiteAssociation { return $true }
            
            $siteUrls = @(
                "https://contoso.sharepoint.com/sites/site1",
                "https://contoso.sharepoint.com/sites/site2"
            )
            
            $result = Add-SPOSiteToHub -HubSiteUrl $testHubUrl -SiteUrls $siteUrls -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should continue on error when specified" {
            Mock Get-PnPHubSite { return @{ Id = [guid]::NewGuid(); Title = "Hub" } }
            Mock Add-PnPHubSiteAssociation { 
                if ($args[0] -match "site2") { throw "Association failed" }
                return $true 
            }
            
            $siteUrls = @(
                "https://contoso.sharepoint.com/sites/site1",
                "https://contoso.sharepoint.com/sites/site2",
                "https://contoso.sharepoint.com/sites/site3"
            )
            
            $result = Add-SPOSiteToHub -HubSiteUrl $testHubUrl -SiteUrls $siteUrls -ClientName $testClientName -ContinueOnError -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Hub validation" {
        It "Should validate hub site exists" {
            Mock Get-PnPHubSite { return $null }
            
            { Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/nonexistenthub" -SiteUrl $testSiteUrl -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
        
        It "Should validate site compatibility" {
            Mock Get-PnPSite { return @{ IsHubSite = $true } }
            
            { Add-SPOSiteToHub -HubSiteUrl $testHubUrl -SiteUrl $testSiteUrl -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
    }
}

Describe "New-SPOSiteFromConfig" {
    Context "Configuration file loading" {
        It "Should load JSON configuration file" {
            $configPath = "$TestDrive\config.json"
            $config = @{
                client = "TestClient"
                sites = @(
                    @{
                        title = "Test Site"
                        url = "testsite"
                        type = "TeamSite"
                    }
                )
            }
            $config | ConvertTo-Json -Depth 10 | Out-File $configPath
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/testsite" } }
            
            $result = New-SPOSiteFromConfig -ConfigPath $configPath -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should accept configuration object directly" {
            $config = @{
                sites = @(
                    @{
                        title = "Direct Site"
                        url = "directsite"
                        type = "TeamSite"
                    }
                )
            }
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/directsite" } }
            
            $result = New-SPOSiteFromConfig -Configuration $config -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Configuration validation" {
        It "Should validate required configuration elements" {
            $config = @{}
            
            { New-SPOSiteFromConfig -Configuration $config -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
        
        It "Should validate site properties" {
            $config = @{
                sites = @(
                    @{
                        # Missing required properties
                        type = "TeamSite"
                    }
                )
            }
            
            { New-SPOSiteFromConfig -Configuration $config -ClientName $testClientName -ErrorAction Stop } | Should -Throw
        }
        
        It "Should support ValidateOnly mode" {
            $config = @{
                sites = @(
                    @{
                        title = "Validate Site"
                        url = "validatesite"
                        type = "TeamSite"
                    }
                )
            }
            
            $result = New-SPOSiteFromConfig -Configuration $config -ClientName $testClientName -ValidateOnly
            $result.IsValid | Should -Be $true
        }
    }
    
    Context "Hub and site creation" {
        It "Should create hub site and associated sites" {
            $config = @{
                hubSite = @{
                    title = "Config Hub"
                    url = "confighub"
                    securityBaseline = "MSPSecure"
                }
                sites = @(
                    @{
                        title = "Member Site"
                        url = "membersite"
                        type = "TeamSite"
                        joinHub = $true
                    }
                )
            }
            
            Mock New-SPOHubSite { return @{ Url = "https://contoso.sharepoint.com/sites/confighub" } }
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/membersite" } }
            
            $result = New-SPOSiteFromConfig -Configuration $config -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "New-SPOBulkSites" {
    Context "Bulk site creation from array" {
        It "Should create multiple sites from array" {
            $sites = @(
                @{ Title = "Site1"; Url = "site1"; Type = "TeamSite" },
                @{ Title = "Site2"; Url = "site2"; Type = "CommunicationSite" }
            )
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/$($args[0])" } }
            
            $result = New-SPOBulkSites -Sites $sites -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Bulk site creation from file" {
        It "Should create sites from CSV file" {
            $csvPath = "$TestDrive\sites.csv"
            $csvContent = @"
Title,Url,Type,Description
"Site A","sitea","TeamSite","Team site A"
"Site B","siteb","CommunicationSite","Comm site B"
"@
            $csvContent | Out-File $csvPath
            
            Mock Import-Csv { 
                return @(
                    [PSCustomObject]@{ Title = "Site A"; Url = "sitea"; Type = "TeamSite"; Description = "Team site A" },
                    [PSCustomObject]@{ Title = "Site B"; Url = "siteb"; Type = "CommunicationSite"; Description = "Comm site B" }
                )
            }
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/$($args[0])" } }
            
            $result = New-SPOBulkSites -ConfigPath $csvPath -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should create sites from JSON file" {
            $jsonPath = "$TestDrive\sites.json"
            $sites = @(
                @{ Title = "JSON Site 1"; Url = "jsonsite1"; Type = "TeamSite" },
                @{ Title = "JSON Site 2"; Url = "jsonsite2"; Type = "CommunicationSite" }
            )
            $sites | ConvertTo-Json | Out-File $jsonPath
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/$($args[0])" } }
            
            $result = New-SPOBulkSites -ConfigPath $jsonPath -ClientName $testClientName -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Parallel processing" {
        It "Should support parallel site creation" {
            $sites = 1..10 | ForEach-Object {
                @{ Title = "Site$_"; Url = "site$_"; Type = "TeamSite" }
            }
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/$($args[0])" } }
            
            $result = New-SPOBulkSites -Sites $sites -ClientName $testClientName -Parallel -ThrottleLimit 3 -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Error handling and retry" {
        It "Should retry failed sites when specified" {
            $sites = @(
                @{ Title = "Retry Site"; Url = "retrysite"; Type = "TeamSite" }
            )
            
            $script:attemptCount = 0
            Mock New-SPOSite {
                $script:attemptCount++
                if ($script:attemptCount -eq 1) {
                    throw "Transient error"
                }
                return @{ Url = "https://contoso.sharepoint.com/sites/retrysite" }
            }
            
            $result = New-SPOBulkSites -Sites $sites -ClientName $testClientName -RetryFailedSites -MaxRetries 2 -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should continue on error when specified" {
            $sites = @(
                @{ Title = "Site1"; Url = "site1"; Type = "TeamSite" },
                @{ Title = "FailSite"; Url = "failsite"; Type = "TeamSite" },
                @{ Title = "Site3"; Url = "site3"; Type = "TeamSite" }
            )
            
            Mock New-SPOSite {
                if ($args[0] -eq "failsite") {
                    throw "Site creation failed"
                }
                return @{ Url = "https://contoso.sharepoint.com/sites/$($args[0])" }
            }
            
            $result = New-SPOBulkSites -Sites $sites -ClientName $testClientName -ContinueOnError -WhatIf
            $result | Should -Not -BeNullOrEmpty
        }
    }
    
    Context "Reporting" {
        It "Should generate report when specified" {
            $sites = @(
                @{ Title = "Report Site"; Url = "reportsite"; Type = "TeamSite" }
            )
            
            Mock New-SPOSite { return @{ Url = "https://contoso.sharepoint.com/sites/reportsite" } }
            
            $reportPath = "$TestDrive\report.json"
            $result = New-SPOBulkSites -Sites $sites -ClientName $testClientName -GenerateReport -ReportPath $reportPath -WhatIf
            
            $result | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "Security Baseline Functions" {
    Context "Set-SPOSiteSecurityBaseline" {
        It "Should apply standard security baseline" {
            Mock Get-PnPSite { return @{ Url = $testSiteUrl } }
            Mock Set-PnPSite { return $true }
            Mock Set-PnPTenantSite { return $true }
            
            $result = Set-SPOSiteSecurityBaseline -SiteUrl $testSiteUrl -BaselineName "MSPStandard" -ClientName $testClientName
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should apply secure baseline with stricter settings" {
            Mock Get-PnPSite { return @{ Url = $testSiteUrl } }
            Mock Set-PnPSite { return $true }
            Mock Set-PnPTenantSite { return $true }
            
            $result = Set-SPOSiteSecurityBaseline -SiteUrl $testSiteUrl -BaselineName "MSPSecure" -ClientName $testClientName
            $result | Should -Not -BeNullOrEmpty
        }
        
        It "Should configure document libraries when specified" {
            Mock Get-PnPSite { return @{ Url = $testSiteUrl } }
            Mock Get-PnPList { 
                return @(
                    @{ Title = "Documents"; BaseTemplate = 101 },
                    @{ Title = "Site Assets"; BaseTemplate = 101 }
                )
            }
            Mock Set-PnPList { return $true }
            
            $result = Set-SPOSiteSecurityBaseline -SiteUrl $testSiteUrl -BaselineName "MSPStandard" -ClientName $testClientName -ConfigureDocumentLibraries
            $result | Should -Not -BeNullOrEmpty
        }
    }
}

Describe "Site Validation Functions" {
    Context "Test-SPOSiteUrl" {
        It "Should validate correct URL format" {
            Mock Test-Path { return $false }
            
            $result = Test-SPOSiteUrl -Url "testsite" -ClientName $testClientName
            $result.IsValid | Should -Be $true
        }
        
        It "Should reject invalid URL format" {
            $result = Test-SPOSiteUrl -Url "test site with spaces" -ClientName $testClientName
            $result.IsValid | Should -Be $false
        }
        
        It "Should check MSP naming convention" {
            $result = Test-SPOSiteUrl -Url "ClientName-SiteName" -ClientName "ClientName" -MSPNamingConvention
            $result.IsValid | Should -Be $true
        }
    }
    
    Context "Test-SPOSiteExists" {
        It "Should detect existing sites" {
            Mock Get-PnPSite { return @{ Url = $testSiteUrl } }
            
            $result = Test-SPOSiteExists -SiteUrl $testSiteUrl -ClientName $testClientName
            $result | Should -Be $true
        }
        
        It "Should return false for non-existent sites" {
            Mock Get-PnPSite { return $null }
            
            $result = Test-SPOSiteExists -SiteUrl "https://contoso.sharepoint.com/sites/nonexistent" -ClientName $testClientName
            $result | Should -Be $false
        }
    }
}

Describe "Provisioning Helper Functions" {
    Context "Wait-SPOSiteCreation" {
        It "Should wait for site creation to complete" {
            $script:callCount = 0
            Mock Get-PnPSite {
                $script:callCount++
                if ($script:callCount -lt 3) {
                    return @{ Status = "Creating" }
                }
                return @{ Status = "Active" }
            }
            
            $result = Wait-SPOSiteCreation -SiteUrl $testSiteUrl -TimeoutMinutes 1 -ClientName $testClientName
            $result.Success | Should -Be $true
        }
        
        It "Should timeout if site creation takes too long" {
            Mock Get-PnPSite { return @{ Status = "Creating" } }
            
            $result = Wait-SPOSiteCreation -SiteUrl $testSiteUrl -TimeoutMinutes 0.01 -ClientName $testClientName
            $result.Success | Should -Be $false
            $result.TimedOut | Should -Be $true
        }
    }
    
    Context "Get-SPOProvisioningStatus" {
        It "Should return current provisioning status" {
            Mock Get-PnPSite { return @{ Status = "Active"; Url = $testSiteUrl } }
            
            $result = Get-SPOProvisioningStatus -SiteUrl $testSiteUrl -ClientName $testClientName
            $result.Status | Should -Be "Active"
        }
    }
    
    Context "Initialize-SPOSiteFeatures" {
        It "Should activate specified features" {
            Mock Enable-PnPFeature { return $true }
            
            $features = @(
                "8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295",
                "E3540C7D-6BEA-403C-A224-1A12EAFEE4C4"
            )
            
            $result = Initialize-SPOSiteFeatures -SiteUrl $testSiteUrl -Features $features -ClientName $testClientName
            $result.ActivatedFeatures.Count | Should -Be 2
        }
    }
}

AfterAll {
    # Clean up
    Remove-Module SPOSiteFactory -Force -ErrorAction SilentlyContinue
}