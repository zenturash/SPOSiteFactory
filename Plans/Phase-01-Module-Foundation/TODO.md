# Phase 1: Module Foundation & Structure

## Objectives
Establish the foundational structure for the SPOSiteFactory PowerShell module with proper organization, manifest configuration, and core infrastructure components.

## Timeline
**Duration**: 2 weeks  
**Priority**: Critical (Blocker for all other phases)

## Prerequisites
- [ ] PowerShell 7.4.6+ installed
- [ ] PnP.PowerShell module installed
- [ ] PSFramework module installed
- [ ] Visual Studio Code with PowerShell extension
- [ ] Git repository initialized

## Tasks

### 1. Module Structure Creation
- [ ] Create root module folder: `SPOSiteFactory/`
- [ ] Create module manifest: `SPOSiteFactory.psd1`
  ```powershell
  # Key manifest properties to configure:
  - RootModule = 'SPOSiteFactory.psm1'
  - ModuleVersion = '0.1.0'
  - Author = 'Your Organization'
  - CompanyName = 'Your Company'
  - Description = 'SharePoint Online Site Factory with Security Auditing'
  - PowerShellVersion = '5.1'
  - RequiredModules = @('PnP.PowerShell', 'PSFramework')
  - FunctionsToExport = @()  # Will be populated as we build
  - CmdletsToExport = @()
  - VariablesToExport = @()
  - AliasesToExport = @()
  ```
- [ ] Create root module file: `SPOSiteFactory.psm1`
- [ ] Create folder structure:
  - [ ] `Public/` - Exported functions
  - [ ] `Public/Provisioning/` - Site creation functions
  - [ ] `Public/Security/` - Auditing functions  
  - [ ] `Public/Hub/` - Hub management functions
  - [ ] `Public/Configuration/` - Config functions
  - [ ] `Private/` - Internal helper functions
  - [ ] `Data/` - Static data and configurations
  - [ ] `Data/Baselines/` - Security baseline files
  - [ ] `Data/Templates/` - Site templates
  - [ ] `Data/Schemas/` - JSON/YAML schemas
  - [ ] `Tests/` - Pester test files
  - [ ] `Docs/` - Documentation
  - [ ] `Examples/` - Usage examples

### 2. Module Loading Logic
- [ ] Implement dot-sourcing in `SPOSiteFactory.psm1`:
  ```powershell
  # Get public and private function files
  $Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -Recurse -ErrorAction SilentlyContinue)
  $Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -Recurse -ErrorAction SilentlyContinue)
  
  # Dot source the files
  foreach ($import in @($Public + $Private)) {
      try {
          . $import.FullName
      }
      catch {
          Write-Error "Failed to import function $($import.FullName): $_"
      }
  }
  
  # Export public functions
  Export-ModuleMember -Function $Public.BaseName
  ```
- [ ] Add module initialization code
- [ ] Configure module variables
- [ ] Set up module scope preferences

### 3. Connection Management Foundation
- [ ] Create `Private/Connect-SPOFactory.ps1`:
  ```powershell
  function Connect-SPOFactory {
      [CmdletBinding()]
      param(
          [Parameter(Mandatory)]
          [string]$TenantUrl,
          
          [ValidateSet('Interactive', 'Certificate', 'AppOnly')]
          [string]$AuthMethod = 'Interactive'
      )
      # Implementation
  }
  ```
- [ ] Implement connection pooling logic
- [ ] Add connection state management
- [ ] Create disconnect function
- [ ] Add connection validation

### 4. Logging Framework Setup
- [ ] Create `Private/Write-SPOFactoryLog.ps1`:
  ```powershell
  function Write-SPOFactoryLog {
      param(
          [string]$Message,
          [ValidateSet('Info', 'Warning', 'Error', 'Debug', 'Verbose')]
          [string]$Level = 'Info'
      )
      # Use PSFramework for logging
  }
  ```
- [ ] Configure PSFramework logging providers
- [ ] Set up log file rotation
- [ ] Create log initialization function
- [ ] Add performance logging

### 5. Error Handling Foundation
- [ ] Create `Private/Invoke-SPOFactoryCommand.ps1`:
  ```powershell
  function Invoke-SPOFactoryCommand {
      param(
          [scriptblock]$ScriptBlock,
          [string]$ErrorMessage,
          [int]$MaxRetries = 3
      )
      # Wrapper for error handling and retry logic
  }
  ```
- [ ] Define custom exception types
- [ ] Create error classification system
- [ ] Implement basic retry logic
- [ ] Add error logging integration

### 6. Configuration Management Base
- [ ] Create `Private/Get-SPOFactoryConfig.ps1`
- [ ] Create `Private/Set-SPOFactoryConfig.ps1`
- [ ] Define default configuration structure:
  ```powershell
  $script:SPOFactoryConfig = @{
      TenantUrl = $null
      DefaultBaseline = 'Standard'
      BatchSize = 100
      MaxRetries = 3
      LogPath = "$env:TEMP\SPOSiteFactory"
      EnableDebugLogging = $false
  }
  ```
- [ ] Add configuration persistence
- [ ] Create configuration validation

### 7. Module Variables & Constants
- [ ] Define module-scoped variables:
  ```powershell
  $script:SPOFactoryConnection = $null
  $script:SPOFactoryBaselines = @{}
  $script:SPOFactoryTemplates = @{}
  ```
- [ ] Create constants file
- [ ] Define SharePoint limits and constraints
- [ ] Set up feature IDs and GUIDs

### 8. Helper Utilities
- [ ] Create `Private/Test-SPOFactoryConnection.ps1`
- [ ] Create `Private/Get-SPOFactoryVersion.ps1`
- [ ] Create `Private/Test-SPOFactoryPrerequisites.ps1`
- [ ] Add parameter validation helpers
- [ ] Create type accelerators

### 9. Initial Data Files
- [ ] Create `Data/Baselines/Standard.json`:
  ```json
  {
    "name": "Standard",
    "version": "1.0",
    "tenantSettings": {
      "sharingCapability": "ExternalUserSharingOnly",
      "requireAnonymousLinksExpireInDays": 30
    },
    "siteSettings": {
      "denyAddAndCustomizePages": true
    }
  }
  ```
- [ ] Create `Data/Templates/TeamSite.json`
- [ ] Create `Data/Templates/CommunicationSite.json`
- [ ] Create `Data/Templates/HubSite.json`

### 10. Basic Module Tests
- [ ] Create `Tests/SPOSiteFactory.Tests.ps1`:
  ```powershell
  Describe "SPOSiteFactory Module" {
      Context "Module Setup" {
          It "Should import without errors" {
              { Import-Module SPOSiteFactory -Force } | Should -Not -Throw
          }
          It "Should export expected functions" {
              $commands = Get-Command -Module SPOSiteFactory
              $commands | Should -Not -BeNullOrEmpty
          }
      }
  }
  ```
- [ ] Add manifest validation tests
- [ ] Create structure validation tests
- [ ] Add prerequisite checking tests

## Success Criteria
- [ ] Module imports successfully without errors
- [ ] All folder structure is in place
- [ ] Connection management works with SharePoint Online
- [ ] Logging writes to file and console
- [ ] Basic error handling is functional
- [ ] Module passes PSScriptAnalyzer
- [ ] Initial tests pass

## Testing Requirements
- [ ] Module loads in PowerShell 5.1
- [ ] Module loads in PowerShell 7.4+
- [ ] Connection to SharePoint Online succeeds
- [ ] Logging creates log files
- [ ] Error handling catches and logs errors

## Documentation Required
- [ ] Module structure diagram
- [ ] Connection management flow
- [ ] Logging configuration guide
- [ ] Error handling patterns

## Dependencies
- PnP.PowerShell 3.0+
- PSFramework 1.7+
- .NET Framework 4.7.2+ or .NET 8.0

## Notes
- Keep all functions small and focused
- Follow PowerShell best practices
- Use approved verbs for function names
- Implement verbose and debug output
- Consider cross-platform compatibility

## Next Phase Prerequisites
Once Phase 1 is complete, we can begin:
- Implementing core provisioning functions (Phase 2)
- Building on the connection management
- Utilizing the logging framework
- Extending error handling

---

**Status**: Not Started  
**Last Updated**: [Current Date]  
**Assigned To**: Development Team