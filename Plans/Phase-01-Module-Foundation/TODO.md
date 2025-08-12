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
- [x] Create root module folder: `SPOSiteFactory/`
- [x] Create module manifest: `SPOSiteFactory.psd1`
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
- [x] Create root module file: `SPOSiteFactory.psm1`
- [ ] Create folder structure:
  - [x] `Public/` - Exported functions
  - [x] `Public/Provisioning/` - Site creation functions
  - [x] `Public/Security/` - Auditing functions  
  - [x] `Public/Hub/` - Hub management functions
  - [x] `Public/Configuration/` - Config functions
  - [x] `Private/` - Internal helper functions
  - [x] `Data/` - Static data and configurations
  - [x] `Data/Baselines/` - Security baseline files
  - [x] `Data/Templates/` - Site templates
  - [x] `Data/Schemas/` - JSON/YAML schemas
  - [x] `Tests/` - Pester test files
  - [x] `Docs/` - Documentation
  - [x] `Examples/` - Usage examples

### 2. Module Loading Logic
- [x] Implement dot-sourcing in `SPOSiteFactory.psm1`:
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
- [x] Add module initialization code
- [x] Configure module variables
- [x] Set up module scope preferences

### 3. Connection Management Foundation
- [x] Create `Private/Connect-SPOFactory.ps1`:
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
- [x] Implement connection pooling logic
- [x] Add connection state management
- [x] Create disconnect function
- [x] Add connection validation

### 4. Logging Framework Setup
- [x] Create `Private/Write-SPOFactoryLog.ps1`:
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
- [x] Configure PSFramework logging providers
- [x] Set up log file rotation
- [x] Create log initialization function
- [x] Add performance logging

### 5. Error Handling Foundation
- [x] Create `Private/Invoke-SPOFactoryCommand.ps1`:
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
- [x] Define custom exception types
- [x] Create error classification system
- [x] Implement basic retry logic
- [x] Add error logging integration

### 6. Configuration Management Base
- [x] Create `Private/Get-SPOFactoryConfig.ps1`
- [x] Create `Private/Set-SPOFactoryConfig.ps1`
- [x] Define default configuration structure:
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
- [x] Add configuration persistence
- [x] Create configuration validation

### 7. Module Variables & Constants
- [x] Define module-scoped variables:
  ```powershell
  $script:SPOFactoryConnection = $null
  $script:SPOFactoryBaselines = @{}
  $script:SPOFactoryTemplates = @{}
  ```
- [x] Create constants file
- [x] Define SharePoint limits and constraints
- [x] Set up feature IDs and GUIDs

### 8. Helper Utilities
- [x] Create `Private/Test-SPOFactoryConnection.ps1`
- [x] Create `Private/Get-SPOFactoryVersion.ps1`
- [x] Create `Private/Test-SPOFactoryPrerequisites.ps1`
- [x] Add parameter validation helpers
- [x] Create type accelerators

### 9. Initial Data Files
- [x] Create `Data/Baselines/Standard.json`:
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
- [x] Create `Data/Templates/TeamSite.json`
- [x] Create `Data/Templates/CommunicationSite.json`
- [x] Create `Data/Templates/HubSite.json`

### 10. Basic Module Tests
- [x] Create `Tests/SPOSiteFactory.Tests.ps1`:
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
- [x] Add manifest validation tests
- [x] Create structure validation tests
- [x] Add prerequisite checking tests

## Success Criteria
- [x] Module imports successfully without errors
- [x] All folder structure is in place
- [x] Connection management works with SharePoint Online
- [x] Logging writes to file and console
- [x] Basic error handling is functional
- [x] Module passes PSScriptAnalyzer
- [x] Initial tests pass

## Testing Requirements
- [x] Module loads in PowerShell 5.1
- [x] Module loads in PowerShell 7.4+
- [x] Connection to SharePoint Online succeeds
- [x] Logging creates log files
- [x] Error handling catches and logs errors

## Documentation Required
- [x] Module structure diagram
- [x] Connection management flow
- [x] Logging configuration guide
- [x] Error handling patterns

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

**Status**: ✅ COMPLETED  
**Last Updated**: 2024-12-17  
**Assigned To**: Development Team