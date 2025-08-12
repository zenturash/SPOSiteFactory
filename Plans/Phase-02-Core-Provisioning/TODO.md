# Phase 2: Core Provisioning Functions

## Objectives
Implement the fundamental site provisioning capabilities including hub site creation, team/communication site creation, and site-to-hub associations with security baseline application.

## Timeline
**Duration**: 2 weeks  
**Priority**: Critical  
**Dependencies**: Phase 1 must be complete

## Prerequisites
- [x] Module foundation complete (Phase 1)
- [x] Connection management functional
- [x] Logging framework operational
- [x] Test SharePoint tenant available

## Tasks

### 1. Hub Site Creation Function
- [x] Create `Public/Provisioning/New-SPOHubSite.ps1`:
  ```powershell
  function New-SPOHubSite {
      [CmdletBinding(SupportsShouldProcess)]
      param(
          [Parameter(Mandatory)]
          [string]$Title,
          
          [Parameter(Mandatory)]
          [string]$Url,
          
          [string]$Description,
          
          [ValidateSet('Standard', 'High', 'Custom')]
          [string]$SecurityBaseline = 'Standard',
          
          [string[]]$Owners,
          
          [switch]$ApplySecurityImmediately
      )
  }
  ```
- [x] Implement site creation logic
- [x] Add hub site registration
- [x] Apply security baseline
- [x] Configure hub settings:
  - [x] Hub permissions
  - [x] Hub theme
  - [x] Hub logo
  - [x] Search scope
- [x] Add logging for each step
- [x] Implement error handling
- [x] Add WhatIf support

### 2. Team Site Creation Function
- [x] Create `Public/Provisioning/New-SPOSite.ps1`:
  ```powershell
  function New-SPOSite {
      [CmdletBinding(DefaultParameterSetName = 'TeamSite')]
      param(
          [Parameter(Mandatory)]
          [string]$Title,
          
          [Parameter(Mandatory)]
          [string]$Url,
          
          [Parameter(ParameterSetName = 'TeamSite')]
          [switch]$TeamSite,
          
          [Parameter(ParameterSetName = 'CommunicationSite')]
          [switch]$CommunicationSite,
          
          [string]$HubSiteUrl,
          
          [string]$SecurityBaseline = 'Standard',
          
          [string[]]$Owners,
          [string[]]$Members
      )
  }
  ```
- [x] Implement team site creation (GROUP#0 template)
- [x] Implement communication site creation
- [x] Add Microsoft 365 Group creation
- [x] Apply security settings:
  - [x] Sharing capabilities
  - [x] External user settings
  - [x] Custom scripts
  - [x] Access request settings
- [x] Configure default document libraries
- [x] Set Office file handling (Feature ID: 8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295)
- [x] Add site to hub if specified

### 3. Site-to-Hub Association
- [x] Create `Public/Hub/Add-SPOSiteToHub.ps1`:
  ```powershell
  function Add-SPOSiteToHub {
      [CmdletBinding()]
      param(
          [Parameter(Mandatory, ValueFromPipeline)]
          [string[]]$SiteUrl,
          
          [Parameter(Mandatory)]
          [string]$HubSiteUrl,
          
          [switch]$InheritHubPermissions,
          
          [switch]$ApplyHubTheme
      )
  }
  ```
- [x] Implement hub association logic
- [x] Validate site compatibility
- [x] Apply hub permissions if requested
- [x] Apply hub theme if requested
- [x] Update site navigation
- [x] Log association details

### 4. Configuration-Based Site Creation
- [ ] Create `Public/Provisioning/New-SPOSiteFromConfig.ps1`:
  ```powershell
  function New-SPOSiteFromConfig {
      [CmdletBinding()]
      param(
          [Parameter(Mandatory, ParameterSetName = 'File')]
          [string]$ConfigPath,
          
          [Parameter(Mandatory, ParameterSetName = 'Object')]
          [hashtable]$Configuration,
          
          [switch]$ValidateOnly
      )
  }
  ```
- [ ] Implement JSON configuration parsing
- [ ] Create configuration schema:
  ```json
  {
    "sites": [
      {
        "title": "Finance Team",
        "url": "finance",
        "type": "TeamSite",
        "hubSite": "corporate-hub",
        "securityBaseline": "High",
        "owners": ["admin@contoso.com"],
        "members": ["finance-team@contoso.com"]
      }
    ]
  }
  ```
- [ ] Add configuration validation
- [ ] Implement batch site creation
- [ ] Add progress reporting
- [ ] Handle partial failures

### 5. Security Baseline Application
- [x] Create `Private/Set-SPOSiteSecurityBaseline.ps1`:
  ```powershell
  function Set-SPOSiteSecurityBaseline {
      param(
          [string]$SiteUrl,
          [string]$BaselineName = 'Standard'
      )
  }
  ```
- [x] Load baseline from Data/Baselines/
- [x] Apply tenant-level settings:
  - [x] SharingCapability
  - [x] DefaultSharingLinkType
  - [x] DefaultLinkPermission  
  - [x] RequireAnonymousLinksExpireInDays
  - [x] ExternalUserExpirationInDays
- [x] Apply site-level settings:
  - [x] DenyAddAndCustomizePages
  - [x] RestrictedAccessControl
  - [x] ConditionalAccessPolicy
- [x] Configure document libraries:
  - [x] DefaultItemOpenInBrowser = $false
  - [x] Enable versioning
  - [x] Set retention policies

### 6. Site Template Management
- [ ] Create `Public/Configuration/Get-SPOSiteTemplate.ps1`
- [ ] Create `Public/Configuration/New-SPOSiteTemplate.ps1`
- [ ] Define default templates:
  - [ ] Standard Team Site
  - [ ] Project Site
  - [ ] Department Site
  - [ ] Communication Site
- [ ] Template structure:
  ```json
  {
    "name": "Project Site",
    "baseTemplate": "STS#3",
    "libraries": [
      {
        "name": "Project Documents",
        "versioning": true,
        "checkOut": false
      }
    ],
    "features": [
      "8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295"
    ]
  }
  ```

### 7. Site Validation Functions
- [x] Create `Private/Test-SPOSiteUrl.ps1`:
  ```powershell
  function Test-SPOSiteUrl {
      param([string]$Url)
      # Validate URL format and availability
  }
  ```
- [x] Create `Private/Test-SPOSiteExists.ps1`
- [x] Create `Private/Get-SPOSiteStatus.ps1`
- [x] Add URL formatting helpers
- [x] Validate against SharePoint limitations

### 8. Provisioning Helper Functions
- [x] Create `Private/Wait-SPOSiteCreation.ps1`
- [x] Create `Private/Get-SPOProvisioningStatus.ps1`
- [x] Create `Private/Initialize-SPOSiteFeatures.ps1`
- [x] Add retry logic for transient failures
- [x] Implement timeout handling

### 9. Bulk Site Creation
- [ ] Create `Public/Provisioning/New-SPOBulkSites.ps1`:
  ```powershell
  function New-SPOBulkSites {
      param(
          [string]$ConfigPath,
          [int]$BatchSize = 10,
          [switch]$Parallel
      )
  }
  ```
- [ ] Implement sequential processing
- [ ] Add parallel processing option
- [ ] Create progress tracking
- [ ] Generate creation report
- [ ] Handle partial failures

### 10. Testing Site Creation
- [ ] Create `Tests/Provisioning.Tests.ps1`:
  ```powershell
  Describe "Site Provisioning" {
      Context "New-SPOHubSite" {
          It "Creates hub site successfully" { }
          It "Applies security baseline" { }
          It "Handles existing site error" { }
      }
      Context "New-SPOSite" {
          It "Creates team site" { }
          It "Creates communication site" { }
          It "Associates with hub" { }
      }
  }
  ```
- [ ] Mock PnP cmdlets
- [ ] Test parameter validation
- [ ] Test error scenarios
- [ ] Validate security application

## Configuration Examples

### Hub Site with Sub-sites
```json
{
  "hubSite": {
    "title": "Corporate Hub",
    "url": "corporate-hub",
    "description": "Main corporate hub",
    "securityBaseline": "High",
    "owners": ["admin@contoso.com"]
  },
  "sites": [
    {
      "title": "Finance",
      "url": "finance",
      "type": "TeamSite",
      "joinHub": true
    },
    {
      "title": "HR",
      "url": "hr", 
      "type": "TeamSite",
      "joinHub": true
    }
  ]
}
```

### Batch Site Creation
```powershell
# Create multiple sites from configuration
$config = Get-Content -Path "sites.json" | ConvertFrom-Json
New-SPOBulkSites -Configuration $config -Parallel -Verbose

# Create hub and associated sites
New-SPOHubSite -Title "Project Hub" -Url "project-hub" -SecurityBaseline High
"alpha", "beta", "gamma" | ForEach-Object {
    New-SPOSite -Title "Project $_" -Url "project-$_" -TeamSite -HubSiteUrl "project-hub"
}
```

## Success Criteria
- [ ] Can create hub sites with security settings
- [ ] Can create team and communication sites
- [ ] Sites can be associated with hubs
- [ ] Configuration-based creation works
- [ ] Security baselines apply correctly
- [ ] Bulk operations complete successfully
- [ ] All functions have proper error handling
- [ ] Progress reporting works

## Testing Requirements
- [ ] Create at least 5 test sites
- [ ] Verify security settings applied
- [ ] Test hub associations
- [ ] Validate configuration parsing
- [ ] Test error scenarios
- [ ] Verify rollback on failure

## Documentation Required
- [ ] Function help for all public cmdlets
- [ ] Configuration file examples
- [ ] Security baseline documentation
- [ ] Troubleshooting guide

## Performance Targets
- Single site creation: < 30 seconds
- Bulk creation: 10 sites < 5 minutes
- Hub association: < 10 seconds
- Security application: < 20 seconds

## Known Limitations
- SharePoint throttling limits
- Maximum 2000 sites per hub
- Group creation may take time
- Some settings require admin consent

## Next Phase Prerequisites
- Core provisioning functions working
- Security baselines applying correctly
- Hub associations functional
- Configuration parsing complete

---

**Status**: ✅ COMPLETED (Core Functions Implemented)
**Last Updated**: 2024-12-17  
**Assigned To**: Development Team

## Completed Items Summary:
- ✅ Hub Site Creation Function (New-SPOHubSite.ps1)
- ✅ Team/Communication Site Creation (New-SPOSite.ps1)
- ✅ Site-to-Hub Association (Add-SPOSiteToHub.ps1)
- ✅ Security Baseline Application (Set-SPOSiteSecurityBaseline.ps1)
- ✅ Site Validation Functions (Test-SPOSiteUrl.ps1)
- ✅ Provisioning Helper Functions (Wait-SPOSiteCreation.ps1, Get-SPOProvisioningStatus.ps1)
- 🔄 Configuration-Based Creation (Pending)
- 🔄 Bulk Site Creation (Pending)
- 🔄 Template Management (Pending)
- 🔄 Testing Suite (Pending)