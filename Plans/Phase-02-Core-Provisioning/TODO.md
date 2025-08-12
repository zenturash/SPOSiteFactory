# Phase 2: Core Provisioning Functions

## Objectives
Implement the fundamental site provisioning capabilities including hub site creation, team/communication site creation, and site-to-hub associations with security baseline application.

## Timeline
**Duration**: 2 weeks  
**Priority**: Critical  
**Dependencies**: Phase 1 must be complete

## Prerequisites
- [ ] Module foundation complete (Phase 1)
- [ ] Connection management functional
- [ ] Logging framework operational
- [ ] Test SharePoint tenant available

## Tasks

### 1. Hub Site Creation Function
- [ ] Create `Public/Provisioning/New-SPOHubSite.ps1`:
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
- [ ] Implement site creation logic
- [ ] Add hub site registration
- [ ] Apply security baseline
- [ ] Configure hub settings:
  - [ ] Hub permissions
  - [ ] Hub theme
  - [ ] Hub logo
  - [ ] Search scope
- [ ] Add logging for each step
- [ ] Implement error handling
- [ ] Add WhatIf support

### 2. Team Site Creation Function
- [ ] Create `Public/Provisioning/New-SPOSite.ps1`:
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
- [ ] Implement team site creation (GROUP#0 template)
- [ ] Implement communication site creation
- [ ] Add Microsoft 365 Group creation
- [ ] Apply security settings:
  - [ ] Sharing capabilities
  - [ ] External user settings
  - [ ] Custom scripts
  - [ ] Access request settings
- [ ] Configure default document libraries
- [ ] Set Office file handling (Feature ID: 8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295)
- [ ] Add site to hub if specified

### 3. Site-to-Hub Association
- [ ] Create `Public/Hub/Add-SPOSiteToHub.ps1`:
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
- [ ] Implement hub association logic
- [ ] Validate site compatibility
- [ ] Apply hub permissions if requested
- [ ] Apply hub theme if requested
- [ ] Update site navigation
- [ ] Log association details

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
- [ ] Create `Private/Set-SPOSiteSecurityBaseline.ps1`:
  ```powershell
  function Set-SPOSiteSecurityBaseline {
      param(
          [string]$SiteUrl,
          [string]$BaselineName = 'Standard'
      )
  }
  ```
- [ ] Load baseline from Data/Baselines/
- [ ] Apply tenant-level settings:
  - [ ] SharingCapability
  - [ ] DefaultSharingLinkType
  - [ ] DefaultLinkPermission  
  - [ ] RequireAnonymousLinksExpireInDays
  - [ ] ExternalUserExpirationInDays
- [ ] Apply site-level settings:
  - [ ] DenyAddAndCustomizePages
  - [ ] RestrictedAccessControl
  - [ ] ConditionalAccessPolicy
- [ ] Configure document libraries:
  - [ ] DefaultItemOpenInBrowser = $false
  - [ ] Enable versioning
  - [ ] Set retention policies

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
- [ ] Create `Private/Test-SPOSiteUrl.ps1`:
  ```powershell
  function Test-SPOSiteUrl {
      param([string]$Url)
      # Validate URL format and availability
  }
  ```
- [ ] Create `Private/Test-SPOSiteExists.ps1`
- [ ] Create `Private/Get-SPOSiteStatus.ps1`
- [ ] Add URL formatting helpers
- [ ] Validate against SharePoint limitations

### 8. Provisioning Helper Functions
- [ ] Create `Private/Wait-SPOSiteCreation.ps1`
- [ ] Create `Private/Get-SPOProvisioningStatus.ps1`
- [ ] Create `Private/Initialize-SPOSiteFeatures.ps1`
- [ ] Add retry logic for transient failures
- [ ] Implement timeout handling

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

**Status**: Not Started  
**Last Updated**: [Current Date]  
**Assigned To**: Development Team