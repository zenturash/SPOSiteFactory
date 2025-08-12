# Phase 2 Implementation Prompt for SPOSiteFactory Module

## Overview
We need to implement Phase 2 (Core Provisioning Functions) of the SPOSiteFactory PowerShell module. This phase builds on the foundation from Phase 1 to add site creation, hub management, and security baseline application capabilities for MSP environments.

## Current Status
- ✅ Phase 1 Complete: Module foundation, connection management, logging, error handling
- 🚀 Phase 2 Starting: Core provisioning functions for SharePoint sites

## Agent Coordination Strategy

### todo-tracker-coordinator Agent Tasks:
1. Load and track all tasks from `Plans/Phase-02-Core-Provisioning/TODO.md`
2. Prioritize provisioning tasks based on dependencies
3. Monitor implementation progress for each function
4. Coordinate between hub site creation and sub-site association tasks
5. Track security baseline application across all site types
6. Ensure all 10 main task groups are completed

### powershell-msp-automation Agent Tasks:
1. Implement core provisioning functions following MSP best practices
2. Create hub site provisioning with multi-tenant support
3. Build team/communication site creation functions
4. Implement site-to-hub association logic
5. Create configuration-based bulk provisioning
6. Apply security baselines consistently across all sites
7. Add MSP-specific features (client isolation, audit trails)

## Core Functions to Implement

### 1. Hub Site Creation (Public/Provisioning/New-SPOHubSite.ps1)
```powershell
function New-SPOHubSite {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [string]$Url,
        
        [string]$ClientName,  # MSP: Client identifier
        
        [string]$Description,
        
        [ValidateSet('Standard', 'High', 'MSPStandard', 'MSPSecure', 'Custom')]
        [string]$SecurityBaseline = 'MSPStandard',
        
        [string[]]$Owners,
        
        [switch]$ApplySecurityImmediately,
        
        [switch]$RegisterAsHub
    )
}
```

**MSP Requirements:**
- Support client-specific URL namespacing
- Apply MSP security baseline by default
- Log all provisioning actions for audit
- Support bulk owner assignment
- Track provisioning in client database

### 2. Team/Communication Site Creation (Public/Provisioning/New-SPOSite.ps1)
```powershell
function New-SPOSite {
    [CmdletBinding(DefaultParameterSetName = 'TeamSite')]
    param(
        [Parameter(Mandatory)]
        [string]$Title,
        
        [Parameter(Mandatory)]
        [string]$Url,
        
        [string]$ClientName,  # MSP: Required for multi-tenant
        
        [Parameter(ParameterSetName = 'TeamSite')]
        [switch]$TeamSite,
        
        [Parameter(ParameterSetName = 'CommunicationSite')]
        [switch]$CommunicationSite,
        
        [string]$HubSiteUrl,
        
        [string]$SecurityBaseline = 'MSPStandard',
        
        [string[]]$Owners,
        [string[]]$Members,
        
        [hashtable]$MSPMetadata  # Client-specific metadata
    )
}
```

**Key Features:**
- Create Microsoft 365 Group for Team sites
- Apply Office file handling settings (Feature: 8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295)
- Set DefaultItemOpenInBrowser = $false for security
- Configure external sharing per MSP policy
- Enable versioning and retention

### 3. Site-to-Hub Association (Public/Hub/Add-SPOSiteToHub.ps1)
```powershell
function Add-SPOSiteToHub {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string[]]$SiteUrl,
        
        [Parameter(Mandatory)]
        [string]$HubSiteUrl,
        
        [string]$ClientName,
        
        [switch]$InheritHubPermissions,
        
        [switch]$ApplyHubTheme,
        
        [switch]$UpdateNavigation
    )
}
```

### 4. Configuration-Based Site Creation (Public/Provisioning/New-SPOSiteFromConfig.ps1)
```powershell
function New-SPOSiteFromConfig {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$ConfigPath,
        
        [Parameter(Mandatory, ParameterSetName = 'Object')]
        [hashtable]$Configuration,
        
        [string]$ClientName,  # MSP: Override config client
        
        [switch]$ValidateOnly,
        
        [switch]$WhatIf
    )
}
```

**Configuration Schema Example:**
```json
{
  "client": "Contoso",
  "hubSite": {
    "title": "Contoso Hub",
    "url": "contoso-hub",
    "description": "Main hub for Contoso",
    "securityBaseline": "MSPSecure",
    "owners": ["admin@msp.com", "admin@contoso.com"]
  },
  "sites": [
    {
      "title": "Finance Team",
      "url": "contoso-finance",
      "type": "TeamSite",
      "hubSite": "contoso-hub",
      "securityBaseline": "MSPSecure",
      "owners": ["finance-lead@contoso.com"],
      "members": ["finance-team@contoso.com"],
      "features": {
        "openInClient": true,
        "versioning": true,
        "retention": 7
      }
    },
    {
      "title": "HR Team",
      "url": "contoso-hr",
      "type": "TeamSite",
      "hubSite": "contoso-hub"
    }
  ]
}
```

### 5. Security Baseline Application (Private/Set-SPOSiteSecurityBaseline.ps1)
```powershell
function Set-SPOSiteSecurityBaseline {
    param(
        [string]$SiteUrl,
        [string]$BaselineName = 'MSPStandard',
        [string]$ClientName
    )
}
```

**Security Settings to Apply:**

**Tenant-Level:**
- SharingCapability (per client policy)
- DefaultSharingLinkType (Internal/Direct)
- DefaultLinkPermission (View only)
- RequireAnonymousLinksExpireInDays (7-30 days)
- ExternalUserExpirationInDays (30-90 days)

**Site-Level:**
- DenyAddAndCustomizePages = $true
- RestrictedAccessControl = $true
- Office file handling (DefaultItemOpenInBrowser = $false)
- Versioning enabled
- Audit logging enabled

### 6. Bulk Site Creation (Public/Provisioning/New-SPOBulkSites.ps1)
```powershell
function New-SPOBulkSites {
    param(
        [string]$ConfigPath,
        [string]$ClientName,
        [int]$BatchSize = 10,
        [switch]$Parallel,
        [switch]$GenerateReport
    )
}
```

## MSP-Specific Implementation Requirements

### 1. Multi-Tenant Isolation
- Each site URL must include client identifier
- Separate security baselines per client
- Isolated logging per client
- Client-specific connection context

### 2. Provisioning Tracking
```powershell
# Track all provisioning in MSP database
$provisioningRecord = @{
    ClientName = $ClientName
    SiteUrl = $SiteUrl
    SiteType = $SiteType
    CreatedBy = $env:USERNAME
    CreatedDate = Get-Date
    SecurityBaseline = $SecurityBaseline
    Status = 'Provisioning'
    MSPTicket = $TicketNumber
}
```

### 3. Audit Requirements
- Log all site creation requests
- Track security baseline applications
- Record hub associations
- Monitor provisioning failures
- Generate client provisioning reports

### 4. Error Handling
- Rollback on provisioning failure
- Clean up partial provisions
- Notify MSP team of failures
- Create support tickets automatically

### 5. Performance Optimization
- Batch operations for multiple sites
- Connection pooling per client
- Parallel processing for bulk creation
- Progress tracking for long operations

## Testing Requirements

### Unit Tests to Create
1. Test hub site creation with various parameters
2. Test team vs communication site creation
3. Test security baseline application
4. Test configuration validation
5. Test bulk operations
6. Test error scenarios and rollback

### Integration Tests
1. Create hub with 5 associated sites
2. Apply security baseline and verify
3. Test cross-client isolation
4. Verify audit trail completeness

## Success Criteria for Phase 2

- [ ] Can create hub sites with MSP security
- [ ] Can create team and communication sites
- [ ] Sites associate with hubs correctly
- [ ] Configuration-based creation works
- [ ] Security baselines apply consistently
- [ ] Bulk operations handle 50+ sites
- [ ] All operations logged for audit
- [ ] Client isolation maintained
- [ ] Error handling with rollback works
- [ ] Tests achieve 80% coverage

## Implementation Priority

### Week 1 Tasks:
1. **Day 1-2**: Implement New-SPOHubSite with MSP features
2. **Day 3-4**: Implement New-SPOSite for both site types
3. **Day 5**: Create Add-SPOSiteToHub association

### Week 2 Tasks:
1. **Day 1-2**: Implement configuration-based creation
2. **Day 3**: Security baseline application
3. **Day 4**: Bulk operations with parallel processing
4. **Day 5**: Testing and documentation

## Files to Create/Modify

### New Files:
- `Public/Provisioning/New-SPOHubSite.ps1`
- `Public/Provisioning/New-SPOSite.ps1`
- `Public/Provisioning/New-SPOSiteFromConfig.ps1`
- `Public/Provisioning/New-SPOBulkSites.ps1`
- `Public/Hub/Add-SPOSiteToHub.ps1`
- `Private/Set-SPOSiteSecurityBaseline.ps1`
- `Private/Test-SPOSiteUrl.ps1`
- `Private/Wait-SPOSiteCreation.ps1`
- `Tests/Provisioning.Tests.ps1`

### Configuration Files:
- `Data/Configurations/MSPSiteTemplates.json`
- `Data/Configurations/ClientDefaults.json`

## Code Quality Standards
- All functions must have comment-based help
- Parameter validation for all inputs
- Proper error handling with try/catch
- MSP client context in all operations
- Logging at Info, Warning, Error levels
- Support for -WhatIf and -Confirm
- Pipeline support where appropriate

## Notes for Agents

**For todo-tracker-coordinator:**
- Track each of the 10 main task groups from TODO.md
- Monitor dependencies between tasks
- Ensure security baseline is implemented before site creation
- Verify all MSP requirements are addressed

**For powershell-msp-automation:**
- Focus on MSP multi-tenant scenarios
- Implement robust error handling with rollback
- Ensure all operations are auditable
- Use existing Phase 1 functions (Connect-SPOFactory, Write-SPOFactoryLog, etc.)
- Test with multiple client contexts
- Consider scale (100+ sites per client)

---

**Start Implementation:**
1. Load Phase 2 TODO.md tasks
2. Create New-SPOHubSite function first (foundation)
3. Implement New-SPOSite with both site types
4. Add configuration-based creation
5. Implement security baseline application
6. Test each component thoroughly