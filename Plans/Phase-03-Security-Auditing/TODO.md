# Phase 3: Security Auditing Integration

## Objectives
Integrate comprehensive security auditing capabilities from the POC, including tenant and site-level security assessments, Office file handling audits, and automated remediation features.

## Timeline
**Duration**: 2 weeks  
**Priority**: Critical  
**Dependencies**: Phase 1 & 2 complete

## Prerequisites
- [ ] Core provisioning functions operational
- [ ] Connection management working
- [ ] POC.md code reviewed and understood
- [ ] Security baseline definitions ready

## Tasks

### 1. Port Core Auditing Class
- [ ] Create `Private/SecurityAuditor.ps1`:
  ```powershell
  class SPOSecurityAuditor {
      [string]$TenantUrl
      [hashtable]$TenantSettings
      [array]$SiteAudits
      [hashtable]$SecurityBaseline
      [string]$RemediationMode
      
      SPOSecurityAuditor([string]$tenantUrl, [string]$remediationMode) {
          # Constructor
      }
      
      [hashtable]AuditTenantSettings() { }
      [array]AuditSites([string[]]$siteUrls) { }
      [void]ExecuteRemediation() { }
  }
  ```
- [ ] Port tenant auditing logic from POC
- [ ] Port site auditing logic from POC
- [ ] Adapt error handling to module framework
- [ ] Integrate with module logging

### 2. Main Audit Function
- [ ] Create `Public/Security/Invoke-SPOSecurityAudit.ps1`:
  ```powershell
  function Invoke-SPOSecurityAudit {
      [CmdletBinding()]
      param(
          [Parameter(Mandatory)]
          [string]$TenantUrl,
          
          [string[]]$Sites,
          
          [ValidateSet('ReportOnly', 'Interactive', 'Automatic')]
          [string]$RemediationMode = 'ReportOnly',
          
          [string]$OutputPath,
          
          [string]$Baseline = 'Standard'
      )
  }
  ```
- [ ] Implement audit orchestration
- [ ] Add progress reporting
- [ ] Handle partial failures
- [ ] Generate audit summary

### 3. Tenant Security Auditing
- [ ] Create `Private/Get-SPOTenantSecuritySettings.ps1`:
  ```powershell
  function Get-SPOTenantSecuritySettings {
      param(
          [hashtable]$SecurityBaseline
      )
  }
  ```
- [ ] Audit critical tenant settings:
  - [ ] SharingCapability
  - [ ] ShowEveryoneExceptExternalUsersClaim
  - [ ] ShowAllUsersClaim
  - [ ] EnableRestrictedAccessControl
  - [ ] DisableDocumentLibraryDefaultLabeling
  - [ ] NoAccessRedirectUrl
  - [ ] HideSyncButtonOnTeamSite
  - [ ] DenyAddAndCustomizePages
  - [ ] ConditionalAccessPolicy
  - [ ] DefaultSharingLinkType
  - [ ] DefaultLinkPermission
  - [ ] RequireAnonymousLinksExpireInDays
  - [ ] ExternalUserExpirationInDays
  - [ ] BlockMacSync
  - [ ] DisableReportProblemDialog
- [ ] Compare against baseline
- [ ] Calculate risk scores
- [ ] Generate recommendations

### 4. Site Security Auditing
- [ ] Create `Private/Get-SPOSiteSecuritySettings.ps1`:
  ```powershell
  function Get-SPOSiteSecuritySettings {
      param(
          [string]$SiteUrl,
          [hashtable]$SecurityBaseline
      )
  }
  ```
- [ ] Audit site-level settings:
  - [ ] SharingCapability
  - [ ] DenyAddAndCustomizePages
  - [ ] RestrictedAccessControl
  - [ ] ConditionalAccessPolicy
  - [ ] SensitivityLabel
  - [ ] ExternalUserExpirationInDays
  - [ ] DefaultSharingLinkType
  - [ ] DefaultLinkPermission
- [ ] Check site collection administrators
- [ ] Audit custom scripts status
- [ ] Verify site classification

### 5. Office File Handling Audit
- [ ] Create `Private/Get-SPOOfficeFileHandling.ps1`:
  ```powershell
  function Get-SPOOfficeFileHandling {
      param([string]$SiteUrl)
  }
  ```
- [ ] Check feature activation (8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295)
- [ ] Audit document library settings:
  ```powershell
  # Check DefaultItemOpenInBrowser property
  $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
  foreach ($library in $libraries) {
      $lib = Get-PnPList -Identity $library.Id -Includes DefaultItemOpenInBrowser
      # Assess compliance
  }
  ```
- [ ] Determine security implications
- [ ] Generate remediation scripts

### 6. Security Baseline Management
- [ ] Create `Public/Security/Get-SPOSecurityBaseline.ps1`:
  ```powershell
  function Get-SPOSecurityBaseline {
      param(
          [string]$Name = 'Standard',
          [switch]$ListAvailable
      )
  }
  ```
- [ ] Create `Public/Security/Set-SPOSecurityBaseline.ps1`
- [ ] Create `Public/Security/New-SPOSecurityBaseline.ps1`
- [ ] Define baseline structure:
  ```json
  {
    "name": "High Security",
    "version": "1.0",
    "tenantSettings": {
      "sharingCapability": "Disabled",
      "requireAnonymousLinksExpireInDays": 7,
      "showAllUsersClaim": false,
      "denyAddAndCustomizePages": true
    },
    "siteSettings": {
      "denyAddAndCustomizePages": true,
      "defaultItemOpenInBrowser": false
    }
  }
  ```

### 7. Compliance Checking
- [ ] Create `Public/Security/Test-SPOSecurityCompliance.ps1`:
  ```powershell
  function Test-SPOSecurityCompliance {
      param(
          [string[]]$Sites,
          [string]$Baseline = 'Standard',
          [int]$PassThreshold = 80
      )
  }
  ```
- [ ] Calculate compliance scores
- [ ] Identify critical violations
- [ ] Generate compliance report
- [ ] Return pass/fail status

### 8. Remediation Engine
- [ ] Create `Public/Security/Start-SPOSecurityRemediation.ps1`:
  ```powershell
  function Start-SPOSecurityRemediation {
      param(
          [PSCustomObject]$AuditResults,
          [ValidateSet('Interactive', 'Automatic', 'WhatIf')]
          [string]$Mode = 'Interactive'
      )
  }
  ```
- [ ] Port remediation logic from POC
- [ ] Implement interactive prompts
- [ ] Add automatic remediation
- [ ] Create rollback capability
- [ ] Log all changes

### 9. Risk Assessment
- [ ] Create `Private/Get-SPOSecurityRisk.ps1`:
  ```powershell
  function Get-SPOSecurityRisk {
      param(
          [string]$Setting,
          [object]$CurrentValue,
          [object]$RecommendedValue
      )
  }
  ```
- [ ] Define risk matrix:
  ```powershell
  $riskMatrix = @{
      'SharingCapability' = @{
          'ExternalUserAndGuestSharing' = 'High'
          'ExternalUserSharingOnly' = 'Medium'
          'ExistingExternalUserSharingOnly' = 'Low'
          'Disabled' = 'None'
      }
      'ShowAllUsersClaim' = @{
          $true = 'High'
          $false = 'None'
      }
  }
  ```
- [ ] Calculate aggregate risk scores
- [ ] Prioritize remediation actions

### 10. Audit Caching
- [ ] Create `Private/Get-SPOAuditCache.ps1`
- [ ] Create `Private/Set-SPOAuditCache.ps1`
- [ ] Implement cache expiration
- [ ] Store audit results for comparison
- [ ] Enable drift detection

### 11. Testing Security Auditing
- [ ] Create `Tests/SecurityAuditing.Tests.ps1`:
  ```powershell
  Describe "Security Auditing" {
      Context "Tenant Auditing" {
          It "Audits all required settings" { }
          It "Compares against baseline" { }
          It "Calculates risk correctly" { }
      }
      Context "Site Auditing" {
          It "Audits site settings" { }
          It "Checks Office file handling" { }
      }
      Context "Remediation" {
          It "Generates remediation scripts" { }
          It "Handles interactive mode" { }
      }
  }
  ```

## Integration Points

### With Provisioning (Phase 2)
- Apply security baseline during site creation
- Validate security before provisioning
- Post-creation security verification

### With Reporting (Phase 7)
- Generate security audit reports
- Create compliance dashboards
- Export findings to various formats

## Success Criteria
- [ ] All POC auditing features ported
- [ ] Tenant auditing completes successfully
- [ ] Site auditing works for all site types
- [ ] Office file handling correctly assessed
- [ ] Remediation applies changes correctly
- [ ] Risk scoring accurate
- [ ] Compliance calculation works
- [ ] Integration with provisioning seamless

## Testing Requirements
- [ ] Audit at least 10 different settings
- [ ] Test against 3 security baselines
- [ ] Verify remediation changes
- [ ] Test rollback functionality
- [ ] Validate risk calculations
- [ ] Check compliance scoring

## Performance Targets
- Tenant audit: < 30 seconds
- Single site audit: < 15 seconds
- 10 sites audit: < 2 minutes
- Remediation per setting: < 5 seconds

## Documentation Required
- [ ] Security baseline definitions
- [ ] Risk matrix explanation
- [ ] Remediation guide
- [ ] Compliance scoring methodology

## Next Phase Prerequisites
- Security auditing fully functional
- Baselines properly defined
- Remediation tested and working
- Risk assessment accurate

---

**Status**: Not Started  
**Last Updated**: [Current Date]  
**Assigned To**: Development Team