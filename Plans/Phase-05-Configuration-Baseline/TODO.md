# Phase 5: Configuration & Baseline Management

## Objectives
Implement comprehensive configuration management including JSON/YAML support, baseline snapshots, drift detection, and template management for consistent SharePoint deployments.

## Timeline
**Duration**: 1 week  
**Priority**: High  
**Dependencies**: Phase 3 complete (Security Auditing)

## Prerequisites
- [ ] Security auditing functional
- [ ] Baseline structure defined
- [ ] JSON schema designed
- [ ] Test configurations available

## Tasks

### 1. Configuration Import/Export
- [ ] Create `Public/Configuration/Import-SPOConfiguration.ps1`:
  ```powershell
  function Import-SPOConfiguration {
      [CmdletBinding()]
      param(
          [Parameter(Mandatory, ParameterSetName='File')]
          [string]$Path,
          
          [Parameter(Mandatory, ParameterSetName='String')]
          [string]$Json,
          
          [ValidateSet('JSON', 'YAML')]
          [string]$Format = 'JSON',
          
          [switch]$ValidateOnly
      )
  }
  ```
- [ ] Implement JSON parsing
- [ ] Add YAML support (future)
- [ ] Validate configuration schema
- [ ] Handle nested configurations
- [ ] Support variable substitution

### 2. Configuration Schema Validation
- [ ] Create `Private/Test-SPOConfiguration.ps1`:
  ```powershell
  function Test-SPOConfiguration {
      param(
          [PSCustomObject]$Configuration,
          [string]$SchemaPath
      )
  }
  ```
- [ ] Define JSON schema:
  ```json
  {
    "$schema": "http://json-schema.org/draft-07/schema#",
    "type": "object",
    "properties": {
      "sites": {
        "type": "array",
        "items": {
          "type": "object",
          "required": ["title", "url", "type"],
          "properties": {
            "title": {"type": "string"},
            "url": {"type": "string"},
            "type": {"enum": ["TeamSite", "CommunicationSite"]},
            "securityBaseline": {"type": "string"}
          }
        }
      }
    }
  }
  ```
- [ ] Implement schema validation
- [ ] Generate validation errors
- [ ] Support custom schemas

### 3. Baseline Snapshot Management
- [ ] Create `Public/Configuration/New-SPOBaselineSnapshot.ps1`:
  ```powershell
  function New-SPOBaselineSnapshot {
      param(
          [string[]]$Sites,
          [string]$Name,
          [string]$OutputPath,
          [switch]$IncludeTenantSettings
      )
  }
  ```
- [ ] Capture current state
- [ ] Store configuration
- [ ] Include security settings
- [ ] Add metadata (date, user, etc.)
- [ ] Support incremental snapshots

### 4. Baseline Comparison
- [ ] Create `Public/Configuration/Compare-SPOBaseline.ps1`:
  ```powershell
  function Compare-SPOBaseline {
      param(
          [string]$ReferenceBaseline,
          [string]$DifferenceBaseline,
          [switch]$DetailedOutput
      )
  }
  ```
- [ ] Compare two baselines
- [ ] Identify differences
- [ ] Calculate drift percentage
- [ ] Generate comparison report
- [ ] Support multiple formats

### 5. Drift Detection
- [ ] Create `Public/Configuration/Get-SPOConfigurationDrift.ps1`:
  ```powershell
  function Get-SPOConfigurationDrift {
      param(
          [string[]]$Sites,
          [string]$BaselinePath,
          [int]$DriftThreshold = 10
      )
  }
  ```
- [ ] Compare current vs baseline
- [ ] Identify configuration changes
- [ ] Calculate drift metrics
- [ ] Flag critical changes
- [ ] Generate drift report

### 6. Template Management
- [ ] Create `Public/Configuration/Get-SPOTemplate.ps1`:
  ```powershell
  function Get-SPOTemplate {
      param(
          [string]$Name,
          [switch]$ListAvailable
      )
  }
  ```
- [ ] Create `Public/Configuration/New-SPOTemplate.ps1`
- [ ] Create `Public/Configuration/Set-SPOTemplate.ps1`
- [ ] Define template structure:
  ```json
  {
    "name": "ProjectSite",
    "version": "1.0",
    "baseTemplate": "STS#3",
    "configuration": {
      "lists": [
        {
          "name": "Tasks",
          "template": "TasksList"
        }
      ],
      "libraries": [
        {
          "name": "Project Documents",
          "versioning": true
        }
      ],
      "features": [
        "8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295"
      ],
      "security": {
        "baseline": "High"
      }
    }
  }
  ```

### 7. Configuration Variables
- [ ] Create `Private/Expand-SPOConfigurationVariables.ps1`:
  ```powershell
  function Expand-SPOConfigurationVariables {
      param(
          [PSCustomObject]$Configuration,
          [hashtable]$Variables
      )
  }
  ```
- [ ] Support variable substitution:
  ```json
  {
    "variables": {
      "tenant": "contoso",
      "environment": "prod"
    },
    "sites": [
      {
        "title": "{{environment}} Site",
        "url": "{{environment}}-site"
      }
    ]
  }
  ```
- [ ] Implement token replacement
- [ ] Support environment variables
- [ ] Add computed variables

### 8. Configuration Merge
- [ ] Create `Public/Configuration/Merge-SPOConfiguration.ps1`:
  ```powershell
  function Merge-SPOConfiguration {
      param(
          [PSCustomObject[]]$Configurations,
          [ValidateSet('Override', 'Combine', 'Error')]
          [string]$ConflictResolution = 'Error'
      )
  }
  ```
- [ ] Merge multiple configs
- [ ] Handle conflicts
- [ ] Support inheritance
- [ ] Validate merged result

### 9. Configuration History
- [ ] Create `Private/Add-SPOConfigurationHistory.ps1`:
  ```powershell
  function Add-SPOConfigurationHistory {
      param(
          [string]$Action,
          [PSCustomObject]$Configuration,
          [string]$User
      )
  }
  ```
- [ ] Track configuration changes
- [ ] Store history log
- [ ] Support rollback
- [ ] Generate audit trail

### 10. Baseline Compliance
- [ ] Create `Public/Configuration/Test-SPOBaselineCompliance.ps1`:
  ```powershell
  function Test-SPOBaselineCompliance {
      param(
          [string[]]$Sites,
          [string]$Baseline,
          [int]$PassThreshold = 80
      )
  }
  ```
- [ ] Check compliance status
- [ ] Calculate compliance score
- [ ] Identify violations
- [ ] Generate compliance report

## Configuration Examples

### Site Configuration with Variables
```json
{
  "variables": {
    "department": "Finance",
    "year": "2024"
  },
  "sites": [
    {
      "title": "{{department}} Team {{year}}",
      "url": "{{department}}-{{year}}",
      "type": "TeamSite",
      "securityBaseline": "High",
      "owners": ["{{department}}-admin@contoso.com"],
      "features": {
        "openInClient": true
      }
    }
  ]
}
```

### Baseline Snapshot Structure
```json
{
  "metadata": {
    "name": "Production Baseline",
    "created": "2024-01-15T10:00:00Z",
    "createdBy": "admin@contoso.com",
    "version": "1.0"
  },
  "tenant": {
    "sharingCapability": "ExternalUserSharingOnly",
    "requireAnonymousLinksExpireInDays": 30
  },
  "sites": [
    {
      "url": "https://contoso.sharepoint.com/sites/finance",
      "settings": {
        "denyAddAndCustomizePages": true,
        "defaultItemOpenInBrowser": false
      }
    }
  ]
}
```

## Success Criteria
- [ ] JSON configuration import works
- [ ] Schema validation catches errors
- [ ] Baseline snapshots capture all settings
- [ ] Drift detection identifies changes
- [ ] Templates apply correctly
- [ ] Variable substitution works
- [ ] Configuration merge handles conflicts

## Testing Requirements
- [ ] Test with 10+ configuration files
- [ ] Validate against schema
- [ ] Create and compare baselines
- [ ] Detect configuration drift
- [ ] Apply templates successfully
- [ ] Test variable substitution

## Performance Targets
- Configuration import: < 5 seconds
- Baseline snapshot: < 30 seconds per site
- Drift detection: < 1 minute for 10 sites
- Template application: < 20 seconds

## Documentation Required
- [ ] Configuration file examples
- [ ] Schema documentation
- [ ] Variable reference
- [ ] Template creation guide

---

**Status**: Not Started  
**Last Updated**: [Current Date]  
**Assigned To**: Development Team