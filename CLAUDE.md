# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a SharePoint Online Security Auditing and Remediation PowerShell module that provides comprehensive security auditing capabilities for SharePoint tenants and sites. The tool audits security settings, offers automated remediation, and generates detailed reports in multiple formats.

## Module Architecture

The project follows PowerShell module best practices with this structure:

```
SharePointSecurityAuditor/
├── SharePointSecurityAuditor.psd1          # Module manifest
├── SharePointSecurityAuditor.psm1          # Root module
├── Public/                                 # Exported functions
│   ├── Invoke-SPSecurityAudit.ps1          # Main audit entry point
│   ├── Get-SPSecurityBaseline.ps1          # Retrieve baselines
│   ├── Start-SPSecurityRemediation.ps1     # Remediation execution
│   └── Export-SPSecurityReport.ps1         # Report generation
├── Private/                                # Internal functions
│   ├── Connect-SPSecurityService.ps1       # Connection management
│   ├── Get-SPTenantSecuritySettings.ps1    # Tenant auditing
│   ├── Get-SPSiteSecuritySettings.ps1      # Site auditing
│   └── Get-SPOfficeFileHandling.ps1        # Office file settings
├── Data/                                   # Configuration
│   └── SecurityBaselines.json              # Default baselines
└── Tests/                                  # Pester tests
```

## Development Commands

### Module Installation
```powershell
# Install required modules
Install-Module PnP.PowerShell -Scope CurrentUser
Install-Module Microsoft.PowerShell.SecretManagement -Scope CurrentUser
Install-Module PSFramework -Scope CurrentUser

# Verify installation
Get-Module PnP.PowerShell -ListAvailable
```

### Running the Auditor
```powershell
# Basic audit with interactive remediation
.\SharePointSecurityAuditor.ps1 -TenantUrl "https://tenant-admin.sharepoint.com" -RemediationMode Interactive

# Audit specific sites with automatic remediation
.\SharePointSecurityAuditor.ps1 -TenantUrl "https://tenant-admin.sharepoint.com" -Sites @("site1","site2") -RemediationMode Automatic

# Report-only mode
.\SharePointSecurityAuditor.ps1 -TenantUrl "https://tenant-admin.sharepoint.com" -RemediationMode ReportOnly
```

### Testing
```powershell
# Run Pester tests
Invoke-Pester -Path .\Tests\

# Run PSScriptAnalyzer
Invoke-ScriptAnalyzer -Path . -Recurse
```

### Building Module
```powershell
# Build module for distribution
.\Scripts\Build.ps1

# Deploy to PowerShell Gallery
.\Scripts\Deploy.ps1
```

## Key Security Settings Audited

### Tenant-Level Settings
- `SharingCapability` - External sharing configuration
- `RequireAnonymousLinksExpireInDays` - Link expiration enforcement
- `ShowAllUsersClaim` - User visibility settings
- `DenyAddAndCustomizePages` - Custom script permissions
- `ConditionalAccessPolicy` - Device access controls
- `DefaultLinkPermission` - Default sharing permissions
- `ExternalUserExpirationInDays` - External user lifecycle

### Office File Handling (Critical Feature)
- **Feature ID**: `8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295` - "Open Documents in Client Applications by Default"
- **Library Setting**: `DefaultItemOpenInBrowser` property
  - `$false` = Open in desktop app (recommended for security)
  - `$true` = Open in browser
  - `$null` = Follow site default

### Site-Level Settings
- Site sharing capabilities
- Custom script permissions
- Document library configurations
- External access controls
- Site collection administrators

## Technical Requirements

- **PowerShell**: 7.4.6+ (cross-platform support)
- **.NET Framework**: .NET 8
- **PnP PowerShell**: v3.0+ (815+ cmdlets)
- **Authentication**: Interactive, Certificate, App-only

## Core Functionality

### Main Functions
- `Invoke-SPSecurityAudit`: Primary audit execution with parameters for TenantUrl, RemediationMode (Automatic/Interactive/ReportOnly), Sites, OutputPath
- `Get-SPSecurityBaseline`: Retrieve/manage security baselines
- `Start-SPSecurityRemediation`: Execute remediation actions
- `Export-SPSecurityReport`: Generate HTML/CSV/JSON reports

### Remediation Modes
1. **ReportOnly**: Audit without changes
2. **Interactive**: Prompt for each remediation
3. **Automatic**: Auto-remediate high/medium risks

### Batch Processing
- Uses PnP PowerShell batch operations (100-request chunks)
- Parallel execution with ForEach-Object -Parallel
- Runspace pools for large-scale operations

## Error Handling Patterns

```powershell
try {
    # Operation
}
catch [Microsoft.Identity.Client.MsalUiRequiredException] {
    # Re-authentication required
    Connect-PnPOnline -Url $TenantUrl -Interactive
}
catch [System.Net.WebException] {
    # Network error - implement retry logic
}
```

## Report Outputs

- **HTML**: Interactive dashboards with risk heat maps
- **CSV**: Integration with SIEM systems  
- **JSON**: Programmatic processing
- **Logs**: PSFramework detailed logging

## Performance Optimizations

- Connection pooling and reuse
- Parallel site processing (ThrottleLimit: 5)
- Batch API operations
- Memory-efficient object disposal
- 10x faster operations with PnP v3.0

## Security Best Practices

- No plaintext secrets in code
- Secure credential management via SecretManagement module
- Audit trail of all changes
- Risk-based remediation prioritization
- Rollback functionality for changes