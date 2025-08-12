# Claude Code: SharePoint Security Auditing PowerShell Module Development

## Project Overview

Create a production-ready SharePoint Online Security Auditing and Remediation PowerShell module following PowerShell best practices with proper folder structure, error handling, and comprehensive functionality.

## Module Structure Requirements

Create a PowerShell module with the following folder structure:

```
SharePointSecurityAuditor/
├── SharePointSecurityAuditor.psd1          # Module manifest
├── SharePointSecurityAuditor.psm1          # Root module file
├── Public/                                 # Public functions (exported)
│   ├── Invoke-SPSecurityAudit.ps1          # Main audit function
│   ├── Get-SPSecurityBaseline.ps1          # Get security baseline
│   ├── Set-SPSecurityBaseline.ps1          # Set custom baseline
│   ├── Export-SPSecurityReport.ps1         # Export reports
│   ├── Start-SPSecurityRemediation.ps1     # Manual remediation
│   ├── Test-SPSecurityCompliance.ps1       # Compliance check only
│   ├── Enable-SPTeamsIntegration.ps1       # Enable Teams for SharePoint
│   ├── Test-SPTeamsIntegration.ps1         # Check Teams integration status
│   ├── New-SPTeamsSite.ps1                # Create site with Teams
│   └── Convert-SPSiteToTeams.ps1           # Convert existing site to Teams
├── Private/                                # Private functions (internal)
│   ├── Connect-SPSecurityService.ps1       # Connection management
│   ├── Get-SPTenantSecuritySettings.ps1    # Tenant audit logic
│   ├── Get-SPSiteSecuritySettings.ps1      # Site audit logic
│   ├── Test-SPSecurityConfiguration.ps1    # Configuration validation
│   ├── Invoke-SPSecurityRemediation.ps1    # Remediation engine
│   ├── New-SPSecurityReport.ps1            # Report generation
│   ├── Write-SPSecurityLog.ps1             # Logging functions
│   ├── Get-SPOfficeFileHandling.ps1        # Office app integration
│   ├── Get-SPTeamsIntegrationStatus.ps1    # Teams status checking
│   ├── Enable-SPBulkTeamsIntegration.ps1   # Bulk Teams enablement
│   └── Get-SPTeamsSecuritySettings.ps1     # Teams-specific security audit
├── Data/                                   # Configuration and templates
│   ├── SecurityBaselines.json              # Default security baselines
│   ├── ReportTemplates/                    # HTML/CSS templates
│   │   ├── SecurityReport.html             # Main report template
│   │   └── styles.css                      # Report styling
│   └── Schemas/                            # Validation schemas
│       └── BaselineSchema.json             # Baseline validation
├── Tests/                                  # Pester tests
│   ├── SharePointSecurityAuditor.Tests.ps1 # Module tests
│   ├── Public.Tests.ps1                   # Public function tests
│   └── Private.Tests.ps1                  # Private function tests
├── Docs/                                   # Documentation
│   ├── README.md                           # Module documentation
│   ├── CHANGELOG.md                        # Version history
│   └── Examples/                           # Usage examples
│       ├── BasicUsage.ps1                  # Simple examples
│       └── AdvancedScenarios.ps1           # Complex scenarios
└── Scripts/                                # Build and deployment
    ├── Build.ps1                           # Module build script
    └── Deploy.ps1                          # Deployment script
```

## Teams Integration Features

### SharePoint-Teams Integration Requirements

The module must include comprehensive Teams integration capabilities for SharePoint sites:

#### 1. Teams Enablement Functions
```powershell
# Public functions for Teams integration
Enable-SPTeamsIntegration      # Enable Teams for SharePoint site
Test-SPTeamsIntegration       # Check if site is Teams-enabled
New-SPTeamsSite              # Create new SharePoint site with Teams
Convert-SPSiteToTeams        # Convert existing site to Teams
```

#### 2. Teams Integration Scenarios
**Scenario A: New Site Creation with Teams**
- Create SharePoint team site with Microsoft 365 Group
- Automatically enable Teams functionality
- Configure default channels and tabs

**Scenario B: Enable Teams for Existing Sites**
- Check if site is group-connected
- Connect to new M365 Group if needed (Add-PnPMicrosoft365GroupToSite)
- Enable Teams functionality (New-PnPTeamsTeam)

**Scenario C: Audit Teams Integration Status**
- Identify which sites have Teams enabled
- Check group connectivity status
- Assess Teams integration compliance

### Key PowerShell Patterns for Teams Integration

#### Enable Teams for Group-Connected Site
```powershell
function Enable-SPTeamsIntegration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$TeamDisplayName,
        
        [Parameter(Mandatory = $false)]
        [switch]$CreateGroupIfMissing
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        $site = Get-PnPSite -Includes GroupId
        
        if ($site.GroupId -eq [Guid]::Empty) {
            if ($CreateGroupIfMissing) {
                # Connect site to new M365 Group first
                $groupAlias = "Group$(Get-Random)"
                Add-PnPMicrosoft365GroupToSite -Url $SiteUrl -Alias $groupAlias -DisplayName $TeamDisplayName
                Start-Sleep -Seconds 10 # Wait for group creation
                $site = Get-PnPSite -Includes GroupId
            }
            else {
                throw "Site is not connected to Microsoft 365 Group"
            }
        }
        
        # Enable Teams functionality
        New-PnPTeamsTeam -GroupId $site.GroupId.Guid
        Write-Host "Successfully enabled Teams for $SiteUrl"
    }
    catch {
        Write-Error "Failed to enable Teams: $_"
    }
}
```

#### Bulk Teams Enablement
```powershell
function Enable-SPBulkTeamsIntegration {
    param(
        [string[]]$SiteUrls,
        [string]$TenantAdminUrl
    )
    
    Connect-PnPOnline -Url $TenantAdminUrl -Interactive
    
    $results = @()
    foreach ($siteUrl in $SiteUrls) {
        try {
            $siteInfo = Get-PnPTenantSite -Identity $siteUrl
            
            if ($siteInfo.Template -eq 'GROUP#0') {
                # Site is group-connected, enable Teams
                New-PnPTeamsTeam -GroupId $siteInfo.GroupId.Guid
                $results += [PSCustomObject]@{
                    SiteUrl = $siteUrl
                    Status = "Teams Enabled"
                    GroupId = $siteInfo.GroupId.Guid
                }
            }
            else {
                $results += [PSCustomObject]@{
                    SiteUrl = $siteUrl
                    Status = "Not Group-Connected"
                    GroupId = $null
                }
            }
        }
        catch {
            $results += [PSCustomObject]@{
                SiteUrl = $siteUrl
                Status = "Failed: $($_.Exception.Message)"
                GroupId = $null
            }
        }
    }
    
    return $results
}
```

### Teams Integration Auditing

#### Private Functions for Teams Auditing
```powershell
function Get-SPTeamsIntegrationStatus {
    param(
        [string]$SiteUrl
    )
    
    try {
        Connect-PnPOnline -Url $SiteUrl -Interactive
        $site = Get-PnPSite -Includes GroupId
        
        $teamsStatus = @{
            SiteUrl = $SiteUrl
            IsGroupConnected = $site.GroupId -ne [Guid]::Empty
            GroupId = $site.GroupId
            TeamsEnabled = $false
            TeamId = $null
        }
        
        if ($teamsStatus.IsGroupConnected) {
            try {
                # Check if Teams is enabled for the group
                $team = Get-PnPTeamsTeam -Identity $site.GroupId.Guid -ErrorAction SilentlyContinue
                $teamsStatus.TeamsEnabled = $null -ne $team
                $teamsStatus.TeamId = $team.Id
            }
            catch {
                # Team doesn't exist for this group
                $teamsStatus.TeamsEnabled = $false
            }
        }
        
        return $teamsStatus
    }
    catch {
        Write-Error "Failed to get Teams status for $SiteUrl: $_"
        return $null
    }
}
```

### Integration with Security Auditing

#### Enhanced Site Auditing with Teams Status
```powershell
# Add to existing Get-SPSiteSecuritySettings function
$siteAudit.TeamsIntegration = @{
    IsGroupConnected = $site.GroupId -ne [Guid]::Empty
    GroupId = $site.GroupId
    TeamsEnabled = $false
    TeamDisplayName = $null
    TeamsComplianceRisk = 'Medium'  # Default risk if not properly configured
}

if ($siteAudit.TeamsIntegration.IsGroupConnected) {
    try {
        $team = Get-PnPTeamsTeam -Identity $site.GroupId.Guid -ErrorAction SilentlyContinue
        if ($team) {
            $siteAudit.TeamsIntegration.TeamsEnabled = $true
            $siteAudit.TeamsIntegration.TeamDisplayName = $team.DisplayName
            $siteAudit.TeamsIntegration.TeamsComplianceRisk = 'Low'
        }
    }
    catch {
        # Teams not enabled for group
    }
}
```

### Teams Security Considerations

#### Teams-Specific Security Auditing
```powershell
function Get-SPTeamsSecuritySettings {
    param([string]$TeamId)
    
    $team = Get-PnPTeamsTeam -Identity $TeamId
    
    return @{
        AllowGuestCreateUpdateChannels = $team.AllowGuestCreateUpdateChannels
        AllowGuestDeleteChannels = $team.AllowGuestDeleteChannels
        AllowCreateUpdateChannels = $team.AllowCreateUpdateChannels
        AllowDeleteChannels = $team.AllowDeleteChannels
        AllowAddRemoveApps = $team.AllowAddRemoveApps
        AllowCreateUpdateRemoveTabs = $team.AllowCreateUpdateRemoveTabs
        AllowCreateUpdateRemoveConnectors = $team.AllowCreateUpdateRemoveConnectors
        ShowInTeamsSearchAndSuggestions = $team.ShowInTeamsSearchAndSuggestions
        Classification = $team.Classification
        Visibility = $team.Visibility
    }
}
```

## Enhanced Module Structure with Teams Integration

### 1. Public Functions

**Invoke-SPSecurityAudit**: Main entry point function
- Parameters: TenantUrl, OutputPath, RemediationMode, Sites, Baseline
- Comprehensive parameter validation with proper error messages
- Support for pipeline input
- Advanced parameter sets for different use cases
- Proper help documentation with examples

**Get-SPSecurityBaseline**: Retrieve security baselines
- Support for built-in and custom baselines
- JSON schema validation
- Baseline versioning support

**Test-SPTeamsIntegration**: Check Teams integration status
- Verify if site is group-connected
- Check if Teams is enabled for the group
- Assess Teams security settings
- Support for bulk checking across multiple sites

**Enable-SPTeamsIntegration**: Enable Teams for SharePoint sites
- Support for existing group-connected sites
- Automatic group creation for non-group sites
- Configurable Teams settings and permissions
- Bulk enablement capabilities

**New-SPTeamsSite**: Create new SharePoint site with Teams
- Integrated site and Teams creation
- Pre-configured security settings
- Template support for different use cases
- Automatic baseline compliance

**Convert-SPSiteToTeams**: Convert existing sites to Teams
- Group connectivity assessment
- Seamless Teams integration
- Preserve existing content and permissions
- Migration validation and reporting

### 2. Private Functions

**Connect-SPSecurityService**: Connection management
- Intelligent connection pooling
- Automatic re-authentication handling
- Multiple authentication methods (Interactive, Certificate, App-only)
- Connection health monitoring

**Get-SPTenantSecuritySettings**: Tenant-level auditing
- All critical tenant settings from our POC
- Parallel processing for efficiency
- Comprehensive error handling with retry logic

**Get-SPSiteSecuritySettings**: Site-level auditing
- Site collection and sub-site auditing
- Document library Office file handling settings
- Feature activation status checking
- Permission auditing
- Teams integration status and security settings

**Get-SPTeamsIntegrationStatus**: Teams integration auditing
- Group connectivity verification
- Teams enablement status
- Teams security configuration assessment
- Channel and tab security evaluation

**Enable-SPBulkTeamsIntegration**: Bulk Teams operations
- Multi-site Teams enablement
- Progress tracking and error handling
- Rollback capabilities for failed operations
- Comprehensive reporting of results

## PowerShell Best Practices Implementation

### 1. Module Structure & Organization
- Proper module manifest with all required fields
- Clean separation of public/private functions
- Consistent file naming conventions
- Proper module loading and initialization

### 2. Error Handling & Logging
```powershell
# Example pattern for error handling
try {
    # Operation
}
catch [Microsoft.Identity.Client.MsalUiRequiredException] {
    Write-SPSecurityLog -Level Warning -Message "Authentication required"
    # Handle re-auth
}
catch [System.Net.WebException] {
    Write-SPSecurityLog -Level Error -Message "Network error: $($_.Exception.Message)"
    throw
}
catch {
    Write-SPSecurityLog -Level Error -Message "Unexpected error: $($_.Exception.Message)"
    throw
}
```

### 3. Parameter Validation
- Use proper parameter attributes: [Parameter(Mandatory, Position, ValueFromPipeline)]
- Implement [ValidateSet], [ValidateScript], [ValidatePattern]
- Custom validation classes where appropriate
- Support for parameter binding from pipeline

### 4. Help Documentation
- Complete comment-based help for all public functions
- .SYNOPSIS, .DESCRIPTION, .PARAMETER, .EXAMPLE sections
- .INPUTS and .OUTPUTS documentation
- .NOTES with version and author information

### 5. Output Objects
- Use proper PowerShell objects with TypeName
- Consistent property naming (PascalCase)
- Support for Format.ps1xml files for display formatting
- Pipeline-friendly output

## Security Auditing Features

### 1. Tenant-Level Settings (from POC)
```powershell
# All these settings from our previous POC
$TenantSettings = @(
    'SharingCapability',
    'ShowEveryoneExceptExternalUsersClaim',
    'ShowAllUsersClaim', 
    'EnableRestrictedAccessControl',
    'DisableDocumentLibraryDefaultLabeling',
    'NoAccessRedirectUrl',
    'HideSyncButtonOnTeamSite',
    'DenyAddAndCustomizePages',
    'ConditionalAccessPolicy',
    'DefaultSharingLinkType',
    'DefaultLinkPermission',
    'RequireAnonymousLinksExpireInDays',
    'ExternalUserExpirationInDays',
    'BlockMacSync'
)
```

### 2. Office File Handling (Key Requirement)
- Site collection feature activation status (8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295)
- Document library DefaultItemOpenInBrowser settings
- Compliance assessment based on security best practices
- Automated remediation capabilities

### 3. Teams Integration Security
- Teams enablement status for SharePoint sites
- Group connectivity verification
- Teams-specific security settings audit
- Channel and guest access controls
- Teams app and connector permissions

### 4. Remediation Capabilities (Enhanced)
- SharePoint security settings remediation
- Office file handling configuration
- Teams integration enablement
- Group connectivity establishment
- Bulk operations with progress tracking

## Teams Integration Use Cases and Examples

### Use Case 1: New Site Creation with Teams Integration
```powershell
# Create a new SharePoint site with automatic Teams enablement
New-SPTeamsSite -Title "Project Alpha" -Alias "ProjectAlpha" `
               -Description "Confidential project workspace" `
               -Owners @("admin@contoso.com") `
               -Members @("user1@contoso.com", "user2@contoso.com") `
               -EnableTeams $true `
               -ApplySecurityBaseline "HighSecurity"
```

### Use Case 2: Convert Existing Sites to Teams
```powershell
# Convert existing SharePoint sites to Teams-enabled sites
$sites = @(
    "https://contoso.sharepoint.com/sites/finance",
    "https://contoso.sharepoint.com/sites/hr",
    "https://contoso.sharepoint.com/sites/legal"
)

Convert-SPSiteToTeams -SiteUrls $sites -CreateGroupIfMissing -ReportPath "C:\Reports"
```

### Use Case 3: Audit Teams Integration Across Tenant
```powershell
# Comprehensive audit including Teams integration status
Invoke-SPSecurityAudit -TenantUrl "https://contoso-admin.sharepoint.com" `
                       -IncludeTeamsIntegration $true `
                       -RemediationMode Interactive `
                       -OutputPath "C:\SecurityAudit"
```

### Use Case 4: Bulk Teams Enablement for Compliant Sites
```powershell
# Enable Teams only for sites that pass security compliance
$auditResults = Test-SPSecurityCompliance -TenantUrl "https://contoso-admin.sharepoint.com"
$compliantSites = $auditResults | Where-Object { $_.ComplianceScore -gt 80 }

Enable-SPTeamsIntegration -Sites $compliantSites.Url -Verbose
```

### 1. Configuration Management
```powershell
# Support for configuration files
class SPSecurityConfiguration {
    [hashtable] $TenantBaseline
    [hashtable] $SiteBaseline  
    [hashtable] $RemediationRules
    [hashtable] $AuditScope
}
```

### 2. Batch Processing & Performance
- Use PnP PowerShell batch operations
- Implement parallel processing with runspace pools
- Progress bars for long-running operations
- Efficient memory management

### 3. Reporting Engine
- Multiple output formats (HTML, CSV, JSON, XML)
- Template-based HTML reports with CSS styling
- Executive summary dashboards
- Detailed technical findings
- Remediation action logs

### 4. Remediation Capabilities
- Three modes: ReportOnly, Interactive, Automatic
- Rollback functionality
- Change tracking and logging
- Risk-based remediation prioritization

## Code Quality Requirements

### 1. PowerShell Script Analyzer
- All code must pass PSScriptAnalyzer with no errors
- Follow PSScriptAnalyzer best practice rules
- Custom rules for organization standards

### 2. Testing Strategy
```powershell
# Implement comprehensive Pester tests
Describe "Invoke-SPSecurityAudit" {
    Context "Parameter Validation" {
        It "Should accept valid tenant URL" {
            # Test cases
        }
        It "Should reject invalid tenant URL" {
            # Test cases  
        }
    }
    Context "Functionality" {
        # Mock PnP PowerShell cmdlets
        Mock Connect-PnPOnline { }
        Mock Get-PnPTenant { return $MockTenant }
        
        It "Should return audit results" {
            # Test audit execution
        }
    }
}
```

### 3. Performance Considerations
- Efficient object creation and disposal
- Minimal memory footprint
- Optimized queries to SharePoint
- Connection reuse and pooling

### 4. Security & Compliance
- Secure credential handling
- No plaintext secrets in code
- Audit trail of all changes
- Compliance with security frameworks

## Dependencies & Prerequisites

### Required Modules
```powershell
# In module manifest
RequiredModules = @(
    @{ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0'},
    @{ModuleName = 'PSFramework'; ModuleVersion = '1.7.0'}
)
```

### PowerShell Version
- PowerShell 7.4+ (for cross-platform support)
- .NET 8 framework compatibility
- Windows PowerShell 5.1 fallback support

## Example Implementation Patterns

### 1. Public Function Template
```powershell
function Invoke-SPSecurityAudit {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateScript({
            if ($_ -match '^https://.*\.sharepoint\.com.*') { $true }
            else { throw "Invalid SharePoint tenant URL format" }
        })]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('ReportOnly', 'Interactive', 'Automatic')]
        [string]$RemediationMode = 'ReportOnly'
    )
    
    begin {
        Write-SPSecurityLog -Level Info -Message "Starting security audit for $TenantUrl"
    }
    
    process {
        try {
            # Implementation
        }
        catch {
            Write-SPSecurityLog -Level Error -Message $_.Exception.Message
            throw
        }
    }
    
    end {
        Write-SPSecurityLog -Level Info -Message "Security audit completed"
    }
}
```

### 2. Private Function Template  
```powershell
function Get-SPTenantSecuritySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$SecurityBaseline
    )
    
    # Implementation following same patterns
}
```

## Development Workflow

1. **Initialize module structure** with proper folders and files
2. **Implement core private functions** first (connection, auditing logic)
3. **Create public wrapper functions** with proper parameter validation
4. **Add comprehensive error handling** and logging throughout
5. **Implement configuration management** and baseline support
6. **Create reporting engine** with multiple output formats
7. **Add comprehensive tests** for all functions
8. **Document everything** with proper help and examples
9. **Performance optimization** and batch processing
10. **Final validation** with PSScriptAnalyzer and testing

## Success Criteria

The completed module should:
- ✅ Follow all PowerShell best practices and conventions
- ✅ Include comprehensive tenant and site security auditing
- ✅ Support Office file handling configuration (desktop vs browser)
- ✅ Provide flexible remediation capabilities
- ✅ Generate professional reports in multiple formats
- ✅ Include complete test coverage with Pester
- ✅ Have proper documentation and examples
- ✅ Pass all PSScriptAnalyzer tests
- ✅ Support both Windows and cross-platform PowerShell
- ✅ Be production-ready for enterprise environments

## Additional Considerations

- **Localization support** for international deployments  
- **Plugin architecture** for extending functionality
- **Integration hooks** for SIEM and monitoring systems
- **Compliance reporting** for SOC, ISO, and other frameworks
- **Automated scheduling** capabilities with task scheduler integration
- **Azure DevOps pipeline** integration for CI/CD

This comprehensive approach will result in a professional, maintainable, and scalable SharePoint security auditing solution that follows industry best practices and can be easily extended for future requirements.