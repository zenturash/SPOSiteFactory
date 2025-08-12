# Phase 2: Core Provisioning Functions

## Overview

Phase 2 functions provide comprehensive SharePoint site provisioning capabilities, including hub sites, team sites, communication sites, bulk operations, and template management.

## Site Creation Functions

### New-SPOSite

Creates SharePoint Online sites with comprehensive MSP features and M365 Group integration.

**Syntax:**
```powershell
New-SPOSite -SiteUrl <String> 
            -Title <String>
            -Owner <String>
            -SiteType <String>
            -ClientName <String>
            [-Description <String>]
            [-SecurityBaseline <String>]
            [-Language <Int32>]
            [-TimeZone <Int32>]
            [-Template <String>]
            [-HubSiteUrl <String>]
            [-CreateM365Group]
            [-GroupAlias <String>]
            [-GroupMembers <String[]>]
            [-GroupOwners <String[]>]
            [-SiteDesignId <String>]
            [-WaitForCompletion]
            [-TimeoutMinutes <Int32>]
            [-ApplySecurityBaseline]
            [-ConfigureOfficeFileHandling]
            [-EnableAuditing]
            [-Force]
            [-WhatIf]
```

**Parameters:**
- **SiteUrl** (Required): Full SharePoint site URL
- **Title** (Required): Site title
- **Owner** (Required): Primary owner email address
- **SiteType** (Required): TeamSite or CommunicationSite
- **ClientName** (Required): MSP client identifier
- **Description**: Site description
- **SecurityBaseline**: MSPStandard, MSPSecure, or MSPStrict (default: MSPStandard)
- **Language**: Language ID (default: 1033 for English)
- **TimeZone**: Time zone ID (default: 13 for Eastern)
- **Template**: Site template (default: GROUP#0 for TeamSite, SITEPAGEPUBLISHING#0 for CommunicationSite)
- **HubSiteUrl**: Hub site to associate with
- **CreateM365Group**: Create Microsoft 365 Group (default: true for TeamSite)
- **GroupAlias**: M365 Group alias
- **GroupMembers**: Array of member emails
- **GroupOwners**: Array of additional owner emails
- **SiteDesignId**: Site design to apply
- **WaitForCompletion**: Wait for site to be fully provisioned (default: true)
- **TimeoutMinutes**: Max wait time (default: 30)
- **ApplySecurityBaseline**: Apply security settings (default: true)
- **ConfigureOfficeFileHandling**: Configure Office files to open in desktop apps (default: true)
- **EnableAuditing**: Enable audit logging (default: true)

**Returns:**
```powershell
@{
    Success = $true
    SiteUrl = "https://contoso.sharepoint.com/sites/finance"
    SiteType = "TeamSite"
    Site = [PSObject]
    M365Group = [PSObject]
    CreationTime = [TimeSpan]
    SecurityBaseline = @{Applied = $true; Results = @()}
    HubAssociation = [PSObject]
    Errors = @()
    Warnings = @()
}
```

**Examples:**
```powershell
# Create team site with M365 Group
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/finance" `
            -Title "Finance Team" `
            -Description "Finance department collaboration" `
            -Owner "cfo@contoso.com" `
            -SiteType "TeamSite" `
            -ClientName "Contoso" `
            -GroupMembers @("accountant1@contoso.com", "accountant2@contoso.com") `
            -SecurityBaseline "MSPSecure"

# Create communication site
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/news" `
            -Title "Company News" `
            -Owner "comms@contoso.com" `
            -SiteType "CommunicationSite" `
            -ClientName "Contoso"

# Create site with hub association
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/marketing" `
            -Title "Marketing Team" `
            -Owner "cmo@contoso.com" `
            -SiteType "TeamSite" `
            -ClientName "Contoso" `
            -HubSiteUrl "https://contoso.sharepoint.com/sites/corp-hub"
```

### New-SPOHubSite

Creates SharePoint hub sites with comprehensive configuration and security settings.

**Syntax:**
```powershell
New-SPOHubSite -Title <String>
               -Url <String>
               -ClientName <String>
               [-Description <String>]
               [-SecurityBaseline <String>]
               [-Owners <String[]>]
               [-ApplySecurityImmediately]
               [-RegisterAsHub]
               [-HubSiteDesignId <String>]
               [-EnableHubSiteJoinApproval]
               [-RequireSecurityClearance]
               [-LogoUrl <String>]
               [-Theme <String>]
               [-WhatIf]
```

**Examples:**
```powershell
# Create department hub
New-SPOHubSite -Title "IT Department Hub" `
               -Url "it-hub" `
               -Description "Central hub for IT department sites" `
               -ClientName "Contoso" `
               -SecurityBaseline "MSPSecure" `
               -Owners @("cio@contoso.com", "itmanager@contoso.com") `
               -ApplySecurityImmediately

# Create corporate hub with branding
New-SPOHubSite -Title "Contoso Corporate" `
               -Url "corporate" `
               -ClientName "Contoso" `
               -LogoUrl "https://contoso.sharepoint.com/assets/logo.png" `
               -Theme "Contoso Blue" `
               -EnableHubSiteJoinApproval
```

## Bulk Operations

### New-SPOBulkSites

Creates multiple SharePoint sites efficiently with parallel processing.

**Syntax:**
```powershell
New-SPOBulkSites -Sites <Object[]> | -ConfigPath <String>
                 [-ClientName <String>]
                 [-BatchSize <Int32>]
                 [-Parallel]
                 [-ThrottleLimit <Int32>]
                 [-ContinueOnError]
                 [-RetryFailedSites]
                 [-MaxRetries <Int32>]
                 [-GenerateReport]
                 [-ReportPath <String>]
                 [-WhatIf]
```

**Parameters:**
- **Sites**: Array of site configuration objects
- **ConfigPath**: Path to CSV or JSON file with site definitions
- **BatchSize**: Sites to process per batch (default: 10)
- **Parallel**: Enable parallel processing
- **ThrottleLimit**: Max concurrent operations (default: 5)
- **ContinueOnError**: Don't stop on failures
- **RetryFailedSites**: Automatically retry failed sites
- **MaxRetries**: Max retry attempts (default: 2)
- **GenerateReport**: Create detailed report
- **ReportPath**: Report file location

**CSV Format:**
```csv
Title,Url,Type,Description,Owner,SecurityBaseline,HubSite
"Finance","finance","TeamSite","Finance team","cfo@contoso.com","MSPSecure","corp-hub"
"Marketing","marketing","TeamSite","Marketing team","cmo@contoso.com","MSPStandard","corp-hub"
```

**Examples:**
```powershell
# Bulk create from CSV
New-SPOBulkSites -ConfigPath "C:\Sites\bulk-sites.csv" `
                 -ClientName "Contoso" `
                 -Parallel `
                 -ThrottleLimit 3 `
                 -RetryFailedSites `
                 -GenerateReport

# Bulk create from array
$sites = @(
    @{Title="Project A"; Url="proj-a"; Type="TeamSite"; Owner="pm1@contoso.com"},
    @{Title="Project B"; Url="proj-b"; Type="TeamSite"; Owner="pm2@contoso.com"}
)
New-SPOBulkSites -Sites $sites -ClientName "Contoso" -BatchSize 5

# Bulk create with error handling
New-SPOBulkSites -ConfigPath "sites.json" `
                 -ClientName "Contoso" `
                 -ContinueOnError `
                 -MaxRetries 3 `
                 -GenerateReport `
                 -ReportPath "C:\Reports\bulk-$(Get-Date -Format 'yyyyMMdd').json"
```

### New-SPOSiteFromConfig

Creates SharePoint sites from configuration files with support for hub/spoke architectures.

**Syntax:**
```powershell
New-SPOSiteFromConfig -ConfigPath <String> | -Configuration <Object>
                      [-ClientName <String>]
                      [-ValidateOnly]
                      [-SkipExisting]
                      [-MaxConcurrent <Int32>]
                      [-GenerateReport]
                      [-ReportPath <String>]
                      [-WhatIf]
```

**JSON Configuration:**
```json
{
  "client": "Contoso",
  "hubSite": {
    "title": "Corporate Hub",
    "url": "corp-hub",
    "description": "Main corporate hub",
    "securityBaseline": "MSPSecure",
    "owners": ["admin@contoso.com"]
  },
  "sites": [
    {
      "title": "Finance",
      "url": "finance",
      "type": "TeamSite",
      "joinHub": true,
      "securityBaseline": "MSPSecure",
      "owners": ["cfo@contoso.com"],
      "members": ["finance-team@contoso.com"]
    },
    {
      "title": "Marketing",
      "url": "marketing",
      "type": "TeamSite",
      "joinHub": true,
      "owners": ["cmo@contoso.com"]
    }
  ]
}
```

**Examples:**
```powershell
# Create from JSON file
New-SPOSiteFromConfig -ConfigPath "C:\Config\contoso-sites.json" `
                      -GenerateReport

# Validate configuration only
New-SPOSiteFromConfig -ConfigPath "sites.json" `
                      -ValidateOnly

# Create from hashtable
$config = @{
    hubSite = @{
        title = "Project Hub"
        url = "project-hub"
        securityBaseline = "MSPSecure"
    }
    sites = @(
        @{title="Alpha"; url="alpha"; type="TeamSite"; joinHub=$true}
        @{title="Beta"; url="beta"; type="TeamSite"; joinHub=$true}
    )
}
New-SPOSiteFromConfig -Configuration $config -ClientName "Contoso"
```

## Hub Management

### Add-SPOSiteToHub

Associates SharePoint sites with hub sites.

**Syntax:**
```powershell
Add-SPOSiteToHub -HubSiteUrl <String>
                 -SiteUrl <String> | -SiteUrls <String[]>
                 -ClientName <String>
                 [-EnablePermissionSync]
                 [-ApplyHubNavigation]
                 [-ApplyHubTheme]
                 [-WaitForCompletion]
                 [-TimeoutMinutes <Int32>]
                 [-ContinueOnError]
                 [-MaxConcurrent <Int32>]
                 [-Force]
                 [-WhatIf]
```

**Examples:**
```powershell
# Associate single site
Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/corp-hub" `
                 -SiteUrl "https://contoso.sharepoint.com/sites/finance" `
                 -ClientName "Contoso" `
                 -ApplyHubNavigation `
                 -ApplyHubTheme

# Bulk association
$sites = @(
    "https://contoso.sharepoint.com/sites/site1",
    "https://contoso.sharepoint.com/sites/site2",
    "https://contoso.sharepoint.com/sites/site3"
)
Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/hub" `
                 -SiteUrls $sites `
                 -ClientName "Contoso" `
                 -EnablePermissionSync `
                 -ContinueOnError
```

## Template Management

### Get-SPOSiteTemplate

Retrieves SharePoint site templates.

**Syntax:**
```powershell
Get-SPOSiteTemplate [-Name <String>]
                    [-Path <String>]
                    [-ClientName <String>]
                    [-BuiltIn]
                    [-Online]
                    [-Category <String>]
```

**Examples:**
```powershell
# Get all local templates
Get-SPOSiteTemplate

# Get specific template
Get-SPOSiteTemplate -Name "ProjectSite"

# Get built-in templates
Get-SPOSiteTemplate -BuiltIn

# Get templates by category
Get-SPOSiteTemplate -Category "Department"

# Get online site designs
Get-SPOSiteTemplate -Online -ClientName "Contoso"
```

### New-SPOSiteTemplate

Creates custom site templates.

**Syntax:**
```powershell
New-SPOSiteTemplate -Name <String>
                    -DisplayName <String>
                    -Category <String>
                    [-Description <String>]
                    [-BaseTemplate <String>]
                    [-Libraries <Object[]>]
                    [-Lists <Object[]>]
                    [-Features <String[]>]
                    [-SecurityBaseline <String>]
                    [-ClientName <String>]
                    [-Navigation <Object>]
                    [-Theme <String>]
                    [-OutputPath <String>]
                    [-ExportToSharePoint]
                    [-Force]
```

**Examples:**
```powershell
# Create project template
New-SPOSiteTemplate -Name "ProjectTemplate" `
                    -DisplayName "Project Site Template" `
                    -Description "Template for project sites" `
                    -Category "Project" `
                    -BaseTemplate "GROUP#0" `
                    -Libraries @(
                        @{name="Project Docs"; versioning=$true; checkOut=$true},
                        @{name="Deliverables"; versioning=$true}
                    ) `
                    -Lists @(
                        @{name="Tasks"; template="TasksList"},
                        @{name="Issues"; template="IssuesList"}
                    ) `
                    -SecurityBaseline "MSPSecure"

# Export to SharePoint as site design
New-SPOSiteTemplate -Name "DepartmentTemplate" `
                    -DisplayName "Department Template" `
                    -Category "Department" `
                    -ExportToSharePoint `
                    -ClientName "Contoso"
```

### Set-SPOSiteTemplate

Updates existing site templates.

**Syntax:**
```powershell
Set-SPOSiteTemplate -Name <String> | -Path <String>
                    [-DisplayName <String>]
                    [-Description <String>]
                    [-Category <String>]
                    [-Libraries <Object[]>]
                    [-Lists <Object[]>]
                    [-AddLibrary <Object>]
                    [-RemoveLibrary <String>]
                    [-AddList <Object>]
                    [-RemoveList <String>]
                    [-UpdateSharePoint]
                    [-Force]
```

**Examples:**
```powershell
# Update template description
Set-SPOSiteTemplate -Name "ProjectTemplate" `
                    -Description "Updated project template with new features"

# Add library to template
Set-SPOSiteTemplate -Name "ProjectTemplate" `
                    -AddLibrary @{name="Contracts"; versioning=$true; checkOut=$true}

# Remove list from template
Set-SPOSiteTemplate -Name "ProjectTemplate" `
                    -RemoveList "Issues"
```

## Security Functions

### Set-SPOSiteSecurityBaseline

Applies security baselines to SharePoint sites.

**Syntax:**
```powershell
Set-SPOSiteSecurityBaseline -SiteUrl <String>
                            -BaselineName <String>
                            -ClientName <String>
                            [-ApplyToSite]
                            [-ApplyToLibraries]
                            [-ConfigureDocumentLibraries]
                            [-EnableAuditing]
                            [-Force]
```

**Security Baselines:**

| Baseline | External Sharing | Anonymous Links | DLP | Audit |
|----------|-----------------|-----------------|-----|-------|
| MSPStandard | External users with auth | 30 days | Standard | Basic |
| MSPSecure | Internal only | Disabled | Enhanced | Full |
| MSPStrict | Disabled | Disabled | Maximum | Complete |

**Examples:**
```powershell
# Apply standard baseline
Set-SPOSiteSecurityBaseline -SiteUrl "https://contoso.sharepoint.com/sites/team" `
                            -BaselineName "MSPStandard" `
                            -ClientName "Contoso" `
                            -ApplyToSite `
                            -ConfigureDocumentLibraries

# Apply strict baseline with full audit
Set-SPOSiteSecurityBaseline -SiteUrl "https://contoso.sharepoint.com/sites/confidential" `
                            -BaselineName "MSPStrict" `
                            -ClientName "Contoso" `
                            -ApplyToSite `
                            -ApplyToLibraries `
                            -EnableAuditing `
                            -Force
```

## Validation Functions

### Test-SPOSiteUrl

Validates SharePoint site URLs.

**Syntax:**
```powershell
Test-SPOSiteUrl -Url <String>
                -ClientName <String>
                [-SiteType <String>]
                [-CheckAvailability]
                [-MSPNamingConvention]
```

**Returns:**
```powershell
@{
    IsValid = $true
    IsAvailable = $true
    Url = "testsite"
    FullUrl = "https://contoso.sharepoint.com/sites/testsite"
    ValidationErrors = @()
    Suggestions = @()
}
```

**Examples:**
```powershell
# Validate URL format
Test-SPOSiteUrl -Url "my-site" -ClientName "Contoso"

# Check availability
Test-SPOSiteUrl -Url "finance" -ClientName "Contoso" -CheckAvailability

# Validate MSP naming convention
Test-SPOSiteUrl -Url "contoso-finance" -ClientName "Contoso" -MSPNamingConvention
```

### Test-SPOSiteExists

Checks if a SharePoint site exists.

**Syntax:**
```powershell
Test-SPOSiteExists -SiteUrl <String> -ClientName <String>
```

**Examples:**
```powershell
if (Test-SPOSiteExists -SiteUrl "https://contoso.sharepoint.com/sites/test" -ClientName "Contoso") {
    Write-Host "Site exists"
} else {
    Write-Host "Site does not exist"
}
```

## Helper Functions

### Wait-SPOSiteCreation

Waits for site provisioning to complete.

**Syntax:**
```powershell
Wait-SPOSiteCreation -SiteUrl <String>
                     -TimeoutMinutes <Int32>
                     -ClientName <String>
                     [-ExpectedTitle <String>]
                     [-ExpectedOwner <String>]
                     [-ShowProgress]
```

**Examples:**
```powershell
# Wait with progress bar
$result = Wait-SPOSiteCreation -SiteUrl "https://contoso.sharepoint.com/sites/newsite" `
                                -TimeoutMinutes 5 `
                                -ClientName "Contoso" `
                                -ShowProgress

if ($result.Success) {
    Write-Host "Site is ready!"
}
```

### Get-SPOProvisioningStatus

Gets current site provisioning status.

**Syntax:**
```powershell
Get-SPOProvisioningStatus -SiteUrl <String> -ClientName <String>
```

**Returns:**
```powershell
@{
    Status = "Active"  # Creating, Active, Failed
    IsReady = $true
    PercentComplete = 100
    LastUpdated = [DateTime]
    Details = "Site is fully provisioned"
}
```

### Initialize-SPOSiteFeatures

Activates features on SharePoint sites.

**Syntax:**
```powershell
Initialize-SPOSiteFeatures -SiteUrl <String>
                          -Features <String[]>
                          -ClientName <String>
                          [-Scope <String>]
```

**Examples:**
```powershell
# Activate Office file handling feature
Initialize-SPOSiteFeatures -SiteUrl "https://contoso.sharepoint.com/sites/team" `
                          -Features @("8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295") `
                          -ClientName "Contoso" `
                          -Scope "Web"
```

## Best Practices

### Site Creation
1. Always apply security baselines
2. Use consistent naming conventions
3. Document site ownership clearly
4. Enable auditing for compliance
5. Configure Office file handling for security

### Bulk Operations
1. Use parallel processing for large batches
2. Implement retry logic for transient failures
3. Generate reports for audit trails
4. Use configuration files for repeatability
5. Test with ValidateOnly before execution

### Template Management
1. Create templates for common site types
2. Version control template files
3. Test templates in dev environment
4. Document template configurations
5. Use categories for organization

### Hub Architecture
1. Plan hub hierarchy before creation
2. Apply consistent security across hubs
3. Use hub themes for branding
4. Configure navigation inheritance
5. Monitor hub association limits

## Performance Considerations

| Operation | Expected Time | Optimization |
|-----------|--------------|-------------|
| Single Site | 15-30 seconds | Use WaitForCompletion=false for async |
| Hub Site | 20-40 seconds | Pre-validate all settings |
| Bulk (10 sites) | 2-5 minutes | Use Parallel with ThrottleLimit=5 |
| Bulk (50 sites) | 10-15 minutes | Increase ThrottleLimit to 10 |
| Template Apply | 10-20 seconds | Cache templates locally |
| Hub Association | 5-10 seconds | Batch associations |

## Troubleshooting

### Site Creation Failures
```powershell
# Enable verbose logging
$VerbosePreference = "Continue"

# Test with WhatIf
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/test" `
            -Title "Test" `
            -Owner "admin@contoso.com" `
            -SiteType "TeamSite" `
            -ClientName "Contoso" `
            -WhatIf

# Check logs
Get-SPOFactoryLog -Level Error -Last 10
```

### Bulk Operation Issues
```powershell
# Validate configuration first
New-SPOSiteFromConfig -ConfigPath "sites.json" -ValidateOnly

# Use smaller batches
New-SPOBulkSites -ConfigPath "sites.csv" `
                 -BatchSize 5 `
                 -ThrottleLimit 2 `
                 -ContinueOnError
```

### Template Problems
```powershell
# Validate template structure
$template = Get-SPOSiteTemplate -Name "ProjectTemplate"
$template.Configuration | ConvertTo-Json -Depth 10

# Test template with single site
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/template-test" `
            -Title "Template Test" `
            -Owner "admin@contoso.com" `
            -SiteType "TeamSite" `
            -ClientName "Contoso" `
            -SiteDesignId $template.Id
```

---

**Note**: Phase 2 functions build on Phase 1 foundation. Ensure proper connection and configuration before using provisioning functions.