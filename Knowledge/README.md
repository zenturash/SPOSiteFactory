# SPOSiteFactory Knowledge Base

## 📚 Documentation Overview

Welcome to the SPOSiteFactory Knowledge Base! This comprehensive documentation covers all functions implemented in Phase 1 (Module Foundation) and Phase 2 (Core Provisioning Functions).

## 📁 Documentation Structure

```
Knowledge/
├── README.md                    # This file - main documentation index
├── QuickStart/
│   ├── Installation.md          # Module installation and setup
│   ├── FirstSite.md            # Create your first SharePoint site
│   ├── HubArchitecture.md     # Setting up hub/spoke architecture
│   └── BulkProvisioning.md    # Bulk site provisioning guide
├── Functions/
│   ├── Phase1-Foundation.md    # Phase 1 function reference
│   ├── Phase2-Provisioning.md  # Phase 2 function reference
│   ├── ConnectionManagement.md # Connection and authentication
│   ├── SecurityBaselines.md    # Security configuration
│   └── TemplateManagement.md   # Site templates reference
├── Examples/
│   ├── BasicExamples.ps1       # Simple usage examples
│   ├── AdvancedScenarios.ps1   # Complex provisioning scenarios
│   ├── MSPWorkflows.ps1        # MSP-specific workflows
│   └── ConfigurationFiles/     # Sample JSON/CSV configurations
└── Troubleshooting/
    ├── CommonIssues.md          # Common problems and solutions
    ├── ErrorCodes.md            # Error code reference
    └── Debugging.md             # Debugging techniques
```

## 🚀 Quick Start Guides

### For New Users
1. [Installation Guide](QuickStart/Installation.md) - Get the module installed and configured
2. [Create Your First Site](QuickStart/FirstSite.md) - Step-by-step guide to create a SharePoint site
3. [Hub Architecture Setup](QuickStart/HubArchitecture.md) - Create hub sites with associated team sites

### For MSP Administrators
1. [Bulk Provisioning](QuickStart/BulkProvisioning.md) - Provision multiple sites efficiently
2. [MSP Workflows](Examples/MSPWorkflows.ps1) - Multi-tenant management patterns
3. [Security Baselines](Functions/SecurityBaselines.md) - Apply consistent security settings

## 📖 Function Categories

### Phase 1: Module Foundation
- **Connection Management**: Connect to SharePoint Online tenants
- **Logging Framework**: Enterprise logging with PSFramework
- **Error Handling**: Comprehensive error management
- **Module Configuration**: Settings and preferences

[View Phase 1 Functions →](Functions/Phase1-Foundation.md)

### Phase 2: Core Provisioning
- **Site Creation**: Create team and communication sites
- **Hub Management**: Create and manage hub sites
- **Bulk Operations**: Provision multiple sites at once
- **Template Management**: Use and create site templates
- **Security Baselines**: Apply security configurations

[View Phase 2 Functions →](Functions/Phase2-Provisioning.md)

## 💡 Common Use Cases

### Single Site Creation
```powershell
# Create a team site with M365 Group
New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/ProjectAlpha" `
            -Title "Project Alpha" `
            -Owner "admin@contoso.com" `
            -SiteType "TeamSite" `
            -ClientName "Contoso" `
            -SecurityBaseline "MSPSecure"
```

### Hub and Spoke Architecture
```powershell
# Create hub site
New-SPOHubSite -Title "Corporate Hub" `
               -Url "corp-hub" `
               -ClientName "Contoso" `
               -SecurityBaseline "MSPSecure"

# Create and associate team sites
@("Finance", "HR", "IT") | ForEach-Object {
    New-SPOSite -Title "$_ Team" `
                -Url $_.ToLower() `
                -SiteType "TeamSite" `
                -HubSiteUrl "https://contoso.sharepoint.com/sites/corp-hub" `
                -ClientName "Contoso"
}
```

### Configuration-Based Provisioning
```powershell
# Create sites from JSON configuration
New-SPOSiteFromConfig -ConfigPath "sites.json" `
                      -ClientName "Contoso" `
                      -GenerateReport
```

### Bulk Site Creation
```powershell
# Create multiple sites from CSV
New-SPOBulkSites -ConfigPath "sites.csv" `
                 -ClientName "Contoso" `
                 -Parallel `
                 -RetryFailedSites `
                 -GenerateReport
```

## 🔧 Module Components

### Core Cmdlets (21 Total)

#### Connection & Configuration (4)
- `Connect-SPOFactory` - Establish SharePoint connection
- `Disconnect-SPOFactory` - Close SharePoint connection
- `Test-SPOFactoryConnection` - Verify connection status
- `Get-SPOFactoryConfig` - Get module configuration

#### Site Provisioning (6)
- `New-SPOHubSite` - Create hub sites
- `New-SPOSite` - Create team/communication sites
- `New-SPOSiteFromConfig` - Create sites from configuration
- `New-SPOBulkSites` - Bulk site creation
- `Add-SPOSiteToHub` - Associate sites with hubs
- `Set-SPOSiteSecurityBaseline` - Apply security settings

#### Template Management (3)
- `Get-SPOSiteTemplate` - Retrieve site templates
- `New-SPOSiteTemplate` - Create custom templates
- `Set-SPOSiteTemplate` - Update existing templates

#### Helper Functions (8)
- `Test-SPOSiteUrl` - Validate site URLs
- `Test-SPOSiteExists` - Check if site exists
- `Wait-SPOSiteCreation` - Wait for provisioning
- `Get-SPOProvisioningStatus` - Check provisioning status
- `Initialize-SPOSiteFeatures` - Activate site features
- `Write-SPOFactoryLog` - Write to module log
- `Invoke-SPOFactoryCommand` - Execute commands with retry
- `Get-SPOSiteStatus` - Get site status information

## 📊 Security Baselines

The module includes pre-configured security baselines:

### MSPStandard
- Balanced security for general use
- External sharing with restrictions
- 30-day anonymous link expiration
- Standard DLP policies

### MSPSecure
- Enhanced security for sensitive data
- Internal sharing only
- No anonymous links
- Strict DLP policies
- Conditional access enforcement

### MSPStrict
- Maximum security for regulated industries
- No external sharing
- Complete audit logging
- Data encryption at rest
- Advanced threat protection

## 🛠️ Configuration Files

### JSON Site Configuration
```json
{
  "client": "Contoso",
  "hubSite": {
    "title": "Contoso Hub",
    "url": "contoso-hub",
    "securityBaseline": "MSPSecure"
  },
  "sites": [
    {
      "title": "Finance Team",
      "url": "finance",
      "type": "TeamSite",
      "joinHub": true,
      "owners": ["cfo@contoso.com"],
      "members": ["finance-team@contoso.com"]
    }
  ]
}
```

### CSV Bulk Sites
```csv
Title,Url,Type,Description,SecurityBaseline,HubSite
"Project Alpha","project-alpha","TeamSite","Alpha team collaboration","MSPSecure","corp-hub"
"Project Beta","project-beta","TeamSite","Beta team collaboration","MSPSecure","corp-hub"
```

## 📈 Performance Guidelines

| Operation | Expected Time | Concurrent Limit |
|-----------|--------------|------------------|
| Single Site | 15-30 seconds | N/A |
| Hub Site | 20-40 seconds | N/A |
| Bulk (10 sites) | 2-5 minutes | 5 parallel |
| Bulk (50 sites) | 10-15 minutes | 10 parallel |
| Template Apply | 10-20 seconds | N/A |

## 🔍 Getting Help

### Documentation
- Function help: `Get-Help New-SPOSite -Full`
- Examples: `Get-Help New-SPOSite -Examples`
- Online help: `Get-Help New-SPOSite -Online`

### Support Resources
- [Common Issues](Troubleshooting/CommonIssues.md)
- [Error Codes](Troubleshooting/ErrorCodes.md)
- [Debugging Guide](Troubleshooting/Debugging.md)

### Module Information
```powershell
# Get module version
Get-Module SPOSiteFactory | Select-Object Version

# List all commands
Get-Command -Module SPOSiteFactory

# Get module configuration
Get-SPOFactoryConfig
```

## 📝 License and Credits

Developed for MSP environments managing multiple SharePoint Online tenants.

- **Version**: 1.0.0
- **PowerShell**: 5.1+
- **Dependencies**: PnP.PowerShell, PSFramework, Microsoft.PowerShell.SecretManagement
- **Author**: MSP Automation Team

---

*For the latest updates and additional phases, check the [Development Plan](../Plans/SPOSiteFactory-Development-Plan.md)*