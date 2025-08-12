# Create Your First SharePoint Site

This guide walks you through creating your first SharePoint site using the SPOSiteFactory module.

## Prerequisites

- SPOSiteFactory module installed ([Installation Guide](Installation.md))
- SharePoint Administrator permissions
- Connected to SharePoint Online

## Step 1: Connect to SharePoint

```powershell
# Import the module
Import-Module SPOSiteFactory

# Connect to your SharePoint tenant
Connect-SPOFactory -TenantUrl "https://yourtenant-admin.sharepoint.com" `
                   -ClientName "YourOrg" `
                   -Interactive

# Verify connection
Test-SPOFactoryConnection -ClientName "YourOrg"
```

## Step 2: Create a Simple Team Site

### Basic Team Site
```powershell
# Create a team site with Microsoft 365 Group
$site = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/firstsite" `
                    -Title "My First Site" `
                    -Description "This is my first site created with SPOSiteFactory" `
                    -Owner "admin@yourtenant.com" `
                    -SiteType "TeamSite" `
                    -ClientName "YourOrg"

# Check the result
if ($site.Success) {
    Write-Host "Site created successfully!" -ForegroundColor Green
    Write-Host "URL: $($site.SiteUrl)"
    Write-Host "Creation time: $($site.CreationTime.TotalSeconds) seconds"
} else {
    Write-Host "Site creation failed: $($site.Errors -join ', ')" -ForegroundColor Red
}
```

### Team Site with Members
```powershell
# Create a team site with specific members and owners
$site = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/teamproject" `
                    -Title "Team Project Site" `
                    -Description "Collaborative space for our team" `
                    -Owner "projectmanager@yourtenant.com" `
                    -SiteType "TeamSite" `
                    -ClientName "YourOrg" `
                    -GroupMembers @(
                        "user1@yourtenant.com",
                        "user2@yourtenant.com",
                        "user3@yourtenant.com"
                    ) `
                    -GroupOwners @(
                        "manager2@yourtenant.com"
                    )
```

## Step 3: Create a Communication Site

```powershell
# Create a communication site (no M365 Group)
$commSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/companynews" `
                        -Title "Company News" `
                        -Description "Corporate communications and announcements" `
                        -Owner "communications@yourtenant.com" `
                        -SiteType "CommunicationSite" `
                        -ClientName "YourOrg"

Write-Host "Communication site created: $($commSite.SiteUrl)" -ForegroundColor Green
```

## Step 4: Apply Security Settings

### Create Site with Security Baseline
```powershell
# Create a secure team site
$secureSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/confidential" `
                          -Title "Confidential Projects" `
                          -Description "Secure collaboration space" `
                          -Owner "security@yourtenant.com" `
                          -SiteType "TeamSite" `
                          -ClientName "YourOrg" `
                          -SecurityBaseline "MSPSecure" `
                          -ConfigureOfficeFileHandling

# The MSPSecure baseline applies:
# - Internal sharing only
# - No anonymous links
# - Strict DLP policies
# - Office files open in desktop apps
```

### Apply Security to Existing Site
```powershell
# Apply security baseline to an existing site
Set-SPOSiteSecurityBaseline -SiteUrl "https://yourtenant.sharepoint.com/sites/existingsite" `
                            -BaselineName "MSPStandard" `
                            -ClientName "YourOrg" `
                            -ApplyToSite `
                            -ConfigureDocumentLibraries `
                            -EnableAuditing
```

## Step 5: Create Site with Hub Association

```powershell
# First, create a hub site
$hub = New-SPOHubSite -Title "Department Hub" `
                      -Url "dept-hub" `
                      -Description "Central hub for all department sites" `
                      -ClientName "YourOrg" `
                      -SecurityBaseline "MSPStandard"

# Then create a site and associate it with the hub
$teamSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/finance" `
                        -Title "Finance Team" `
                        -Owner "cfo@yourtenant.com" `
                        -SiteType "TeamSite" `
                        -ClientName "YourOrg" `
                        -HubSiteUrl $hub.Url

Write-Host "Site created and associated with hub: $($teamSite.HubAssociation.Status)" -ForegroundColor Green
```

## Step 6: Verify Your Site

### Check Site Status
```powershell
# Get site provisioning status
$status = Get-SPOProvisioningStatus -SiteUrl "https://yourtenant.sharepoint.com/sites/firstsite" `
                                    -ClientName "YourOrg"

Write-Host "Site Status: $($status.Status)"
Write-Host "Is Ready: $($status.IsReady)"
```

### Test Site Existence
```powershell
# Check if site exists
$exists = Test-SPOSiteExists -SiteUrl "https://yourtenant.sharepoint.com/sites/firstsite" `
                             -ClientName "YourOrg"

if ($exists) {
    Write-Host "Site exists and is accessible" -ForegroundColor Green
} else {
    Write-Host "Site does not exist or is not accessible" -ForegroundColor Yellow
}
```

## Common Patterns

### Pattern 1: Department Site with Standard Setup
```powershell
function New-DepartmentSite {
    param(
        [string]$DepartmentName,
        [string]$OwnerEmail,
        [string[]]$Members
    )
    
    $siteUrl = "https://yourtenant.sharepoint.com/sites/$($DepartmentName.ToLower() -replace ' ', '')"
    
    New-SPOSite -SiteUrl $siteUrl `
                -Title "$DepartmentName Department" `
                -Description "Collaboration space for $DepartmentName" `
                -Owner $OwnerEmail `
                -SiteType "TeamSite" `
                -ClientName "YourOrg" `
                -SecurityBaseline "MSPStandard" `
                -GroupMembers $Members `
                -ConfigureOfficeFileHandling `
                -EnableAuditing
}

# Use the function
New-DepartmentSite -DepartmentName "Marketing" `
                   -OwnerEmail "cmo@yourtenant.com" `
                   -Members @("marketer1@yourtenant.com", "marketer2@yourtenant.com")
```

### Pattern 2: Project Site with Expiration
```powershell
function New-ProjectSite {
    param(
        [string]$ProjectCode,
        [string]$ProjectName,
        [string]$ProjectManager,
        [datetime]$EndDate
    )
    
    $site = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/proj-$ProjectCode" `
                        -Title "Project: $ProjectName" `
                        -Description "Project $ProjectCode - Ends $($EndDate.ToString('yyyy-MM-dd'))" `
                        -Owner $ProjectManager `
                        -SiteType "TeamSite" `
                        -ClientName "YourOrg" `
                        -SecurityBaseline "MSPSecure"
    
    # Add expiration metadata (requires additional configuration)
    if ($site.Success) {
        Write-Host "Project site created for $ProjectName"
        # Additional logic for site expiration could go here
    }
    
    return $site
}
```

## Troubleshooting

### Site Creation Fails
```powershell
# Enable verbose logging
$VerbosePreference = "Continue"

# Try creating site with WhatIf first
New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/testsite" `
            -Title "Test Site" `
            -Owner "admin@yourtenant.com" `
            -SiteType "TeamSite" `
            -ClientName "YourOrg" `
            -WhatIf

# Check the logs
Get-SPOFactoryLog -ClientName "YourOrg" -Last 10
```

### URL Already Exists
```powershell
# Check if URL is available
$urlCheck = Test-SPOSiteUrl -Url "testsite" `
                            -ClientName "YourOrg" `
                            -CheckAvailability

if ($urlCheck.IsAvailable) {
    Write-Host "URL is available" -ForegroundColor Green
} else {
    Write-Host "URL is taken. Suggestions:" -ForegroundColor Yellow
    Write-Host "- testsite-$(Get-Date -Format 'yyyy')"
    Write-Host "- testsite-v2"
    Write-Host "- testsite-$([guid]::NewGuid().ToString('N').Substring(0,4))"
}
```

### Permission Issues
```powershell
# Verify your permissions
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive
$ctx = Get-PnPContext
Write-Host "Connected as: $($ctx.Web.CurrentUser.Email)"

# Check if you're a SharePoint admin
# This needs to be verified in M365 Admin Center
```

## Best Practices

1. **Always use descriptive names**: Sites should have clear, meaningful names
2. **Apply security baselines**: Never create sites without appropriate security
3. **Document ownership**: Always specify clear owners and members
4. **Use consistent naming**: Follow organizational naming conventions
5. **Test with WhatIf**: Always test complex operations with -WhatIf first

## Next Steps

- [Set Up Hub Architecture](HubArchitecture.md) - Create hub and spoke topology
- [Bulk Provisioning](BulkProvisioning.md) - Create multiple sites at once
- [Template Management](../Functions/TemplateManagement.md) - Create reusable templates

---

**Tip**: Save your commonly used site creation commands as functions or scripts for reuse!