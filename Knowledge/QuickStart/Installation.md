# Installation Guide

## Prerequisites

### Required Software
- **PowerShell**: Version 5.1 or PowerShell 7+
- **Operating System**: Windows 10/11 or Windows Server 2016+
- **Network**: Internet connectivity to SharePoint Online

### Required Permissions
- SharePoint Administrator or Global Administrator role
- Application consent for PnP PowerShell (first-time setup)

## Step 1: Install Dependencies

### Install PnP PowerShell
```powershell
# Install latest PnP PowerShell module
Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser

# Verify installation
Get-Module -Name PnP.PowerShell -ListAvailable
```

### Install PSFramework
```powershell
# Install PSFramework for logging
Install-Module -Name PSFramework -Force -Scope CurrentUser

# Verify installation
Get-Module -Name PSFramework -ListAvailable
```

### Install SecretManagement
```powershell
# Install PowerShell SecretManagement
Install-Module -Name Microsoft.PowerShell.SecretManagement -Force -Scope CurrentUser
Install-Module -Name Microsoft.PowerShell.SecretStore -Force -Scope CurrentUser

# Configure SecretStore (first time only)
Set-SecretStoreConfiguration -Authentication None -Confirm:$false
```

## Step 2: Install SPOSiteFactory Module

### Option A: Install from Repository
```powershell
# Clone the repository
git clone https://github.com/yourorg/SPO-Prep.git
cd SPO-Prep

# Import the module
Import-Module .\SPOSiteFactory\SPOSiteFactory.psd1 -Force
```

### Option B: Install to PowerShell Modules
```powershell
# Copy to user modules directory
$modulePath = "$env:USERPROFILE\Documents\PowerShell\Modules\SPOSiteFactory"
Copy-Item -Path ".\SPOSiteFactory" -Destination $modulePath -Recurse -Force

# Import the module
Import-Module SPOSiteFactory -Force
```

## Step 3: Initial Configuration

### Configure Module Settings
```powershell
# Set default client name
Set-SPOFactoryConfig -DefaultClientName "YourOrganization"

# Set default security baseline
Set-SPOFactoryConfig -DefaultSecurityBaseline "MSPStandard"

# Configure logging
Set-SPOFactoryConfig -LogPath "C:\Logs\SPOSiteFactory" -LogLevel "Info"
```

### Store Credentials (Optional)
```powershell
# Store credentials securely
$cred = Get-Credential -Message "Enter SharePoint Admin credentials"
Set-Secret -Name "SPOAdmin" -Secret $cred -Vault "Microsoft.PowerShell.SecretStore"
```

## Step 4: Verify Installation

### Check Module Loading
```powershell
# Get module information
Get-Module SPOSiteFactory | Format-List

# List available commands
Get-Command -Module SPOSiteFactory | Group-Object Verb | Select-Object Count, Name, @{n='Commands';e={$_.Group.Name -join ', '}}
```

### Test Connection
```powershell
# Connect to SharePoint (Interactive)
Connect-SPOFactory -TenantUrl "https://yourtenant-admin.sharepoint.com" -ClientName "YourOrg" -Interactive

# Verify connection
Test-SPOFactoryConnection -ClientName "YourOrg"

# Disconnect when done
Disconnect-SPOFactory -ClientName "YourOrg"
```

## Step 5: Configure for Multi-Tenant (MSP)

### Set Up Multiple Tenant Profiles
```powershell
# Configure Client A
Connect-SPOFactory -TenantUrl "https://clienta-admin.sharepoint.com" `
                   -ClientName "ClientA" `
                   -StoredCredential "ClientA-SPOAdmin"

# Configure Client B
Connect-SPOFactory -TenantUrl "https://clientb-admin.sharepoint.com" `
                   -ClientName "ClientB" `
                   -StoredCredential "ClientB-SPOAdmin"
```

### Create Client-Specific Configurations
```powershell
# Set client-specific defaults
@{
    "ClientA" = @{
        SecurityBaseline = "MSPSecure"
        DefaultTimeZone = 10  # Eastern Time
        NamingPrefix = "CA"
    }
    "ClientB" = @{
        SecurityBaseline = "MSPStandard"
        DefaultTimeZone = 13  # Pacific Time
        NamingPrefix = "CB"
    }
} | ForEach-Object {
    $_.GetEnumerator() | ForEach-Object {
        Set-SPOFactoryConfig -ClientName $_.Key -Settings $_.Value
    }
}
```

## Troubleshooting Installation

### Common Issues

#### Module Import Fails
```powershell
# Check execution policy
Get-ExecutionPolicy -List

# Set appropriate execution policy
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### PnP PowerShell Connection Issues
```powershell
# Register PnP Management Shell application
Register-PnPManagementShellAccess

# Clear cached credentials
Disconnect-PnPOnline
Clear-PnPAzureADAccessToken
```

#### Permission Errors
```powershell
# Ensure you have SharePoint Admin role
# Check in Azure AD or M365 Admin Center

# For app-only authentication, register an app
$app = Register-PnPAzureADApp -ApplicationName "SPOSiteFactory" `
                               -Tenant "yourtenant.onmicrosoft.com" `
                               -SharePointApplicationPermissions "Sites.FullControl.All" `
                               -Interactive
```

### Verify All Components
```powershell
# Run installation verification script
$components = @{
    "PowerShell Version" = $PSVersionTable.PSVersion
    "PnP.PowerShell" = Get-Module -Name PnP.PowerShell -ListAvailable | Select-Object -First 1 -ExpandProperty Version
    "PSFramework" = Get-Module -Name PSFramework -ListAvailable | Select-Object -First 1 -ExpandProperty Version
    "SecretManagement" = Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable | Select-Object -First 1 -ExpandProperty Version
    "SPOSiteFactory" = Get-Module -Name SPOSiteFactory -ListAvailable | Select-Object -First 1 -ExpandProperty Version
}

$components | Format-Table -AutoSize
```

## Next Steps

1. [Create Your First Site](FirstSite.md) - Start provisioning SharePoint sites
2. [Hub Architecture Setup](HubArchitecture.md) - Set up hub and spoke topology
3. [Security Baselines](../Functions/SecurityBaselines.md) - Configure security settings

## Uninstallation

To remove the module:
```powershell
# Remove module from session
Remove-Module SPOSiteFactory -Force

# Delete module files
Remove-Item -Path "$env:USERPROFILE\Documents\PowerShell\Modules\SPOSiteFactory" -Recurse -Force

# Optional: Remove dependencies
Uninstall-Module -Name PnP.PowerShell -Force
Uninstall-Module -Name PSFramework -Force
Uninstall-Module -Name Microsoft.PowerShell.SecretManagement -Force
```

---

**Note**: Always test the module in a development environment before using in production.