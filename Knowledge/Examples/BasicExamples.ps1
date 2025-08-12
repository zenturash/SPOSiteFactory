# SPOSiteFactory - Basic Examples
# These examples demonstrate common usage patterns for the SPOSiteFactory module

#region Module Setup
# Import the module
Import-Module SPOSiteFactory -Force

# Connect to SharePoint
$tenantUrl = "https://yourtenant-admin.sharepoint.com"
$clientName = "YourOrganization"

Connect-SPOFactory -TenantUrl $tenantUrl `
                   -ClientName $clientName `
                   -Interactive

# Verify connection
if (Test-SPOFactoryConnection -ClientName $clientName) {
    Write-Host "Connected successfully!" -ForegroundColor Green
}
#endregion

#region Example 1: Create a Simple Team Site
Write-Host "`n=== Example 1: Simple Team Site ===" -ForegroundColor Cyan

$teamSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/marketing" `
                        -Title "Marketing Team" `
                        -Description "Marketing department collaboration space" `
                        -Owner "marketingmanager@yourtenant.com" `
                        -SiteType "TeamSite" `
                        -ClientName $clientName

if ($teamSite.Success) {
    Write-Host "✓ Team site created: $($teamSite.SiteUrl)" -ForegroundColor Green
    Write-Host "  Creation time: $($teamSite.CreationTime.TotalSeconds) seconds"
}
#endregion

#region Example 2: Create a Communication Site
Write-Host "`n=== Example 2: Communication Site ===" -ForegroundColor Cyan

$commSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/companynews" `
                        -Title "Company News" `
                        -Description "Corporate announcements and news" `
                        -Owner "communications@yourtenant.com" `
                        -SiteType "CommunicationSite" `
                        -ClientName $clientName `
                        -SecurityBaseline "MSPStandard"

if ($commSite.Success) {
    Write-Host "✓ Communication site created: $($commSite.SiteUrl)" -ForegroundColor Green
    Write-Host "  Security baseline applied: MSPStandard"
}
#endregion

#region Example 3: Create a Hub Site
Write-Host "`n=== Example 3: Hub Site Creation ===" -ForegroundColor Cyan

$hubSite = New-SPOHubSite -Title "Department Hub" `
                          -Url "dept-hub" `
                          -Description "Central hub for all department sites" `
                          -ClientName $clientName `
                          -SecurityBaseline "MSPSecure" `
                          -Owners @("admin@yourtenant.com", "depthead@yourtenant.com")

if ($hubSite.Success) {
    Write-Host "✓ Hub site created: $($hubSite.Url)" -ForegroundColor Green
    Write-Host "  Hub ID: $($hubSite.HubSiteId)"
}
#endregion

#region Example 4: Create Site with M365 Group Members
Write-Host "`n=== Example 4: Team Site with Members ===" -ForegroundColor Cyan

$projectSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/projectalpha" `
                           -Title "Project Alpha" `
                           -Description "Project Alpha collaboration" `
                           -Owner "projectmanager@yourtenant.com" `
                           -SiteType "TeamSite" `
                           -ClientName $clientName `
                           -GroupMembers @(
                               "developer1@yourtenant.com",
                               "developer2@yourtenant.com",
                               "designer@yourtenant.com"
                           ) `
                           -GroupOwners @(
                               "teamlead@yourtenant.com"
                           ) `
                           -SecurityBaseline "MSPSecure"

if ($projectSite.Success) {
    Write-Host "✓ Project site created with team members" -ForegroundColor Green
    Write-Host "  M365 Group ID: $($projectSite.M365Group.Id)"
}
#endregion

#region Example 5: Associate Site with Hub
Write-Host "`n=== Example 5: Hub Association ===" -ForegroundColor Cyan

# Create a site and associate it with the hub
$financeSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/finance" `
                           -Title "Finance Department" `
                           -Owner "cfo@yourtenant.com" `
                           -SiteType "TeamSite" `
                           -ClientName $clientName `
                           -HubSiteUrl "https://yourtenant.sharepoint.com/sites/dept-hub"

if ($financeSite.Success) {
    Write-Host "✓ Finance site created and associated with hub" -ForegroundColor Green
    Write-Host "  Hub association status: $($financeSite.HubAssociation.Status)"
}

# Or associate existing site with hub
$association = Add-SPOSiteToHub -HubSiteUrl "https://yourtenant.sharepoint.com/sites/dept-hub" `
                               -SiteUrl "https://yourtenant.sharepoint.com/sites/hr" `
                               -ClientName $clientName `
                               -ApplyHubNavigation `
                               -ApplyHubTheme

if ($association.Success) {
    Write-Host "✓ HR site associated with hub" -ForegroundColor Green
}
#endregion

#region Example 6: Apply Security Baseline
Write-Host "`n=== Example 6: Security Baseline ===" -ForegroundColor Cyan

# Apply security baseline to existing site
$securityResult = Set-SPOSiteSecurityBaseline `
    -SiteUrl "https://yourtenant.sharepoint.com/sites/confidential" `
    -BaselineName "MSPStrict" `
    -ClientName $clientName `
    -ApplyToSite `
    -ConfigureDocumentLibraries `
    -EnableAuditing

if ($securityResult.Success) {
    Write-Host "✓ Security baseline applied successfully" -ForegroundColor Green
    Write-Host "  Applied settings: $($securityResult.AppliedSettings.Count)"
    Write-Host "  Failed settings: $($securityResult.FailedSettings.Count)"
}
#endregion

#region Example 7: Bulk Site Creation from Array
Write-Host "`n=== Example 7: Bulk Site Creation ===" -ForegroundColor Cyan

$bulkSites = @(
    @{
        Title = "Sales Team"
        Url = "sales"
        Type = "TeamSite"
        Owner = "salesmanager@yourtenant.com"
        Description = "Sales team collaboration"
        SecurityBaseline = "MSPStandard"
    },
    @{
        Title = "Support Team"
        Url = "support"
        Type = "TeamSite"
        Owner = "supportmanager@yourtenant.com"
        Description = "Customer support team"
        SecurityBaseline = "MSPStandard"
    },
    @{
        Title = "R&D Team"
        Url = "research"
        Type = "TeamSite"
        Owner = "rdmanager@yourtenant.com"
        Description = "Research and development"
        SecurityBaseline = "MSPSecure"
    }
)

$bulkResult = New-SPOBulkSites -Sites $bulkSites `
                               -ClientName $clientName `
                               -BatchSize 2 `
                               -ContinueOnError `
                               -GenerateReport

Write-Host "✓ Bulk creation completed" -ForegroundColor Green
Write-Host "  Total sites: $($bulkResult.TotalSites)"
Write-Host "  Successful: $($bulkResult.Successful)" -ForegroundColor Green
Write-Host "  Failed: $($bulkResult.Failed)" -ForegroundColor $(if ($bulkResult.Failed -gt 0) { 'Red' } else { 'Gray' })
Write-Host "  Duration: $($bulkResult.Duration.ToString('mm\:ss'))"
#endregion

#region Example 8: Create Site from Template
Write-Host "`n=== Example 8: Site from Template ===" -ForegroundColor Cyan

# Get available templates
$templates = Get-SPOSiteTemplate -Category "Project"
Write-Host "Available project templates:"
$templates | ForEach-Object { Write-Host "  - $($_.DisplayName)" }

# Create site using template (if template exists)
if ($templates) {
    $templateSite = New-SPOSite -SiteUrl "https://yourtenant.sharepoint.com/sites/project-beta" `
                                -Title "Project Beta" `
                                -Owner "pm@yourtenant.com" `
                                -SiteType "TeamSite" `
                                -ClientName $clientName `
                                -SiteDesignId $templates[0].Id
    
    if ($templateSite.Success) {
        Write-Host "✓ Site created from template: $($templates[0].DisplayName)" -ForegroundColor Green
    }
}
#endregion

#region Example 9: Test Site URL Availability
Write-Host "`n=== Example 9: URL Validation ===" -ForegroundColor Cyan

$urlsToTest = @("testsite1", "test site", "test_site", "test-site-2024")

foreach ($url in $urlsToTest) {
    $validation = Test-SPOSiteUrl -Url $url `
                                  -ClientName $clientName `
                                  -CheckAvailability
    
    if ($validation.IsValid -and $validation.IsAvailable) {
        Write-Host "✓ '$url' is valid and available" -ForegroundColor Green
    } elseif ($validation.IsValid) {
        Write-Host "⚠ '$url' is valid but already taken" -ForegroundColor Yellow
    } else {
        Write-Host "✗ '$url' is invalid: $($validation.ValidationErrors -join ', ')" -ForegroundColor Red
    }
}
#endregion

#region Example 10: Check Module Configuration
Write-Host "`n=== Example 10: Module Configuration ===" -ForegroundColor Cyan

$config = Get-SPOFactoryConfig

Write-Host "Current Configuration:"
Write-Host "  Default Client: $($config.DefaultClientName)"
Write-Host "  Default Security: $($config.DefaultSecurityBaseline)"
Write-Host "  Log Path: $($config.LogPath)"
Write-Host "  Log Level: $($config.LogLevel)"
Write-Host "  MSP Mode: $($config.MSPMode)"

# View recent logs
Write-Host "`nRecent Log Entries:"
Get-SPOFactoryLog -Last 5 | Format-Table TimeStamp, Level, Message -AutoSize
#endregion

#region Cleanup
Write-Host "`n=== Cleanup ===" -ForegroundColor Yellow

# Disconnect when done
Disconnect-SPOFactory -ClientName $clientName
Write-Host "Disconnected from SharePoint" -ForegroundColor Gray
#endregion

<#
.NOTES
    These examples demonstrate basic usage of the SPOSiteFactory module.
    Always test in a development environment before running in production.
    
    Key Points:
    - Always specify ClientName for multi-tenant isolation
    - Apply security baselines to all sites
    - Use WhatIf parameter to test operations
    - Generate reports for bulk operations
    - Check logs when troubleshooting
#>