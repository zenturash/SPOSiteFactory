# Test script to check module functions
Import-Module ./SPOSiteFactory/SPOSiteFactory.psd1 -Force

Write-Host "`nChecking for SPOSiteFactory module..." -ForegroundColor Cyan
$module = Get-Module SPOSiteFactory
if ($module) {
    Write-Host "Module loaded: $($module.Name) v$($module.Version)" -ForegroundColor Green
} else {
    Write-Host "Module not loaded!" -ForegroundColor Red
    exit
}

Write-Host "`nChecking for Connect-SPOFactory function..." -ForegroundColor Cyan
$connectCmd = Get-Command Connect-SPOFactory -ErrorAction SilentlyContinue
if ($connectCmd) {
    Write-Host "Found: Connect-SPOFactory" -ForegroundColor Green
    Write-Host "  Type: $($connectCmd.CommandType)" -ForegroundColor Gray
    Write-Host "  Module: $($connectCmd.Module)" -ForegroundColor Gray
} else {
    Write-Host "Connect-SPOFactory not found!" -ForegroundColor Red
}

Write-Host "`nAll available commands in module:" -ForegroundColor Cyan
$commands = Get-Command -Module SPOSiteFactory
if ($commands) {
    $commands | Select-Object Name, CommandType | Sort-Object Name | Format-Table -AutoSize
} else {
    Write-Host "No commands found in module!" -ForegroundColor Red
}

Write-Host "`nChecking for Connection functions:" -ForegroundColor Cyan
$connectionFuncs = Get-Command -Module SPOSiteFactory -Name "*Connection*", "*Connect*", "*Disconnect*" -ErrorAction SilentlyContinue
if ($connectionFuncs) {
    $connectionFuncs | Select-Object Name | Format-Table -AutoSize
} else {
    Write-Host "No connection functions found!" -ForegroundColor Red
}