#=========================================================
# Compare-Main-LHN.ps1
# MUST be run under PowerShell 7 (pwsh)
# Compares the QuickLaunch (Left Hand Navigation) between
# on-prem SP2019 and SPO, including nested nav items and all subsites.
# Uses modern PnP.PowerShell for SPO directly (this session),
# and delegates the on-prem scan to a Windows PowerShell 5.1
# child process (OnPrem-Scan-LHN.ps1, must sit in the same folder),
# since the two module lines cannot coexist in one PS session.
#=========================================================

#=========================
# Site URLs
#=========================
$OnPremSite = ""
$SPOSite    = ""

#=========================
# Azure AD App ClientId (used for SPO login)
#=========================
$ClientId = ""

#=========================
# On-Prem service account (passed to the child PS5.1 process)
#=========================
$OnPremUser     = ""
$OnPremPassword = ""     # plaintext per requirement - lock down this file's ACL / avoid committing it

#=========================
# OUTPUT LOCATION - set this to wherever you want the report saved
# This file is APPENDED to on every run, never overwritten.
#=========================
$OutputFolder = "C:\Dishita\Reports\SPCompare"
$OutputFile   = Join-Path $OutputFolder "LHN_Mismatch_Report.csv"
$OnPremJson   = Join-Path $OutputFolder "OnPremNav.json"
$OnPremLog    = Join-Path $OutputFolder "OnPremScan.log"
$OnPremScript = Join-Path $PSScriptRoot "OnPrem-Scan-LHN.ps1"
 
if (-not (Test-Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
 
if (-not (Test-Path $OnPremScript)) {
    Write-Host "OnPrem-Scan-LHN.ps1 not found at $OnPremScript" -ForegroundColor Red
    Write-Host "Make sure OnPrem-Scan-LHN.ps1 is saved in the same folder as this script." -ForegroundColor Red
    exit 1
}
if (Test-Path $OnPremJson) {
    Remove-Item $OnPremJson -Force   # clear any stale scan result from a previous run (JSON only, not the report CSV)
}
 
#=========================================================
# 1. On-prem LHN scan (Windows PowerShell 5.1 child process)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint 2019 LHN Scan (via Windows PowerShell 5.1)" -ForegroundColor Green
 
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $OnPremScript `
    -SiteUrl $OnPremSite `
    -UserName $OnPremUser `
    -Password $OnPremPassword `
    -OutputJson $OnPremJson 2>&1 | Tee-Object -FilePath $OnPremLog
 
if ($LASTEXITCODE -ne 0) {
    Write-Host "OnPrem scan process exited with code $LASTEXITCODE. Full log: $OnPremLog" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $OnPremJson)) {
    Write-Host "OnPrem LHN scan failed - no output file produced. Check log: $OnPremLog" -ForegroundColor Red
    exit 1
}
 
$OnPremRaw = Get-Content $OnPremJson -Raw | ConvertFrom-Json
$OnPremNav = [ordered]@{}
foreach ($prop in $OnPremRaw.PSObject.Properties) {
    $OnPremNav[$prop.Name] = [string]$prop.Value
}
Write-Host "OnPrem LHN scan returned $($OnPremNav.Count) visible nav items." -ForegroundColor Green
 
#=========================================================
# 2. SPO LHN scan (this PS7 session, modern PnP.PowerShell)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint Online LHN Scan (PS7 / PnP.PowerShell)" -ForegroundColor Green
 
Import-Module PnP.PowerShell -ErrorAction Stop
 
Write-Host "Please login to SharePoint Online..."
$Conn = Connect-PnPOnline -Url $SPOSite -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
 
$web = Get-PnPWeb -Connection $Conn
Write-Host "[SPO] Connected: $($web.Title)" -ForegroundColor Green
 
$SPONav = [ordered]@{}
 
# Converts a server-relative URL into a complete, clickable absolute link
# using the connected site's scheme+host. Leaves absolute URLs untouched.
function Get-AbsoluteUrl {
    param([string]$BaseUrl, [string]$NodeUrl)
    if ([string]::IsNullOrWhiteSpace($NodeUrl)) { return $NodeUrl }
    if ($NodeUrl -match '^https?://') { return $NodeUrl }
    $base = [Uri]$BaseUrl
    $root = "$($base.Scheme)://$($base.Host)"
    if (-not $NodeUrl.StartsWith('/')) { $NodeUrl = "/$NodeUrl" }
    return "$root$NodeUrl"
}
 
function Add-SPONavNodeToResults {
    param($Node, $PathPrefix)
 
    $NavPath = if ($PathPrefix) { "$PathPrefix>$($Node.Title)" } else { $Node.Title }
 
    if (-not $Node.IsVisible) {
        Write-Host "[SPO] Skipped (hidden, subtree excluded): $NavPath" -ForegroundColor DarkGray
        return
    }
 
    if ([string]::IsNullOrWhiteSpace($Node.Url)) {
        Write-Host "[SPO] Skipped (no target URL - effectively deleted): $NavPath" -ForegroundColor DarkGray
        return
    }
 
    if ($Node.Title -ieq "default.aspx") {
        Write-Host "[SPO] Skipped (default.aspx): $NavPath" -ForegroundColor DarkGray
        return
    }
 
    Write-Host "[SPO] Nav item: $NavPath" -ForegroundColor Yellow
    $SPONav[$NavPath] = Get-AbsoluteUrl -BaseUrl $SPOSite -NodeUrl $Node.Url
 
    if ($Node.Children -and $Node.Children.Count -gt 0) {
        foreach ($Child in $Node.Children) {
            Add-SPONavNodeToResults -Node $Child -PathPrefix $NavPath
        }
    }
}
 
$Nodes = Get-PnPNavigationNode -Location QuickLaunch -Connection $Conn
foreach ($Node in $Nodes) {
    Add-SPONavNodeToResults -Node $Node -PathPrefix ""
}
 
if ($SPONav.Count -eq 0) {
    Write-Host "SPO LHN scan returned no visible nav items - check connection/permissions/site URL. Aborting." -ForegroundColor Red
    exit 1
}
Write-Host "SPO LHN scan returned $($SPONav.Count) visible nav items." -ForegroundColor Green
 
#=========================================================
# 3. Merge OnPrem and SPO nav items by name - no comparison/status,
#    just both complete redirect links side by side for manual review.
#=========================================================
$AllRows = @()
 
# Preserve the order items actually appear in on the site - on-prem's nav
# order first (source of truth), then any SPO-only items appended after
# in the order SPO returned them. No alphabetical sorting.
$AllKeys = [System.Collections.Generic.List[string]]::new()
foreach ($k in $OnPremNav.Keys) { $AllKeys.Add($k) }
foreach ($k in $SPONav.Keys) { if (-not $AllKeys.Contains($k)) { $AllKeys.Add($k) } }
 
foreach ($Key in $AllKeys) {
    $AllRows += [PSCustomObject]@{
        NavItem   = $Key
        OnPremUrl = if ($OnPremNav.Contains($Key)) { $OnPremNav[$Key] } else { "N/A" }
        SPOUrl    = if ($SPONav.Contains($Key))    { $SPONav[$Key] }    else { "N/A" }
    }
}
 
#=========================================================
# 4. Display as table + Append to file - links only, no mismatch details
#=========================================================
Write-Host ""
Write-Host "===================================="
Write-Host "LHN (QuickLaunch) - Fetched Links"
Write-Host "===================================="
$AllRows | Format-Table -AutoSize -Property NavItem, OnPremUrl, SPOUrl | Out-Host
 
$RunTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
 
# -Width set high so columns never wrap - same fixed-width alignment as the console table
$TableText = ($AllRows | Format-Table -AutoSize -Property NavItem, OnPremUrl, SPOUrl | Out-String -Width 300).TrimEnd("`r","`n")
 
$Lines = @()
$Lines += "Run Date/Time : $RunTimestamp"
$Lines += "OnPrem Site   : $OnPremSite"
$Lines += "SPO Site      : $SPOSite"
$Lines += ""
$Lines += ($TableText -split "`r?`n")
$Lines += ""
$Lines += "Total Nav Items (OnPrem) : $($OnPremNav.Count)"
$Lines += "Total Nav Items (SPO)    : $($SPONav.Count)"
 
if (Test-Path $OutputFile) {
    Add-Content -Path $OutputFile -Value ""
    Add-Content -Path $OutputFile -Value ""
}
Add-Content -Path $OutputFile -Value $Lines
 
Write-Host ""
Write-Host "Report appended to: $OutputFile" -ForegroundColor Green
 