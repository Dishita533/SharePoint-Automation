#=========================================================
# Compare-Main.ps1
# MUST be run under PowerShell 7 (pwsh)
# Compares every VISIBLE list/library's item count between on-prem
# SP2019 and SPO, single site only (no subsites). Hidden lists are
# excluded on both sides - only what actually shows on Site Contents.
# Uses modern PnP.PowerShell for SPO directly, delegates the on-prem
# scan to a Windows PowerShell 5.1 child process (OnPrem-Scan.ps1,
# must sit in the same folder). Only Read access needed anywhere.
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
# OUTPUT LOCATION
#=========================
$OutputFolder = "C:\Dishita\Reports\SPCompare"
$OutputFile   = Join-Path $OutputFolder "Migration_Count_Report.csv"
$OnPremJson   = Join-Path $OutputFolder "OnPremCounts.json"
$OnPremLog    = Join-Path $OutputFolder "OnPremScan.log"
$OnPremScript = Join-Path $PSScriptRoot "OnPrem-Scan.ps1"

if (-not (Test-Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }

if (-not (Test-Path $OnPremScript)) {
    Write-Host "OnPrem-Scan.ps1 not found at $OnPremScript" -ForegroundColor Red
    Write-Host "Make sure OnPrem-Scan.ps1 is saved in the same folder as this script." -ForegroundColor Red
    exit 1
}
if (Test-Path $OnPremJson) { Remove-Item $OnPremJson -Force }

#=========================================================
# 1. On-prem scan (Windows PowerShell 5.1 child process)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint 2019 Scan (via Windows PowerShell 5.1)" -ForegroundColor Green

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
    Write-Host "OnPrem scan failed - no output file produced. Check log: $OnPremLog" -ForegroundColor Red
    exit 1
}

$OnPremRaw    = Get-Content $OnPremJson -Raw | ConvertFrom-Json
$OnPremCounts = @{}
foreach ($prop in $OnPremRaw.PSObject.Properties) {
    $OnPremCounts[$prop.Name] = [int]$prop.Value
}
Write-Host "OnPrem scan returned $($OnPremCounts.Count) visible lists/libraries." -ForegroundColor Green

#=========================================================
# 2. SPO scan (this PS7 session, modern PnP.PowerShell)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint Online Scan (PS7 / PnP.PowerShell)" -ForegroundColor Green

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Please login to SharePoint Online..."
$Conn = Connect-PnPOnline -Url $SPOSite -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop

$web = Get-PnPWeb -Connection $Conn
Write-Host "[SPO] Connected: $($web.Title)" -ForegroundColor Green

$SPOCounts = @{}
# Modern module loads Hidden by default, but -Includes Hidden is kept explicit
# for consistency/clarity with the on-prem side.
$Lists = Get-PnPList -Connection $Conn -Includes Hidden | Where-Object { -not $_.Hidden }

foreach ($List in $Lists) {
    Write-Host "[SPO] Checking: $($List.Title)" -ForegroundColor Yellow
    $SPOCounts[$List.Title] = [int]$List.ItemCount
}

if ($SPOCounts.Count -eq 0) {
    Write-Host "SPO scan returned no lists/libraries - check connection/permissions. Aborting." -ForegroundColor Red
    exit 1
}
Write-Host "SPO scan returned $($SPOCounts.Count) visible lists/libraries." -ForegroundColor Green

#=========================================================
# 3. Compare
#=========================================================
$Report = @()
$AllNames = ($OnPremCounts.Keys + $SPOCounts.Keys) | Sort-Object -Unique

foreach ($Name in $AllNames) {
    $InOnPrem = $OnPremCounts.ContainsKey($Name)
    $InSPO    = $SPOCounts.ContainsKey($Name)

    $OnPremCount = if ($InOnPrem) { $OnPremCounts[$Name] } else { $null }
    $SPOCount    = if ($InSPO)    { $SPOCounts[$Name] }    else { $null }

    if ($InOnPrem -and $InSPO) {
        $Status = if ($OnPremCount -eq $SPOCount) { "MATCH" } else { "COUNT MISMATCH" }
    }
    elseif ($InOnPrem -and -not $InSPO) { $Status = "MISSING IN SPO" }
    else { $Status = "EXTRA IN SPO" }

    $Difference = if ($InOnPrem -and $InSPO) { $SPOCount - $OnPremCount } else { "N/A" }

    $Report += [PSCustomObject]@{
        LibraryName = $Name
        OnPremCount = if ($InOnPrem) { $OnPremCount } else { "N/A" }
        SPOCount    = if ($InSPO)    { $SPOCount }    else { "N/A" }
        Difference  = $Difference
        Status      = $Status
    }
}

$SortedReport = $Report | Sort-Object Status, LibraryName

#=========================================================
# 4. Display as table (console)
#=========================================================
Write-Host ""
Write-Host "===================================="
Write-Host "Migration Validation - Visible Items Only"
Write-Host "===================================="
$SortedReport | Format-Table -AutoSize -Property LibraryName, OnPremCount, SPOCount, Difference, Status | Out-Host

$summary = $SortedReport | Group-Object Status | Select-Object Name, Count
Write-Host "------------------------------------"
Write-Host "Summary"
Write-Host "------------------------------------"
$summary | Format-Table -AutoSize | Out-Host

#=========================================================
# 5. Build report file: Date/Time + site links at top, then the table
#    formatted exactly like the terminal (fixed-width columns, no commas),
#    then overall counts separately. APPENDED each run, with a double
#    blank line between runs instead of overwriting.
#=========================================================
$RunTimestamp     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$TotalOnPremItems = ($SortedReport | Where-Object { $_.OnPremCount -ne "N/A" } | Measure-Object -Property OnPremCount -Sum).Sum
$TotalSPOItems    = ($SortedReport | Where-Object { $_.SPOCount -ne "N/A" } | Measure-Object -Property SPOCount -Sum).Sum

# -Width set high so columns never wrap - same fixed-width alignment as the console table
$TableText = ($SortedReport | Format-Table -AutoSize -Property LibraryName, OnPremCount, SPOCount, Difference, Status | Out-String -Width 300).TrimEnd("`r","`n")

$Lines = @()
$Lines += "Run Date/Time : $RunTimestamp"
$Lines += "OnPrem Site   : $OnPremSite"
$Lines += "SPO Site      : $SPOSite"
$Lines += ""
$Lines += ($TableText -split "`r?`n")
$Lines += ""
$Lines += "OVERALL COUNTS"
$Lines += "Total Lists/Libraries Compared : $($SortedReport.Count)"
$Lines += "Matched                        : $(($SortedReport | Where-Object Status -eq 'MATCH').Count)"
$Lines += "Count Mismatches               : $(($SortedReport | Where-Object Status -eq 'COUNT MISMATCH').Count)"
$Lines += "Missing in SPO                 : $(($SortedReport | Where-Object Status -eq 'MISSING IN SPO').Count)"
$Lines += "Extra in SPO                   : $(($SortedReport | Where-Object Status -eq 'EXTRA IN SPO').Count)"
$Lines += "Total OnPrem Item Count        : $TotalOnPremItems"
$Lines += "Total SPO Item Count           : $TotalSPOItems"

if (Test-Path $OutputFile) {
    Add-Content -Path $OutputFile -Value ""
    Add-Content -Path $OutputFile -Value ""
}
Add-Content -Path $OutputFile -Value $Lines

Write-Host ""
Write-Host "Report Generated: $OutputFile" -ForegroundColor Green