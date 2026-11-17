#=========================================================
# Compare-Main-Schema.ps1
# MUST be run under PowerShell 7 (pwsh)
# Compares COLUMNS (name, type, required), CONTENT TYPES, and VIEWS
# across every non-excluded, VISIBLE list/library, between on-prem
# SP2019 and SPO. Single site only, no subsites, only Read access
# needed anywhere. Output file contains MISMATCHES ONLY and is
# APPENDED to on every run.
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
# OUTPUT LOCATION - appended to on every run, never overwritten
#=========================
$OutputFolder = "C:\Dishita\Reports\SPCompare"
$OutputFile   = Join-Path $OutputFolder "Schema_Mismatch_Report.csv"
$OnPremJson   = Join-Path $OutputFolder "OnPremSchema.json"
$OnPremLog    = Join-Path $OutputFolder "OnPremSchemaScan_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$OnPremScript = Join-Path $PSScriptRoot "OnPrem-Scan-Schema.ps1"
 
# Same exclusion list as on-prem, so both sides scan the same lists
$ExcludeLists = @(
    "Apps","App Catalog","microfeed","MicroFeed",
    "Site Pages","Site Assets","Images","footer","Footer",
    "Calendar","Tasks","Workflow Tasks","Workflow History",
    "Important Links","Pages"
)
$ExcludeViews = @("Threaded")
 
# SPO injects these internal records-management/compliance fields onto every list.
# Their Hidden flag is often $false in the API, but they never appear on the actual
# List Settings > Columns page - SharePoint's UI filters them by a special-case rule,
# not by Hidden alone. This list is scoped ONLY to that known, verified category.
$ExcludeFields = @(
    "_ColorTag","_ComplianceFlags","_ComplianceTag","_ComplianceTagUserId",
    "_ComplianceTagWrittenTime","_DisplayName","_IsRecord","_CheckinComment",
    "_ModerationStatus","_ModerationComments","_EditMenuTableStart","_EditMenuTableStart2",
    "_EditMenuTableEnd","_ShortcutUrl","_ShortcutSiteId","_ShortcutWebId","_ShortcutUniqueId",
    "_SPSitePageFlags","_SourceUrl","_SharedFileIndex","_CommentCount","_LikeCount",
    "ComplianceAssetId"
)
 
# NOTE: beyond the precise list above, no other manual column/content-type exclusion -
# everything else relies on each field/content type's own properties, matching what
# the List Settings page actually shows.
 
if (-not (Test-Path $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
 
if (-not (Test-Path $OnPremScript)) {
    Write-Host "OnPrem-Scan-Schema.ps1 not found at $OnPremScript" -ForegroundColor Red
    Write-Host "Make sure OnPrem-Scan-Schema.ps1 is saved in the same folder as this script." -ForegroundColor Red
    exit 1
}
if (Test-Path $OnPremJson) { Remove-Item $OnPremJson -Force }
 
#=========================================================
# 1. On-prem schema scan (Windows PowerShell 5.1 child process)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint 2019 Schema Scan (via Windows PowerShell 5.1)" -ForegroundColor Green
 
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $OnPremScript `
        -SiteUrl $OnPremSite `
        -UserName $OnPremUser `
        -Password $OnPremPassword `
        -OutputJson $OnPremJson 2>&1 | Tee-Object -FilePath $OnPremLog
}
catch {
    Write-Host "Could not write to log file ($OnPremLog) - $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host "Continuing without a log file - re-running the on-prem scan without Tee-Object..." -ForegroundColor Yellow
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $OnPremScript `
        -SiteUrl $OnPremSite `
        -UserName $OnPremUser `
        -Password $OnPremPassword `
        -OutputJson $OnPremJson
}
 
if ($LASTEXITCODE -ne 0) {
    Write-Host "OnPrem scan process exited with code $LASTEXITCODE. Full log: $OnPremLog" -ForegroundColor Red
    exit 1
}
if (-not (Test-Path $OnPremJson)) {
    Write-Host "OnPrem schema scan failed - no output file produced. Check log: $OnPremLog" -ForegroundColor Red
    exit 1
}
 
try {
    $OnPremRaw = Get-Content $OnPremJson -Raw | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Host "Failed to parse OnPrem scan JSON: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Raw file is at: $OnPremJson - inspect it directly to see what's malformed." -ForegroundColor Red
    exit 1
}
Write-Host "OnPrem schema scan returned $($OnPremRaw.PSObject.Properties.Count) lists/libraries." -ForegroundColor Green
 
#=========================================================
# 2. SPO schema scan (this PS7 session, modern PnP.PowerShell)
#=========================================================
Write-Host ""
Write-Host "Starting SharePoint Online Schema Scan (PS7 / PnP.PowerShell)" -ForegroundColor Green
 
Import-Module PnP.PowerShell -ErrorAction Stop
 
Write-Host "Please login to SharePoint Online..."
$Conn = Connect-PnPOnline -Url $SPOSite -Interactive -ClientId $ClientId -ReturnConnection -ErrorAction Stop
 
$web = Get-PnPWeb -Connection $Conn
Write-Host "[SPO] Connected: $($web.Title)" -ForegroundColor Green
 
$SPOSchema = @{}
$AllSPOLists = Get-PnPList -Connection $Conn -Includes Hidden, ItemCount
$VisibleSPOLists = $AllSPOLists | Where-Object {
    (-not $_.Hidden) -and ($ExcludeLists -notcontains $_.Title)
}
$Lists = $VisibleSPOLists | Where-Object { $_.ItemCount -gt 0 }
Write-Host "[SPO] Lists after Hidden/Exclude/zero-count filters: $($Lists.Count)" -ForegroundColor Magenta
 
foreach ($List in $Lists) {
    Write-Host "[SPO] Scanning list: $($List.Title)" -ForegroundColor Cyan
 
    # --- Columns (visible only) ---
    $FieldsInfo = @{}
    $AllFields = Get-PnPField -List $List.Title -Connection $Conn
 
    # Defensive: force-load Hidden if it came back $null, so the visible-only
    # filter below can't silently pass hidden columns through.
    foreach ($F in $AllFields) {
        if ($null -eq $F.Hidden) {
            try { Get-PnPProperty -ClientObject $F -Property Hidden -Connection $Conn | Out-Null }
            catch { }
        }
    }
 
    $Fields = $AllFields | Where-Object {
        (-not $_.Hidden) -and ($ExcludeFields -notcontains $_.InternalName) -and
        ($_.TypeAsString -ne "WorkflowStatus") -and
        $_.InternalName -and ($_.InternalName.Trim() -ne "")
    }
    foreach ($F in $Fields) {
        $FieldsInfo[$F.InternalName] = @{
            DisplayName = $F.Title
            Type        = $F.TypeAsString
            Required    = [bool]$F.Required
        }
    }
 
    # --- Content Types ---
    $ContentTypesInfo = @()
    try {
        $CTs = Get-PnPContentType -List $List.Title -Connection $Conn
        $ContentTypesInfo = @($CTs | ForEach-Object { $_.Name })
    }
    catch {
        Write-Host "[SPO] Could not read content types for $($List.Title) - $($_.Exception.Message)" -ForegroundColor DarkGray
    }
 
    # --- Views (visible only) - existence + settings flags only, not the
    # columns inside each view: View name, Show In, Default View, Mobile
    # View, Default Mobile View. ---
    $ViewsInfo = @{}
    $Views = Get-PnPView -List $List.Title -Connection $Conn | Where-Object {
        $_.Title -and ($_.Title.Trim() -ne "") -and (-not $_.Hidden) -and
        ($ExcludeViews -notcontains $_.Title)
    }
    foreach ($V in $Views) {
        $ViewsInfo[$V.Title] = @{
            ShowIn            = if ([bool]$V.PersonalView) { "Personal" } else { "All" }
            DefaultView       = [bool]$V.DefaultView
            MobileView        = [bool]$V.MobileView
            MobileDefaultView = [bool]$V.MobileDefaultView
        }
    }
 
    $SPOSchema[$List.Title] = @{
        Fields       = $FieldsInfo
        ContentTypes = $ContentTypesInfo
        Views        = $ViewsInfo
    }
}
 
if ($SPOSchema.Count -eq 0) {
    Write-Host "SPO schema scan found no lists/libraries after filtering (Hidden/Excluded/zero-count)." -ForegroundColor Yellow
    Write-Host "This can be legitimate on a small or newly-created site - continuing rather than aborting." -ForegroundColor Yellow
}
else {
    Write-Host "SPO schema scan returned $($SPOSchema.Count) lists/libraries." -ForegroundColor Green
}
 
#=========================================================
# 3. Compare - MISMATCHES ONLY
#=========================================================
$Mismatches = @()
 
# Convert on-prem JSON (PSCustomObject tree) into a plain hashtable of the same shape
$OnPremSchema = @{}
foreach ($listProp in $OnPremRaw.PSObject.Properties) {
    $listName = $listProp.Name
    $listVal  = $listProp.Value
    $fieldsHt = @{}
    if ($listVal.Fields) {
        foreach ($fp in $listVal.Fields.PSObject.Properties) {
            $fieldsHt[$fp.Name] = @{
                DisplayName = $fp.Value.DisplayName
                Type        = $fp.Value.Type
                Required    = [bool]$fp.Value.Required
            }
        }
    }
    $viewsHt = @{}
    if ($listVal.Views) {
        foreach ($vp in $listVal.Views.PSObject.Properties) {
            $viewsHt[$vp.Name] = @{
                ShowIn            = [string]$vp.Value.ShowIn
                DefaultView       = [bool]$vp.Value.DefaultView
                MobileView        = [bool]$vp.Value.MobileView
                MobileDefaultView = [bool]$vp.Value.MobileDefaultView
            }
        }
    }
    $ctList = @()
    if ($listVal.ContentTypes) { $ctList = @($listVal.ContentTypes) }
    $OnPremSchema[$listName] = @{ Fields = $fieldsHt; ContentTypes = $ctList; Views = $viewsHt }
}
 
$AllListNames = ($OnPremSchema.Keys + $SPOSchema.Keys) | Sort-Object -Unique
 
foreach ($ListName in $AllListNames) {
 
    $InOnPrem = $OnPremSchema.ContainsKey($ListName)
    $InSPO    = $SPOSchema.ContainsKey($ListName)
 
    if ($InOnPrem -and -not $InSPO) {
        $Mismatches += [PSCustomObject]@{
            ListName = $ListName; CheckType = "List"; ItemName = $ListName
            OnPrem = "Present"; SPO = "N/A"; Status = "MISSING IN SPO"
        }
        continue
    }
    if (-not $InOnPrem -and $InSPO) {
        $Mismatches += [PSCustomObject]@{
            ListName = $ListName; CheckType = "List"; ItemName = $ListName
            OnPrem = "N/A"; SPO = "Present"; Status = "EXTRA IN SPO"
        }
        continue
    }
 
    $OnPremList = $OnPremSchema[$ListName]
    $SPOList    = $SPOSchema[$ListName]
 
    #--- Columns ---
    $AllFieldNames = ($OnPremList.Fields.Keys + $SPOList.Fields.Keys) | Sort-Object -Unique
    foreach ($FieldName in $AllFieldNames) {
        $InOP = $OnPremList.Fields.ContainsKey($FieldName)
        $InSP = $SPOList.Fields.ContainsKey($FieldName)
 
        if ($InOP -and -not $InSP) {
            $f = $OnPremList.Fields[$FieldName]
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "Column"; ItemName = $FieldName
                OnPrem = "Type: $($f.Type), Required: $($f.Required)"; SPO = "N/A"; Status = "MISSING IN SPO"
            }
        }
        elseif (-not $InOP -and $InSP) {
            $f = $SPOList.Fields[$FieldName]
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "Column"; ItemName = $FieldName
                OnPrem = "N/A"; SPO = "Type: $($f.Type), Required: $($f.Required)"; Status = "EXTRA IN SPO"
            }
        }
        else {
            $fo = $OnPremList.Fields[$FieldName]
            $fs = $SPOList.Fields[$FieldName]
            if ($fo.Type -ne $fs.Type -or $fo.Required -ne $fs.Required) {
                $Mismatches += [PSCustomObject]@{
                    ListName = $ListName; CheckType = "Column"; ItemName = $FieldName
                    OnPrem = "Type: $($fo.Type), Required: $($fo.Required)"
                    SPO    = "Type: $($fs.Type), Required: $($fs.Required)"
                    Status = "COLUMN MISMATCH"
                }
            }
        }
    }
 
    #--- Content Types ---
    $AllCTNames = ($OnPremList.ContentTypes + $SPOList.ContentTypes) | Sort-Object -Unique
    foreach ($CTName in $AllCTNames) {
        $InOP = $OnPremList.ContentTypes -contains $CTName
        $InSP = $SPOList.ContentTypes -contains $CTName
 
        if ($InOP -and -not $InSP) {
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "ContentType"; ItemName = $CTName
                OnPrem = "Present"; SPO = "N/A"; Status = "MISSING IN SPO"
            }
        }
        elseif (-not $InOP -and $InSP) {
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "ContentType"; ItemName = $CTName
                OnPrem = "N/A"; SPO = "Present"; Status = "EXTRA IN SPO"
            }
        }
    }
 
    #--- Views (existence + settings flags only - not the columns inside each view) ---
    $AllViewNames = ($OnPremList.Views.Keys + $SPOList.Views.Keys) | Sort-Object -Unique
    foreach ($ViewName in $AllViewNames) {
        $InOP = $OnPremList.Views.ContainsKey($ViewName)
        $InSP = $SPOList.Views.ContainsKey($ViewName)
 
        if ($InOP -and -not $InSP) {
            $ov = $OnPremList.Views[$ViewName]
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "View"; ItemName = $ViewName
                OnPrem = "ShowIn: $($ov.ShowIn), Default: $($ov.DefaultView), Mobile: $($ov.MobileView), MobileDefault: $($ov.MobileDefaultView)"
                SPO = "N/A"; Status = "MISSING IN SPO"
            }
        }
        elseif (-not $InOP -and $InSP) {
            $sv = $SPOList.Views[$ViewName]
            $Mismatches += [PSCustomObject]@{
                ListName = $ListName; CheckType = "View"; ItemName = $ViewName
                OnPrem = "N/A"
                SPO = "ShowIn: $($sv.ShowIn), Default: $($sv.DefaultView), Mobile: $($sv.MobileView), MobileDefault: $($sv.MobileDefaultView)"
                Status = "EXTRA IN SPO"
            }
        }
        else {
            $ov = $OnPremList.Views[$ViewName]
            $sv = $SPOList.Views[$ViewName]
 
            $flagsDiffer = ($ov.ShowIn -ne $sv.ShowIn) -or
                           ($ov.DefaultView -ne $sv.DefaultView) -or
                           ($ov.MobileView -ne $sv.MobileView) -or
                           ($ov.MobileDefaultView -ne $sv.MobileDefaultView)
 
            if ($flagsDiffer) {
                $Mismatches += [PSCustomObject]@{
                    ListName = $ListName; CheckType = "View"; ItemName = $ViewName
                    OnPrem = "ShowIn: $($ov.ShowIn), Default: $($ov.DefaultView), Mobile: $($ov.MobileView), MobileDefault: $($ov.MobileDefaultView)"
                    SPO    = "ShowIn: $($sv.ShowIn), Default: $($sv.DefaultView), Mobile: $($sv.MobileView), MobileDefault: $($sv.MobileDefaultView)"
                    Status = "VIEW SETTINGS MISMATCH"
                }
            }
        }
    }
}
 
$Mismatches = $Mismatches | Sort-Object ListName, CheckType, ItemName
 
#=========================================================
# 4. Display as table + Append to file (separator line, no overwrite)
#=========================================================
Write-Host ""
Write-Host "===================================="
Write-Host "Schema Validation (Columns + Content Types + Views) - Mismatches Only"
Write-Host "===================================="
if ($Mismatches.Count -gt 0) {
    $Mismatches | Format-Table -AutoSize -Property ListName, CheckType, ItemName, OnPrem, SPO, Status | Out-Host
}
else {
    Write-Host "No mismatches found - all columns and views match." -ForegroundColor Green
}
 
if ($Mismatches.Count -gt 0) {
    $RunTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
 
    # -Width set high so columns never wrap - same fixed-width alignment as the console table
    $TableText = ($Mismatches | Format-Table -AutoSize -Property ListName, CheckType, ItemName, OnPrem, SPO, Status | Out-String -Width 300).TrimEnd("`r","`n")
 
    $Lines = @()
    $Lines += "Run Date/Time : $RunTimestamp"
    $Lines += "OnPrem Site   : $OnPremSite"
    $Lines += "SPO Site      : $SPOSite"
    $Lines += ""
    $Lines += ($TableText -split "`r?`n")
    $Lines += ""
    $Lines += "OVERALL COUNTS"
    $Lines += "Total Mismatches Found         : $($Mismatches.Count)"
    $Lines += "  List-level (Missing/Extra)    : $(($Mismatches | Where-Object CheckType -eq 'List').Count)"
    $Lines += "  Column Mismatches             : $(($Mismatches | Where-Object Status -eq 'COLUMN MISMATCH').Count)"
    $Lines += "  Columns Missing in SPO        : $(($Mismatches | Where-Object { $_.CheckType -eq 'Column' -and $_.Status -eq 'MISSING IN SPO' }).Count)"
    $Lines += "  Columns Extra in SPO          : $(($Mismatches | Where-Object { $_.CheckType -eq 'Column' -and $_.Status -eq 'EXTRA IN SPO' }).Count)"
    $Lines += "  Content Types Missing in SPO  : $(($Mismatches | Where-Object { $_.CheckType -eq 'ContentType' -and $_.Status -eq 'MISSING IN SPO' }).Count)"
    $Lines += "  Content Types Extra in SPO    : $(($Mismatches | Where-Object { $_.CheckType -eq 'ContentType' -and $_.Status -eq 'EXTRA IN SPO' }).Count)"
    $Lines += "  View Settings Mismatches      : $(($Mismatches | Where-Object Status -eq 'VIEW SETTINGS MISMATCH').Count)"
    $Lines += "  Views Missing in SPO          : $(($Mismatches | Where-Object { $_.CheckType -eq 'View' -and $_.Status -eq 'MISSING IN SPO' }).Count)"
    $Lines += "  Views Extra in SPO            : $(($Mismatches | Where-Object { $_.CheckType -eq 'View' -and $_.Status -eq 'EXTRA IN SPO' }).Count)"
 
    if (Test-Path $OutputFile) {
        Add-Content -Path $OutputFile -Value ""
        Add-Content -Path $OutputFile -Value ""
    }
    Add-Content -Path $OutputFile -Value $Lines
}
 
Write-Host ""
Write-Host "Report appended to: $OutputFile" -ForegroundColor Green
 