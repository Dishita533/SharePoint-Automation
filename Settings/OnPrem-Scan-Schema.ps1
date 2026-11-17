#=========================================================
# OnPrem-Scan-Schema.ps1
# MUST be run under Windows PowerShell 5.1 (powershell.exe)
# Uses the legacy SharePointPnPPowerShell2019 module
# For every non-excluded, VISIBLE list/library on the given site, captures:
#   - Columns: Internal Name, Display Name, Type, Required (visible columns only)
#   - Content Types: names attached to the list
#   - Views: Title, the set of fields shown in each view (visible views only)
# Only Read access is needed - single site, no subsites.
#=========================================================
param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$UserName,
    [Parameter(Mandatory = $true)][string]$Password,
    [Parameter(Mandatory = $true)][string]$OutputJson
)

$ErrorActionPreference = "Stop"

# Lists/libraries to skip entirely - default/system content, not real migrated data
$ExcludeLists = @(
    "Apps","App Catalog","microfeed","MicroFeed",
    "Site Pages","Site Assets","Images","footer","Footer",
    "Calendar","Tasks","Workflow Tasks","Workflow History",
    "Important Links","Pages"
)

# Views to skip entirely regardless of list (e.g. "Threaded" on discussion boards)
$ExcludeViews = @("Threaded")

# SPO injects these internal records-management/compliance fields onto every list.
# Their Hidden flag is often $false in the API, but they never appear on the actual
# List Settings > Columns page - SharePoint's UI filters them by a special-case rule,
# not by Hidden alone. This list is scoped ONLY to that known, verified category -
# it is NOT a general "hide anything that looks like noise" list (that approach
# wrongly excluded real visible content types like "Document" before).
$ExcludeFields = @(
    "_ColorTag","_ComplianceFlags","_ComplianceTag","_ComplianceTagUserId",
    "_ComplianceTagWrittenTime","_DisplayName","_IsRecord","_CheckinComment",
    "_ModerationStatus","_ModerationComments","_EditMenuTableStart","_EditMenuTableStart2",
    "_EditMenuTableEnd","_ShortcutUrl","_ShortcutSiteId","_ShortcutWebId","_ShortcutUniqueId",
    "_SPSitePageFlags","_SourceUrl","_SharedFileIndex","_CommentCount","_LikeCount",
    "ComplianceAssetId"
)

# NOTE: beyond the precise list above, no other manual column/content-type exclusion.
# The List Settings page shown in the site UI is otherwise the source of truth for
# what's "visible" - driven by each field/content type's own properties, not by
# guessing names. A broad hand-maintained exclude list would risk dropping things
# (like the OOB "Document" content type) that genuinely appear on that page.

try {
    Import-Module SharePointPnPPowerShell2019 -ErrorAction Stop

    $SecurePwd = ConvertTo-SecureString $Password -AsPlainText -Force
    $Cred      = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePwd)

    Write-Host "[OnPrem] Connecting to $SiteUrl" -ForegroundColor Cyan
    $Conn = Connect-PnPOnline -Url $SiteUrl -Credentials $Cred -ReturnConnection -ErrorAction Stop

    $web = Get-PnPWeb -Connection $Conn
    Write-Host "[OnPrem] Connected: $($web.Title)" -ForegroundColor Green

    $Schema = @{}

    $AllLists = Get-PnPList -Connection $Conn -Includes Hidden, ItemCount
    Write-Host "[OnPrem] Total lists/libraries found (before any filtering): $($AllLists.Count)" -ForegroundColor Magenta

    $VisibleLists = $AllLists | Where-Object { -not $_.Hidden }
    Write-Host "[OnPrem] After Hidden filter: $($VisibleLists.Count)" -ForegroundColor Magenta

    $NonExcluded = $VisibleLists | Where-Object { $ExcludeLists -notcontains $_.Title }
    Write-Host "[OnPrem] After ExcludeLists filter: $($NonExcluded.Count)" -ForegroundColor Magenta

    $Lists = $NonExcluded | Where-Object { $_.ItemCount -gt 0 }
    Write-Host "[OnPrem] After zero-count filter: $($Lists.Count)" -ForegroundColor Magenta

    foreach ($List in $Lists) {
        Write-Host "[OnPrem] Scanning list: $($List.Title)" -ForegroundColor Cyan

        # --- Columns (visible only) ---
        $FieldsInfo = @{}
        $AllFields = Get-PnPField -List $List.Title -Connection $Conn

        # Defensive: PnP sometimes doesn't populate Hidden on fields by default
        # (same class of issue as Get-PnPList). Force-load it if it came back $null,
        # so the visible-only filter below can't silently pass hidden columns through.
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
            Write-Host "[OnPrem] Could not read content types for $($List.Title) - $($_.Exception.Message)" -ForegroundColor DarkGray
        }

        # --- Views (visible only) - existence + settings flags only, NOT the
        # columns inside each view. "Don't dig inside" per the site UI: only
        # View name, Show In, Default View, Mobile View, Default Mobile View. ---
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

        $Schema[$List.Title] = @{
            Fields       = $FieldsInfo
            ContentTypes = $ContentTypesInfo
            Views        = $ViewsInfo
        }
    }

    $Schema | ConvertTo-Json -Depth 6 | Out-File -FilePath $OutputJson -Encoding UTF8
    Write-Host "[OnPrem] Scan complete. $($Schema.Count) lists/libraries written to $OutputJson" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "[OnPrem] FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}