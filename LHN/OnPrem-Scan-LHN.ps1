#=========================================================
# OnPrem-Scan-LHN.ps1
# MUST be run under Windows PowerShell 5.1 (powershell.exe)
# Uses the legacy SharePointPnPPowerShell2019 module
# Scans the VISIBLE QuickLaunch (Left Hand Navigation) items on the
# given site only - no subsites. Only Read access is needed.
#=========================================================
param(
    [Parameter(Mandatory = $true)][string]$SiteUrl,
    [Parameter(Mandatory = $true)][string]$UserName,
    [Parameter(Mandatory = $true)][string]$Password,
    [Parameter(Mandatory = $true)][string]$OutputJson
)

$ErrorActionPreference = "Stop"

try {
    Import-Module SharePointPnPPowerShell2019 -ErrorAction Stop

    $SecurePwd = ConvertTo-SecureString $Password -AsPlainText -Force
    $Cred      = New-Object System.Management.Automation.PSCredential ($UserName, $SecurePwd)

    Write-Host "[OnPrem] Connecting to $SiteUrl" -ForegroundColor Cyan
    $Conn = Connect-PnPOnline -Url $SiteUrl -Credentials $Cred -ReturnConnection -ErrorAction Stop

    $web = Get-PnPWeb -Connection $Conn
    Write-Host "[OnPrem] Connected: $($web.Title)" -ForegroundColor Green

    $Results = [ordered]@{}

    # Converts a server-relative URL (e.g. "/sites/HR/SitePages/Home.aspx") into a
    # complete, clickable absolute link using the connected site's scheme+host.
    # Leaves already-absolute URLs (http/https) untouched.
    function Get-AbsoluteUrl {
        param([string]$BaseUrl, [string]$NodeUrl)
        if ([string]::IsNullOrWhiteSpace($NodeUrl)) { return $NodeUrl }
        if ($NodeUrl -match '^https?://') { return $NodeUrl }
        $base = [Uri]$BaseUrl
        $root = "$($base.Scheme)://$($base.Host)"
        if (-not $NodeUrl.StartsWith('/')) { $NodeUrl = "/$NodeUrl" }
        return "$root$NodeUrl"
    }

    # Recursively walks a nav node and its Children, keeping only visible items,
    # building a "Parent>Child>Grandchild" path for matching. Stops descending
    # into a hidden node's children entirely - a hidden parent hides its subtree.
    function Add-NavNodeToResults {
        param($Node, $PathPrefix)

        $NavPath = if ($PathPrefix) { "$PathPrefix>$($Node.Title)" } else { $Node.Title }

        if (-not $Node.IsVisible) {
            Write-Host "[OnPrem] Skipped (hidden, subtree excluded): $NavPath" -ForegroundColor DarkGray
            return
        }

        if ([string]::IsNullOrWhiteSpace($Node.Url)) {
            Write-Host "[OnPrem] Skipped (no target URL - effectively deleted): $NavPath" -ForegroundColor DarkGray
            return
        }

        if ($Node.Title -ieq "default.aspx") {
            Write-Host "[OnPrem] Skipped (default.aspx): $NavPath" -ForegroundColor DarkGray
            return
        }

        Write-Host "[OnPrem] Nav item: $NavPath" -ForegroundColor Yellow
        $Results[$NavPath] = Get-AbsoluteUrl -BaseUrl $SiteUrl -NodeUrl $Node.Url

        if ($Node.Children -and $Node.Children.Count -gt 0) {
            foreach ($Child in $Node.Children) {
                Add-NavNodeToResults -Node $Child -PathPrefix $NavPath
            }
        }
    }

    $Nodes = Get-PnPNavigationNode -Location QuickLaunch -Connection $Conn
    foreach ($Node in $Nodes) {
        Add-NavNodeToResults -Node $Node -PathPrefix ""
    }

    $Results | ConvertTo-Json | Out-File -FilePath $OutputJson -Encoding UTF8
    Write-Host "[OnPrem] Scan complete. $($Results.Count) visible nav items written to $OutputJson" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "[OnPrem] FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}