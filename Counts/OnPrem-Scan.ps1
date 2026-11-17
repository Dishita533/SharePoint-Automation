#=========================================================
# OnPrem-Scan.ps1
# MUST be run under Windows PowerShell 5.1 (powershell.exe)
# Uses the legacy SharePointPnPPowerShell2019 module
# Single connection, single site (no subsites).
# Scans every VISIBLE list/library's item count - hidden ones excluded.
# Only Read access is needed.
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

    $Counts = @{}

    # -Includes Hidden is required: the legacy module doesn't load the Hidden
    # property by default, so without it $_.Hidden is always $null and the
    # "-not $_.Hidden" filter below would silently pass every list through,
    # including hidden ones that don't actually show on the Site Contents page.
    $Lists = Get-PnPList -Connection $Conn -Includes Hidden | Where-Object { -not $_.Hidden }

    foreach ($List in $Lists) {
        Write-Host "[OnPrem] Checking: $($List.Title)" -ForegroundColor Yellow
        $Counts[$List.Title] = [int]$List.ItemCount
    }

    $Counts | ConvertTo-Json | Out-File -FilePath $OutputJson -Encoding UTF8
    Write-Host "[OnPrem] Scan complete. $($Counts.Count) visible lists/libraries written to $OutputJson" -ForegroundColor Green
    exit 0
}
catch {
    Write-Host "[OnPrem] FATAL ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}