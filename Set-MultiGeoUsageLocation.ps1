#Requires -Version 5.1
<#
.SYNOPSIS
    Automates Microsoft 365 Multi-Geo usage location assignment and license provisioning.

.DESCRIPTION
    Analyzes all users in a Microsoft 365 tenant, detects the correct UsageLocation
    from OfficeLocation/City attributes, assigns the chosen Multi-Geo license where
    missing, and generates a full HTML dashboard + CSV report.

    Use -DryRun to preview all planned changes as a full HTML report without
    applying anything.

.PARAMETER Domain
    Tenant domain to filter users (e.g. contoso.com).

.PARAMETER SkuPartNumber
    SKU part number to assign. If omitted, an interactive picker lists all tenant SKUs.

.PARAMETER DefaultLocation
    Fallback ISO country code for users with no detectable region. Default: FR

.PARAMETER ExcludeUsers
    Array of UPNs to skip entirely.

.PARAMETER DryRun
    Preview mode: analyses users, builds full HTML report showing planned changes,
    but applies nothing to the tenant.

.PARAMETER OutputPath
    Folder for HTML / CSV / log output. Default: script directory.

.PARAMETER SkipLicenseAssignment
    Only update UsageLocation, do not assign any license.

.PARAMETER IncludeDisabledAccounts
    Include disabled accounts (skipped by default).

.EXAMPLE
    .\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -DryRun

.EXAMPLE
    .\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com"

.EXAMPLE
    .\Set-MultiGeoUsageLocation.ps1 -Domain "contoso.com" -SkuPartNumber "OFFICE365_MULTIGEO" -OutputPath "C:\Reports"

.NOTES
    Author  : Souhaiel Morhag (@msendpoint)
    Blog    : https://msendpoint.com
    Version : 2.4.0
    License : MIT
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$Domain,

    [Parameter()]
    [string]$SkuPartNumber = "",

    [Parameter()]
    [ValidateLength(2,2)]
    [string]$DefaultLocation = "FR",

    [Parameter()]
    [string[]]$ExcludeUsers = @(),

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [string]$OutputPath = $PSScriptRoot,

    [Parameter()]
    [switch]$SkipLicenseAssignment

)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =============================================================================
# REGION MAPPING ENGINE
# =============================================================================
$RegionRules = [ordered]@{
    "FR" = "paris|besancon|lyon|bordeaux|nantes|france"
    "CA" = "montr|toronto|vancouver|canada|quebec"
    "US" = "memphis|new york|chicago|boston|seattle|san francisco|usa|united states"
    "GB" = "london|manchester|birmingham|england|united kingdom"
    "DE" = "frankfurt|berlin|munich|hamburg|germany|allemagne"
    "SG" = "singapore|singapour"
    "CN" = "shanghai|beijing|shenzhen|china|chine"
    "JP" = "tokyo|osaka|japan|japon"
    "AU" = "sydney|melbourne|brisbane|australia|australie"
    "IN" = "mumbai|bangalore|delhi|india|inde"
    "BR" = "sao paulo|rio|brazil|bresil"
    "ES" = "madrid|barcelona|spain|espagne"
    "IT" = "rome|milan|italy|italie"
    "NL" = "amsterdam|rotterdam|netherlands|pays-bas"
}

# =============================================================================
# HELPERS
# =============================================================================
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    # Always write to log file — bypasses any inherited preference variables
    [System.IO.File]::AppendAllText($script:LogFile, ($line + [System.Environment]::NewLine))
    switch ($Level) {
        "SUCCESS" { Write-Host $line -ForegroundColor Green  }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "ERROR"   { Write-Host $line -ForegroundColor Red    }
        default   { Write-Host $line -ForegroundColor Cyan   }
    }
}

function Resolve-TargetLocation {
    param([string]$OfficeInfo, [string]$CurrentLocation)
    $lower = $OfficeInfo.ToLower().Trim()
    foreach ($code in $RegionRules.Keys) {
        if ($lower -match $RegionRules[$code]) { return $code }
    }
    if (-not [string]::IsNullOrWhiteSpace($CurrentLocation)) { return $CurrentLocation }
    return $script:DefaultLocation
}

function Ensure-GraphModules {
    $required = @(
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )
    foreach ($mod in $required) {
        if (-not (Get-Module -ListAvailable -Name $mod)) {
            Write-Log "Module '$mod' not found - installing..." "WARN"
            try {
                Install-Module $mod -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Module '$mod' installed." "SUCCESS"
            }
            catch {
                Write-Log "Failed to install '$mod': $_" "ERROR"
                exit 1
            }
        }
        if (-not (Get-Module -Name $mod)) {
            Write-Log "Importing module '$mod'..."
            Import-Module $mod -ErrorAction Stop
        }
        else {
            Write-Log "Module '$mod' already loaded."
        }
    }
}

function Build-HtmlReport {
    param(
        [System.Collections.Generic.List[PSCustomObject]]$ResultData,
        [string]$ReportPath,
        [string]$Domain,
        [string]$SkuLabel,
        [bool]$IsDryRun
    )

    $totalCount    = $ResultData.Count
    $runDate       = Get-Date -Format "dd/MM/yyyy HH:mm"
    $modeLabel     = if ($IsDryRun) { "DRY RUN - Preview only, nothing changed" } else { "LIVE EXECUTION" }
    $modeBanner    = if ($IsDryRun) {
        "<div class=`"dryrun-banner`">&#9888; DRY RUN MODE &mdash; This is a compliance preview. No changes have been applied to the tenant. Run without -DryRun to execute.</div>"
    } else { "" }

    if ($IsDryRun) {
        # DryRun: count what IS PLANNED
        $locCount     = @($ResultData | Where-Object { $_.NeedsLocFix -eq $true }).Count
        $licCount     = @($ResultData | Where-Object { $_.NeedsLicense -eq $true }).Count
        $disabledCount= @($ResultData | Where-Object { $_.IsDisabled -eq $true }).Count
        $errorCount   = 0
        $successCount = $totalCount
        $partialCount = 0
        $kpiWarnNum   = $totalCount
        $kpiWarnLabel = "Accounts to remediate"
        $kpiLocLabel  = "Location fixes needed"
        $kpiLicLabel  = "Licenses to assign"
        $kpiErrLabel  = "Disabled (location only)"
        $kpiErrNum    = $disabledCount
    }
    else {
        # Live: count what WAS ACTUALLY DONE
        $locCount     = @($ResultData | Where-Object { $_.LocationUpdated -eq $true }).Count
        $licCount     = @($ResultData | Where-Object { $_.LicenseAssigned -eq $true }).Count
        $successCount = @($ResultData | Where-Object { $_.Status -eq "Success" }).Count
        $partialCount = @($ResultData | Where-Object { $_.Status -eq "Partial" }).Count
        $errorCount   = @($ResultData | Where-Object { $_.Status -eq "Error" }).Count
        $kpiWarnNum   = $successCount + $partialCount
        $kpiWarnLabel = "Successfully processed"
        $kpiLocLabel  = "Locations updated"
        $kpiLicLabel  = "Licenses assigned"
        $kpiErrLabel  = "Errors"
        $kpiErrNum    = $errorCount
    }

    $cardList = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $ResultData) {

        # --- Card style ---
        if ($r.Status -eq "Preview") {
            $cardClass = "card preview"
        }
        elseif ($r.Status -eq "Success") {
            $cardClass = "card success-card"
        }
        elseif ($r.Status -eq "Partial") {
            $cardClass = "card partial"
        }
        else {
            $cardClass = "card error"
        }

        # --- Status badge ---
        if ($r.Status -eq "Preview") {
            $statusBadge = "<span class=`"badge preview`">&#9888; Needs action</span>"
        }
        elseif ($r.Status -eq "Success") {
            $statusBadge = "<span class=`"badge success`">&#10003; Success</span>"
        }
        elseif ($r.Status -eq "Partial") {
            $statusBadge = "<span class=`"badge partial`">&#9888; Partial</span>"
        }
        else {
            $statusBadge = "<span class=`"badge error`">&#10007; Error</span>"
        }

        # --- Disabled account badge ---
        $disabledBadge = if ($r.IsDisabled -eq $true) { "<span class=`"badge disabled`">&#128274; Disabled account</span>" } else { "" }

        # --- Action badges: DryRun = what IS needed, Live = what WAS actually done ---
        $actionBadges = ""
        if ($IsDryRun) {
            # Show what is planned
            if ($r.NeedsLocFix -eq $true) {
                $actionBadges += "<span class=`"badge loc`">&#8987; Location to fix</span> "
            }
            if ($r.NeedsLicense -eq $true) {
                $actionBadges += "<span class=`"badge lic-needed`">&#128273; License missing</span> "
            }
            if ($r.IsDisabled -eq $true) {
                $actionBadges += "<span class=`"badge disabled`">Location only (disabled)</span>"
            }
        }
        else {
            # Show what actually happened
            if ($r.LocationUpdated -eq $true) {
                $actionBadges += "<span class=`"badge loc-done`">&#10003; Location updated</span> "
            }
            elseif ($r.NeedsLocFix -eq $true) {
                $actionBadges += "<span class=`"badge loc`">&#10007; Location failed</span> "
            }
            if ($r.LicenseAssigned -eq $true) {
                $actionBadges += "<span class=`"badge license`">&#10003; License assigned</span>"
            }
            elseif ($r.NeedsLicense -eq $true) {
                $actionBadges += "<span class=`"badge lic-needed`">&#10007; License failed</span>"
            }
            elseif ($r.IsDisabled -eq $true -and $r.NeedsLocFix -eq $true) {
                $actionBadges += "<span class=`"badge disabled`">Location only (disabled)</span>"
            }
        }

        # --- Department: show N/A if looks like a person name (no dept data) ---
        $deptRaw = $r.Department
        $dept    = if ([string]::IsNullOrWhiteSpace($deptRaw)) { "<span class='na'>Not set in AAD</span>" } else { [System.Security.SecurityElement]::Escape($deptRaw) }

        # --- Office info ---
        $officeDisplay = if ($r.OfficeInfo -eq "(empty)") { "<span class='na'>Not set in AAD</span>" } else { [System.Security.SecurityElement]::Escape($r.OfficeInfo) }

        # --- Location row: only show arrow if location actually changes ---
        if ($r.PreviousLocation -ne $r.NewLocation) {
            $locDisplay = "<span class='loc-change'>$($r.PreviousLocation) &rarr; $($r.NewLocation)</span>"
        }
        else {
            $locDisplay = "<span class='loc-ok'>$($r.NewLocation) (unchanged)</span>"
        }

        # --- Message row ---
        $msgRow = if ($r.Message) { "<div class=`"row msg`"><i>$([System.Security.SecurityElement]::Escape($r.Message))</i></div>" } else { "" }

        # --- Compliance note in DryRun ---
        $complianceNote = ""
        if ($IsDryRun) {
            $notes = [System.Collections.Generic.List[string]]::new()
            if ($r.NeedsLicense -eq $true)  { $notes.Add("License <b>$SkuLabel</b> must be assigned before data residency is enforced.") }
            if ($r.NeedsLocFix -eq $true)   { $notes.Add("UsageLocation must be set to <b>$($r.NewLocation)</b> to enable geo-specific services.") }
            if ($notes.Count -gt 0) {
                $noteItems = ($notes | ForEach-Object { "<li>$_</li>" }) -join ""
                $complianceNote = "<div class=`"compliance-note`"><b>Required actions:</b><ul>$noteItems</ul></div>"
            }
        }

        $cardList.Add("
        <div class=`"$cardClass`">
            <div class=`"card-header`">$statusBadge $disabledBadge</div>
            <div class=`"action-badges`">$actionBadges</div>
            <h3>$([System.Security.SecurityElement]::Escape($r.DisplayName))</h3>
            <p class=`"upn`">$($r.UPN)</p>
            <div class=`"meta`">
                <div class=`"row`"><span>Department</span><b>$dept</b></div>
                <div class=`"row`"><span>Office</span><b>$officeDisplay</b></div>
                <div class=`"row`"><span>UsageLocation</span><b>$locDisplay</b></div>
                <div class=`"row`"><span>Analysed at</span><b>$($r.Timestamp)</b></div>
                $msgRow
            </div>
            $complianceNote
        </div>")
    }

    $allCards = $cardList -join "`n"

    $css = @"
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0f1117; color: #e1e4e8; min-height: 100vh; }
.dryrun-banner { background: #2d2200; color: #f0a500; text-align: center; padding: 16px 24px; font-weight: 700; font-size: 0.9rem; border-bottom: 2px solid #f0a500; letter-spacing: .03em; }
header { background: linear-gradient(135deg, #1a1f2e 0%, #0d1117 100%); border-bottom: 1px solid #30363d; padding: 32px 48px; }
header h1 { font-size: 1.8rem; font-weight: 700; color: #fff; }
header p { color: #8b949e; font-size: 0.9rem; margin-top: 6px; }
header .mode { display: inline-block; margin-top: 10px; font-size: 0.8rem; font-weight: 600; padding: 4px 12px; border-radius: 20px; background: #1f2937; border: 1px solid #30363d; color: #58a6ff; }
.kpi-bar { display: flex; gap: 16px; flex-wrap: wrap; padding: 24px 48px; background: #161b22; border-bottom: 1px solid #30363d; }
.kpi { background: #1f2937; border-radius: 12px; padding: 18px 24px; flex: 1; min-width: 140px; border: 1px solid #30363d; }
.kpi .num { font-size: 2.2rem; font-weight: 700; color: #58a6ff; }
.kpi .lbl { font-size: 0.75rem; color: #8b949e; text-transform: uppercase; letter-spacing: .05em; margin-top: 4px; line-height: 1.4; }
.kpi.total .num { color: #58a6ff; }
.kpi.warn  .num { color: #f0a500; }
.kpi.loc   .num { color: #79c0ff; }
.kpi.lic   .num { color: #d2a8ff; }
.kpi.ok    .num { color: #3fb950; }
.kpi.err   .num { color: #f85149; }
.grid { display: flex; flex-wrap: wrap; gap: 20px; padding: 32px 48px; }
.card { background: #1f2937; border-radius: 12px; width: 340px; padding: 22px; border: 1px solid #30363d; border-top: 4px solid #3fb950; transition: transform .15s, box-shadow .15s; }
.card:hover { transform: translateY(-3px); box-shadow: 0 8px 24px rgba(0,0,0,.4); }
.card.error { border-top-color: #f85149; }
.card.success-card { border-top-color: #3fb950; }
.card.preview { border-top-color: #f0a500; }
.card.partial { border-top-color: #f0a500; }
.card-header { margin-bottom: 8px; }
.action-badges { display: flex; gap: 6px; flex-wrap: wrap; margin-bottom: 12px; min-height: 24px; }
.card h3 { font-size: 1rem; font-weight: 600; color: #e1e4e8; margin-bottom: 2px; }
.upn { font-size: 0.76rem; color: #58a6ff; word-break: break-all; margin-bottom: 14px; }
.meta { border-top: 1px solid #30363d; padding-top: 10px; }
.row { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; font-size: 0.82rem; padding: 6px 0; border-bottom: 1px solid #21262d; }
.row:last-child { border-bottom: none; }
.row span { color: #6e7681; flex-shrink: 0; padding-top: 1px; }
.row b { color: #e1e4e8; text-align: right; word-break: break-word; font-weight: 500; }
.row.msg { display: block; color: #f85149; font-size: 0.78rem; padding: 6px 0; }
.na { color: #484f58; font-style: italic; font-weight: 400; }
.loc-change { color: #f0a500; font-weight: 600; }
.loc-ok { color: #6e7681; }
.compliance-note { background: #161b22; border: 1px solid #30363d; border-left: 3px solid #f0a500; border-radius: 6px; padding: 12px 14px; margin-top: 14px; font-size: 0.8rem; color: #8b949e; }
.compliance-note b { color: #e1e4e8; }
.compliance-note ul { margin: 8px 0 0 16px; }
.compliance-note li { margin-bottom: 4px; line-height: 1.5; }
.badge { font-size: 0.71rem; font-weight: 600; padding: 3px 8px; border-radius: 20px; white-space: nowrap; }
.badge.success  { background: #1a4731; color: #3fb950; border: 1px solid #3fb950; }
.badge.error    { background: #3d1a1a; color: #f85149; border: 1px solid #f85149; }
.badge.license  { background: #2d1f42; color: #d2a8ff; border: 1px solid #d2a8ff; }
.badge.preview  { background: #2d1e00; color: #f0a500; border: 1px solid #f0a500; }
.badge.loc      { background: #0d2137; color: #79c0ff; border: 1px solid #388bfd; }
.badge.loc-done { background: #0a2a1a; color: #56d364; border: 1px solid #2ea043; }
.badge.lic-needed { background: #2a1a3a; color: #d2a8ff; border: 1px solid #8957e5; }
.badge.partial  { background: #2d2200; color: #f0a500; border: 1px solid #f0a500; }
.badge.disabled { background: #1a1f2e; color: #6e7681; border: 1px solid #30363d; }
footer { text-align: center; padding: 28px; color: #484f58; font-size: 0.8rem; border-top: 1px solid #30363d; margin-top: 20px; }
footer a { color: #58a6ff; text-decoration: none; }
"@

    $html = "<!DOCTYPE html>`n<html lang=`"en`">`n<head>`n<meta charset=`"UTF-8`">`n" +
            "<meta name=`"viewport`" content=`"width=device-width, initial-scale=1.0`">`n" +
            "<title>Multi-Geo Compliance Report - $Domain</title>`n<style>`n$css`n</style>`n</head>`n<body>`n" +
            "$modeBanner`n" +
            "<header>`n  <h1>&#127760; Multi-Geo Compliance Report</h1>`n" +
            "  <p>Domain: <b>$Domain</b> &nbsp;|&nbsp; Generated: $runDate &nbsp;|&nbsp; SKU: $SkuLabel</p>`n" +
            "  <span class=`"mode`">$modeLabel</span>`n</header>`n" +
            "<div class=`"kpi-bar`">`n" +
            "  <div class=`"kpi warn`"><div class=`"num`">$kpiWarnNum</div><div class=`"lbl`">$kpiWarnLabel</div></div>`n" +
            "  <div class=`"kpi loc`"><div class=`"num`">$locCount</div><div class=`"lbl`">$kpiLocLabel</div></div>`n" +
            "  <div class=`"kpi lic`"><div class=`"num`">$licCount</div><div class=`"lbl`">$kpiLicLabel</div></div>`n" +
            "  <div class=`"kpi err`"><div class=`"num`">$kpiErrNum</div><div class=`"lbl`">$kpiErrLabel</div></div>`n" +
            "</div>`n<div class=`"grid`">`n$allCards`n</div>`n" +
            "<footer>Generated by <a href=`"https://msendpoint.com`" target=`"_blank`">msendpoint.com</a> &mdash; M365 Multi-Geo Manager v2.4</footer>`n" +
            "</body>`n</html>"

    [System.IO.File]::WriteAllText($ReportPath, $html, [System.Text.Encoding]::UTF8)
}

# =============================================================================
# INIT
# =============================================================================
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = $PSScriptRoot }
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$RunId          = Get-Date -Format "yyyyMMdd_HHmmss"
$script:LogFile = Join-Path $OutputPath "MultiGeo_$RunId.log"
$HtmlReport     = Join-Path $OutputPath "MultiGeo_Report_$RunId.html"
$CsvReport      = Join-Path $OutputPath "MultiGeo_Report_$RunId.csv"

# Create log file immediately
[System.IO.File]::WriteAllText($script:LogFile, "")

Write-Log "============================================================"
Write-Log "  M365 Multi-Geo Manager v2.4  |  msendpoint.com"
Write-Log "============================================================"
Write-Log "Domain         : $Domain"
Write-Log "Default Region : $DefaultLocation"
Write-Log "Mode           : $(if ($DryRun) { 'DRY RUN (no changes will be made)' } else { 'LIVE EXECUTION' })"
Write-Log "Output Path    : $OutputPath"
Write-Log "HTML Report    : $HtmlReport"
Write-Log "CSV Export     : $CsvReport"
Write-Log "============================================================"

# =============================================================================
# AUTHENTICATION
# =============================================================================
Ensure-GraphModules

try {
    $existingCtx = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -ne $existingCtx -and $existingCtx.Account) {
        Write-Log "Existing Graph session ($($existingCtx.Account)) - disconnecting first..." "WARN"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
catch { }

Write-Log "Opening Microsoft Graph authentication - sign in with a Global Admin account..."
$scopes = @("User.ReadWrite.All", "Organization.Read.All")

try {
    Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
    $ctx = Get-MgContext
    Write-Log "Authenticated as : $($ctx.Account)" "SUCCESS"
    Write-Log "Tenant ID        : $($ctx.TenantId)" "SUCCESS"
}
catch {
    Write-Log "Failed to connect to Microsoft Graph: $_" "ERROR"
    exit 1
}

# =============================================================================
# LICENSE DISCOVERY - Interactive SKU picker
# =============================================================================
$targetSku = $null

if (-not $SkipLicenseAssignment) {

    Write-Log "Fetching all subscribed SKUs from tenant..."
    $allSkus = Get-MgSubscribedSku -All | Sort-Object SkuPartNumber

    if (-not $allSkus) {
        Write-Log "No subscribed SKUs found in this tenant." "ERROR"
        exit 1
    }

    # -- MODE 1: SKU passed as parameter (CI / automation) --------------------
    if (-not [string]::IsNullOrWhiteSpace($SkuPartNumber)) {

        Write-Log "SKU parameter provided: '$SkuPartNumber'"
        $targetSku = $allSkus | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }

        if (-not $targetSku) {
            Write-Log "SKU '$SkuPartNumber' not found in this tenant." "ERROR"
            Write-Log "Available SKUs:" "WARN"
            foreach ($s in $allSkus) {
                $avail = $s.PrepaidUnits.Enabled - $s.ConsumedUnits
                Write-Log ("  {0,-50} Total:{1,6}  Used:{2,6}  Available:{3,6}" -f $s.SkuPartNumber, $s.PrepaidUnits.Enabled, $s.ConsumedUnits, $avail) "WARN"
            }
            exit 1
        }
    }
    # -- MODE 2: Interactive picker -------------------------------------------
    else {
        Write-Host ""
        Write-Host "========================================================" -ForegroundColor Cyan
        Write-Host "  SELECT THE LICENSE SKU TO ASSIGN" -ForegroundColor Cyan
        Write-Host "========================================================" -ForegroundColor Cyan
        Write-Host ""

        $skuTable = @()
        $idx = 1
        foreach ($s in $allSkus) {
            $avail      = $s.PrepaidUnits.Enabled - $s.ConsumedUnits
            $availColor = if ($avail -gt 0) { "Green" } else { "Red" }
            $line = "  [{0,2}]  {1,-50}  Total:{2,6}  Used:{3,6}" -f $idx, $s.SkuPartNumber, $s.PrepaidUnits.Enabled, $s.ConsumedUnits
            Write-Host $line -NoNewline
            Write-Host ("  Available:{0,6}" -f $avail) -ForegroundColor $availColor
            $skuTable += [PSCustomObject]@{ Index = $idx; Sku = $s }
            $idx++
        }

        Write-Host ""
        Write-Host "  [ 0]  Skip license assignment (update UsageLocation only)" -ForegroundColor Yellow
        Write-Host ""

        $pickInt = -1
        do {
            $pick = Read-Host "  Enter the number of the SKU to assign"
            $tmp  = 0
            $ok   = [int]::TryParse($pick.Trim(), [ref]$tmp)
            if ($ok -and $tmp -ge 0 -and $tmp -le $skuTable.Count) {
                $pickInt = $tmp
            }
            else {
                Write-Host ("  Invalid - enter a number between 0 and {0}." -f $skuTable.Count) -ForegroundColor Red
            }
        } while ($pickInt -lt 0)

        if ($pickInt -eq 0) {
            Write-Log "Skipping license assignment - UsageLocation only." "WARN"
            $SkipLicenseAssignment = $true
        }
        else {
            $targetSku = $skuTable[$pickInt - 1].Sku
            Write-Log "Selected SKU: $($targetSku.SkuPartNumber)" "SUCCESS"
        }
    }

    if ($null -ne $targetSku) {
        $remaining = $targetSku.PrepaidUnits.Enabled - $targetSku.ConsumedUnits
        Write-Log "SKU         : $($targetSku.SkuPartNumber)" "SUCCESS"
        Write-Log "Total seats : $($targetSku.PrepaidUnits.Enabled)"
        Write-Log "Consumed    : $($targetSku.ConsumedUnits)"
        if ($remaining -gt 0) {
            Write-Log "Available   : $remaining seats" "SUCCESS"
        }
        else {
            Write-Log "Available   : $remaining - WARNING no seats left, assignments may fail." "WARN"
            if ([Environment]::UserInteractive) {
                $cont = Read-Host "Continue anyway? (Y/N)"
                if ($cont.ToUpper() -ne "Y") { exit 0 }
            }
        }
    }
}

# =============================================================================
# USER DISCOVERY & ANALYSIS
# =============================================================================
Write-Log "Fetching users for domain '$Domain'..."

# Fetch ALL accounts - disabled accounts are always included for location-only updates
$allUsers = Get-MgUser -All `
    -Property "Id,DisplayName,UserPrincipalName,UsageLocation,AssignedLicenses,AccountEnabled,OfficeLocation,City,Department" |
    Where-Object { $_.UserPrincipalName -like "*@$Domain" }

Write-Log "Total users retrieved: $($allUsers.Count)"

if ($ExcludeUsers.Count -gt 0) {
    $allUsers = $allUsers | Where-Object { $ExcludeUsers -notcontains $_.UserPrincipalName }
    Write-Log "After exclusion list: $($allUsers.Count) users remaining"
}

$toProcess = [System.Collections.Generic.List[PSCustomObject]]::new()
$i = 0

foreach ($user in $allUsers) {
    $i++
    Write-Progress -Activity "Analysing users" -Status $user.UserPrincipalName -PercentComplete (($i / $allUsers.Count) * 100)

    $officeInfo  = "$($user.OfficeLocation) $($user.City)".Trim()
    $targetLoc   = Resolve-TargetLocation -OfficeInfo $officeInfo -CurrentLocation $user.UsageLocation
    $needsLocFix = ($user.UsageLocation -ne $targetLoc)

    # Safe SkuId check - handles users with no licenses (empty collection under StrictMode)
    $hasLicense = $true
    if ($null -ne $targetSku) {
        $assignedIds = @($user.AssignedLicenses | Select-Object -ExpandProperty SkuId -ErrorAction SilentlyContinue)
        $hasLicense  = $assignedIds -contains $targetSku.SkuId
    }

    # Disabled accounts: location fix only - NEVER assign license (covers Teams Phone resource accounts)
    $isDisabled   = ($user.AccountEnabled -eq $false)
    $needsLicense = (-not $hasLicense) -and (-not $SkipLicenseAssignment) -and (-not $isDisabled)

    # Skip disabled accounts that don't need a location fix (no reason to process them)
    if ($isDisabled -and -not $needsLocFix) { continue }

    if ($needsLocFix -or $needsLicense) {
        $actionParts = [System.Collections.Generic.List[string]]::new()
        if ($needsLocFix)  { $actionParts.Add("Set location: $($user.UsageLocation) -> $targetLoc") }
        if ($needsLicense) { $actionParts.Add("Add license: $SkuPartNumber") }
        if ($isDisabled)   { $actionParts.Add("Disabled account - location only") }

        $toProcess.Add([PSCustomObject]@{
            Id              = $user.Id
            UPN             = $user.UserPrincipalName
            DisplayName     = $user.DisplayName
            Department      = $user.Department
            OfficeInfo      = if ($officeInfo) { $officeInfo } else { "(empty)" }
            IsDisabled      = $isDisabled
            CurrentLocation = if ($user.UsageLocation) { $user.UsageLocation } else { "N/A" }
            TargetLocation  = $targetLoc
            HasLicense      = $hasLicense
            Actions         = ($actionParts -join " | ")
            NeedsLocFix     = $needsLocFix
            NeedsLicense    = $needsLicense
        })
    }
}

Write-Progress -Completed -Activity "Analysing users"
Write-Log "Analysis complete: $($toProcess.Count) accounts need action out of $($allUsers.Count) total."

if ($toProcess.Count -eq 0) {
    Write-Log "All accounts are already compliant. No changes needed." "SUCCESS"
    Disconnect-MgGraph | Out-Null
    exit 0
}

# Show summary table in console
Write-Host ""
$toProcess | Format-Table UPN, OfficeInfo, CurrentLocation, TargetLocation, Actions -AutoSize
Write-Host ""

# =============================================================================
# DRY RUN MODE - Generate preview report, no changes applied
# =============================================================================
if ($DryRun) {
    Write-Log "DRY RUN: Building preview report for $($toProcess.Count) planned actions..." "WARN"

    $previewResults = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($item in $toProcess) {
        $previewResults.Add([PSCustomObject]@{
            UPN              = $item.UPN
            DisplayName      = $item.DisplayName
            Department       = $item.Department
            OfficeInfo       = $item.OfficeInfo
            IsDisabled       = $item.IsDisabled
            PreviousLocation = $item.CurrentLocation
            NewLocation      = $item.TargetLocation
            NeedsLocFix      = $item.NeedsLocFix
            NeedsLicense     = $item.NeedsLicense
            LocationUpdated  = $false
            LicenseAssigned  = $false
            Actions          = $item.Actions
            Status           = "Preview"
            Message          = ""
            Timestamp        = (Get-Date -Format "HH:mm:ss")
        })
    }

    $previewResults | Export-Csv -Path $CsvReport -NoTypeInformation -Encoding UTF8
    Write-Log "Preview CSV: $CsvReport" "SUCCESS"

    $skuLabel = if ($targetSku) { $targetSku.SkuPartNumber } else { "N/A (location only)" }
    Build-HtmlReport -ResultData $previewResults -ReportPath $HtmlReport -Domain $Domain -SkuLabel $skuLabel -IsDryRun $true

    Write-Log "Preview HTML report: $HtmlReport" "SUCCESS"
    Write-Log "DRY RUN complete. $($toProcess.Count) changes would be applied. Run without -DryRun to execute." "WARN"

    Disconnect-MgGraph | Out-Null
    Invoke-Item $HtmlReport
    exit 0
}

# =============================================================================
# LIVE EXECUTION - Interactive confirmation
# =============================================================================
if ([Environment]::UserInteractive) {
    $confirm = Read-Host "Apply changes to $($toProcess.Count) accounts? (Y/N)"
    if ($confirm.ToUpper() -ne "Y") {
        Write-Log "Cancelled by user." "WARN"
        Disconnect-MgGraph | Out-Null
        exit 0
    }
}

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$i = 0

foreach ($item in $toProcess) {
    $i++
    Write-Progress -Activity "Applying changes" -Status $item.UPN -PercentComplete (($i / $toProcess.Count) * 100)

    # Track each operation independently — a license failure must not hide a location success
    $locUpdated  = $false
    $licAssigned = $false
    $errors      = [System.Collections.Generic.List[string]]::new()
    $messages    = [System.Collections.Generic.List[string]]::new()

    # --- Location update (independent try/catch) ---
    if ($item.NeedsLocFix) {
        try {
            Update-MgUser -UserId $item.Id -UsageLocation $item.TargetLocation
            $locUpdated = $true
            $messages.Add("Location set to $($item.TargetLocation)")
            Write-Log "$($item.UPN) - Location: $($item.CurrentLocation) -> $($item.TargetLocation)" "SUCCESS"
        }
        catch {
            $errors.Add("Location error: $($_.Exception.Message)")
            Write-Log "$($item.UPN) - Location ERROR: $($_.Exception.Message)" "ERROR"
        }
    }

    # --- License assignment (independent try/catch, never for disabled accounts) ---
    if ($item.NeedsLicense -and ($null -ne $targetSku) -and (-not $item.IsDisabled)) {
        try {
            Set-MgUserLicense -UserId $item.Id -AddLicenses @{ SkuId = $targetSku.SkuId } -RemoveLicenses @()
            $licAssigned = $true
            $messages.Add("License $($targetSku.SkuPartNumber) assigned")
            Write-Log "$($item.UPN) - License assigned: $($targetSku.SkuPartNumber)" "SUCCESS"
        }
        catch {
            $errors.Add("License error: $($_.Exception.Message)")
            Write-Log "$($item.UPN) - License ERROR: $($_.Exception.Message)" "ERROR"
        }
    }

    # --- Determine overall status ---
    if ($errors.Count -eq 0) {
        $status = "Success"
    }
    elseif ($locUpdated -or $licAssigned) {
        $status = "Partial"   # at least one operation succeeded
    }
    else {
        $status = "Error"
    }

    $combinedMessage = ($messages + $errors) -join " | "

    $results.Add([PSCustomObject]@{
        UPN              = $item.UPN
        DisplayName      = $item.DisplayName
        Department       = $item.Department
        OfficeInfo       = $item.OfficeInfo
        IsDisabled       = $item.IsDisabled
        PreviousLocation = $item.CurrentLocation
        NewLocation      = $item.TargetLocation
        NeedsLocFix      = $item.NeedsLocFix
        NeedsLicense     = $item.NeedsLicense
        LocationUpdated  = $locUpdated
        LicenseAssigned  = $licAssigned
        Actions          = $item.Actions
        Status           = $status
        Message          = $combinedMessage
        Timestamp        = (Get-Date -Format "HH:mm:ss")
    })
}

Write-Progress -Completed -Activity "Applying changes"

# =============================================================================
# EXPORT RESULTS
# =============================================================================
$results | Export-Csv -Path $CsvReport -NoTypeInformation -Encoding UTF8
Write-Log "CSV exported: $CsvReport" "SUCCESS"

$skuLabel = if ($targetSku) { $targetSku.SkuPartNumber } else { "N/A (location only)" }
Build-HtmlReport -ResultData $results -ReportPath $HtmlReport -Domain $Domain -SkuLabel $skuLabel -IsDryRun $false

Write-Log "HTML report: $HtmlReport" "SUCCESS"

# =============================================================================
# SUMMARY
# =============================================================================
$successCount  = @($results | Where-Object { $_.Status -eq "Success" }).Count
$partialCount  = @($results | Where-Object { $_.Status -eq "Partial" }).Count
$errorCount    = @($results | Where-Object { $_.Status -eq "Error" }).Count
$licenseCount  = @($results | Where-Object { $_.LicenseAssigned -eq $true }).Count
$locationCount = @($results | Where-Object { $_.LocationUpdated -eq $true }).Count

Write-Log "============================================================"
Write-Log " EXECUTION SUMMARY"
Write-Log "============================================================"
Write-Log " Total processed  : $($results.Count)"
Write-Log " Successful       : $successCount" "SUCCESS"
if ($partialCount -gt 0) {
    Write-Log " Partial (mixed)  : $partialCount" "WARN"
}
if ($errorCount -gt 0) {
    Write-Log " Errors           : $errorCount" "ERROR"
}
else {
    Write-Log " Errors           : $errorCount"
}
Write-Log " Locations updated: $locationCount"
Write-Log " Licenses assigned: $licenseCount"
Write-Log " Log              : $script:LogFile"
Write-Log " HTML report     : $HtmlReport"
Write-Log " CSV export      : $CsvReport"
Write-Log "============================================================"

Disconnect-MgGraph | Out-Null
Invoke-Item $HtmlReport
