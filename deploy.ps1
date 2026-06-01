# deploy.ps1
# Full deploy pipeline:
#   1. clasp push  — uploads local files to the LocalisationMasterTool library
#   2. clasp deploy — publishes a new numbered version
#   3. Updates all 5 stub spreadsheet projects to that new version via the Apps Script API
#
# USAGE
#   .\deploy.ps1
#   .\deploy.ps1 -Message "fix save bug"

param(
    [string]$Message = "deploy"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ── Config ────────────────────────────────────────────────────────────────────

$LibraryIdentifier = "PDGLib"

$StubScriptIds = @(
    "1xavof4et8zzdz3tdbnPmS0_fNGJ-L9pJCATg39e0wlYgVKCiBE6DZ9Vq",
    "1xbHo2zpxLmXVAST-3fupf4b0gs4Qdy7aQmkZZS-4n-9JxGYflcP6G92N",
    "1oo1GoFop3r_5Vn_2nodcJBcMIDWUR6XHbFCWjsa3NNZAe-QRQJsHbp3V",
    "1zVLv9faqncXB3fS0Ek-WKkfcCkOohMrwz1oxd5eQlam2TO7oZfDtJ4XN",
    "1N5OSKlgAhdak45FUen_GbjBJM3ElWmXUhbbp8LpuRyzfftW2C1n4VkCR"
)

# ── Step 1: clasp push ────────────────────────────────────────────────────────

Write-Host ""
Write-Host "── Step 1: Pushing files to library ─────────────────────────────" -ForegroundColor Cyan
$pushOutput = clasp push --force 2>&1
Write-Host ($pushOutput -join "`n")
if ($LASTEXITCODE -ne 0) { Write-Error "clasp push failed."; exit 1 }

# ── Step 2: clasp deploy ─────────────────────────────────────────────────────

Write-Host ""
Write-Host "── Step 2: Publishing new library version ────────────────────────" -ForegroundColor Cyan
$deployOutput = clasp deploy --description $Message 2>&1
Write-Host ($deployOutput -join "`n")
if ($LASTEXITCODE -ne 0) { Write-Error "clasp deploy failed."; exit 1 }

# Parse the version number from output like: "Deployed AKfycb... @5"
$versionMatch = ($deployOutput -join " ") | Select-String -Pattern '@(\d+)'
if (-not $versionMatch) { Write-Error "Could not parse version from clasp deploy output."; exit 1 }
$Version = [int]$versionMatch.Matches[0].Groups[1].Value
Write-Host "  → Published as version $Version" -ForegroundColor Green

# ── Step 3: Read OAuth token from ~/.clasprc.json ─────────────────────────────

$ClasprcPath = "$env:USERPROFILE\.clasprc.json"
if (-not (Test-Path $ClasprcPath)) { Write-Error "Not logged in. Run: clasp login"; exit 1 }

$Clasprc = Get-Content $ClasprcPath -Raw | ConvertFrom-Json
$AccessToken = $Clasprc.tokens.default.access_token

if (-not $AccessToken) { Write-Error "No access token in ~/.clasprc.json. Run: clasp login"; exit 1 }

$Headers = @{
    "Authorization" = "Bearer $AccessToken"
    "Content-Type"  = "application/json"
}

# ── Step 4: Update each stub project ─────────────────────────────────────────

Write-Host ""
Write-Host "── Step 3: Updating stubs to library v$Version ───────────────────" -ForegroundColor Cyan

$ApiBase     = "https://script.googleapis.com/v1/projects"
$SuccessCount = 0

foreach ($ScriptId in $StubScriptIds) {
    Write-Host "  $ScriptId ..." -NoNewline

    try {
        # GET current project content
        $Content = Invoke-RestMethod -Uri "$ApiBase/$ScriptId/content" -Headers $Headers -Method GET

        # Find appsscript.json manifest file
        $ManifestFile = $Content.files | Where-Object { $_.name -eq "appsscript" }
        if (-not $ManifestFile) { Write-Host " no manifest, skipping." -ForegroundColor Yellow; continue }

        $Manifest = $ManifestFile.source | ConvertFrom-Json

        # Find PDGLib in the libraries array
        $LibEntry = $null
        if ($Manifest.dependencies -and $Manifest.dependencies.libraries) {
            $LibEntry = $Manifest.dependencies.libraries | Where-Object { $_.userSymbol -eq $LibraryIdentifier }
        }
        if (-not $LibEntry) { Write-Host " PDGLib not in manifest, skipping." -ForegroundColor Yellow; continue }

        $OldVersion = $LibEntry.version
        $LibEntry.version = [string]$Version

        # PUT updated content back
        $ManifestFile.source = ($Manifest | ConvertTo-Json -Depth 10 -Compress)
        $Body = @{ files = $Content.files } | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Uri "$ApiBase/$ScriptId/content" -Headers $Headers -Method PUT -Body $Body | Out-Null

        Write-Host " v$OldVersion → v$Version ✓" -ForegroundColor Green
        $SuccessCount++
    }
    catch {
        Write-Host " FAILED" -ForegroundColor Red
        Write-Warning $_.Exception.Message
    }
}

# ── Summary ───────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Deployed library v$Version — $SuccessCount / $($StubScriptIds.Count) stubs updated." -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""
