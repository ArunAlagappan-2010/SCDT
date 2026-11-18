<#
  Adds (or replaces) the Gemini API key, so Gemini reads register photos first
  and Tesseract is only the fallback. Run it with SETUP_GEMINI_KEY.bat.

    1. asks for the key (get one free at https://aistudio.google.com/apikey)
    2. checks it with Google (one small request; no photo is sent)
    3. saves it as GEMINI_API_KEY in .env (never uploaded to GitHub)
    4. restarts the tracker so it uses Gemini straight away

  Flags (testing): -Key <key> skips the question; -NoVerify skips the check;
  -NoRestart leaves the tracker alone; -Remove deletes the key (Tesseract only).
#>
param(
    [string]$Key = "",
    [switch]$NoVerify,
    [switch]$NoRestart,
    [switch]$Remove,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$AppDir = Split-Path -Parent $PSScriptRoot
$EnvFile = Join-Path $AppDir ".env"
$Template = Join-Path $AppDir ".env.example"

function Done($code) {
    if (-not $Quiet) { Write-Host ""; Read-Host "Press Enter to close" | Out-Null }
    exit $code
}

function Find-Python {
    $done = Join-Path $AppDir "data\setup-done.json"
    if (Test-Path $done) {
        try { $p = (Get-Content $done -Raw | ConvertFrom-Json).python; if ($p -and (Test-Path $p)) { return $p } } catch { }
    }
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Restart-Tracker {
    if ($NoRestart) { return }
    $python = Find-Python
    if (-not $python) { Write-Host "  Restart the tracker yourself to use the new setting."; return }
    $pythonw = Join-Path (Split-Path $python) "pythonw.exe"
    if (-not (Test-Path $pythonw)) { $pythonw = $python }
    $running = $false
    try { $running = [bool](Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:5000/" -TimeoutSec 2) } catch { }
    if ($running) {
        & $python (Join-Path $PSScriptRoot "stop_server.pyw") --quiet | Out-Null
        Start-Sleep -Milliseconds 800
    }
    Start-Process -FilePath $pythonw -ArgumentList "`"$(Join-Path $AppDir 'MCTM.pyw')`" --no-window --skip-setup" -WorkingDirectory $AppDir
    Write-Host "  Tracker restarted. Refresh its window (F5)." -ForegroundColor Green
}

function Write-Key($value) {
    $text = if (Test-Path $EnvFile) { Get-Content $EnvFile -Raw } elseif (Test-Path $Template) { Get-Content $Template -Raw } else { "" }
    if ($text -match "(?m)^\s*GEMINI_API_KEY\s*=.*$") {
        $text = [regex]::Replace($text, "(?m)^\s*GEMINI_API_KEY\s*=.*$", "GEMINI_API_KEY=$value")
    } else {
        $text = "GEMINI_API_KEY=$value`r`n" + $text
    }
    # UTF-8 without a BOM: python-dotenv would read a BOM as part of the first name.
    [System.IO.File]::WriteAllText($EnvFile, $text, (New-Object System.Text.UTF8Encoding($false)))
}

Write-Host "============================================================"
Write-Host "  MCTM Tracker - Gemini key"
Write-Host "  Gemini reads the register photos; Tesseract is the backup."
Write-Host "============================================================"

if ($Remove) {
    Write-Key ""
    Write-Host ""
    Write-Host "  Key removed. Photos are read offline by Tesseract only."
    Restart-Tracker
    Done 0
}

if (-not $Key) {
    Write-Host ""
    Write-Host "  1. Open https://aistudio.google.com/apikey and sign in with a Google account."
    Write-Host "  2. Click 'Create API key' and copy it."
    Write-Host "  3. Paste it below (right-click pastes in this window)."
    Write-Host ""
    Write-Host "  Note: with a key, register photos are sent to Google to be read."
    Write-Host ""
    $Key = (Read-Host "  Gemini API key").Trim().Trim('"').Trim("'")
}
if (-not $Key) {
    Write-Host "  Nothing entered; no change made."
    Done 0
}
if ($Key -match "\s") {
    Write-Host "  That doesn't look like a key (it has spaces). No change made." -ForegroundColor Yellow
    Done 1
}

if (-not $NoVerify) {
    Write-Host ""
    Write-Host "  Checking the key with Google..."
    try {
        $resp = Invoke-WebRequest -UseBasicParsing -TimeoutSec 20 -Headers @{ "x-goog-api-key" = $Key } `
            -Uri "https://generativelanguage.googleapis.com/v1beta/models?pageSize=1"
        Write-Host "  The key works." -ForegroundColor Green
    } catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        if ($status -in 400, 401, 403) {
            Write-Host "  Google rejected that key (HTTP $status). Check you copied all of it. No change made." -ForegroundColor Yellow
            Done 1
        }
        Write-Host "  Couldn't reach Google to check it ($($_.Exception.Message))." -ForegroundColor Yellow
        Write-Host "  Saving it anyway; scans fall back to Tesseract while Gemini can't be reached."
    }
}

Write-Key $Key
Write-Host "  Saved in .env (never uploaded to GitHub)."
Restart-Tracker
Write-Host ""
Write-Host "  Done. Gemini is now the main reader; Tesseract is the backup." -ForegroundColor Green
Done 0
