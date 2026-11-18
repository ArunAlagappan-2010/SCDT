<#
  One-time setup for MCTM Tracker. Run it with "Setup MCTM.cmd"; the app also
  runs it by itself the first time it's opened. Safe to run again: anything
  already installed is skipped.

    1. Python 3.9+           installed with winget if missing (per user, no admin)
    2. Python packages       from requirements.txt
    3. Tesseract OCR         installed with winget if missing (Windows may ask to allow it)
    4. Gemini key (optional) paste one to read handwriting better; Enter skips
    5. Database, icon, Desktop + Start menu shortcuts, Settings > Apps entry
       (installer/install.ps1)

  When it finishes it writes data\setup-done.json, which tells the app that
  setup has run.

  Flags:
    -Check              only report what is installed and what would be done
    -SkipDownloads      never install Python or Tesseract (for testing / offline)
    -Quiet              no questions, no pause at the end
    -ShortcutRoot <dir> -NoRegistry   passed on to install.ps1 (testing)
#>
param(
    [switch]$Check,
    [switch]$SkipDownloads,
    [switch]$Quiet,
    [string]$ShortcutRoot = "",
    [switch]$NoRegistry
)

$ErrorActionPreference = "Stop"
$AppDir = Split-Path -Parent $PSScriptRoot
$DataDir = Join-Path $AppDir "data"
$DoneFile = Join-Path $DataDir "setup-done.json"
$HasWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)

function Step($n, $text) { Write-Host ""; Write-Host "[$n/5] $text" -ForegroundColor Cyan }
function Info($text) { Write-Host "  $text" }
function Warn($text) { Write-Host "  $text" -ForegroundColor Yellow }
function Finish($code) {
    if (-not $Quiet -and -not $Check) { Write-Host ""; Read-Host "Press Enter to close" | Out-Null }
    exit $code
}

function Refresh-Path {
    # Pick up PATH changes made by an installer that just ran, without a restart.
    $machine = [Environment]::GetEnvironmentVariable("Path", "Machine")
    $user = [Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path = "$machine;$user"
}

function Find-Python {
    # Every Python 3.9+ we can see; prefer one that already has the packages.
    $found = @()
    $probe = "import sys; print(sys.executable); print(sys.version_info >= (3, 9))"
    foreach ($cmd in @(@("python"), @("py", "-3"))) {
        if (Get-Command $cmd[0] -ErrorAction SilentlyContinue) {
            try {
                $a = @(); if ($cmd.Count -gt 1) { $a += $cmd[1] }; $a += @("-c", $probe)
                $out = & $cmd[0] @a 2>$null
                if ($LASTEXITCODE -eq 0 -and $out.Count -ge 2 -and $out[1] -eq "True" -and $found -notcontains $out[0].Trim()) {
                    # The Microsoft Store "python" stub isn't a real install.
                    if ($out[0] -notmatch "WindowsApps") { $found += $out[0].Trim() }
                }
            } catch { }
        }
    }
    $common = @(Get-ChildItem -Path (Join-Path $env:LOCALAPPDATA "Programs\Python\Python3*\python.exe") -ErrorAction SilentlyContinue |
                Sort-Object FullName -Descending | ForEach-Object { $_.FullName })
    foreach ($p in $common) { if ($found -notcontains $p) { $found += $p } }
    foreach ($p in $found) {
        & $p -c "import flask, openpyxl, PIL, requests, dotenv" 2>$null
        if ($LASTEXITCODE -eq 0) { return $p }
    }
    if ($found.Count -gt 0) { return $found[0] }
    return $null
}

function Find-Tesseract {
    $paths = @(
        "C:\Program Files\Tesseract-OCR\tesseract.exe",
        "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\Tesseract-OCR\tesseract.exe")
    )
    foreach ($p in $paths) { if (Test-Path $p) { return $p } }
    $cmd = Get-Command tesseract -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    return $null
}

function Winget-Install($id, $label, $extra) {
    if ($SkipDownloads) { Warn "Skipped installing $label (-SkipDownloads)."; return $false }
    if (-not $HasWinget) {
        Warn "Can't install $label automatically: winget (App Installer) isn't on this PC."
        return $false
    }
    Info "Installing $label with winget. This downloads it from the internet and can take a few minutes..."
    $args2 = @("install", "-e", "--id", $id, "--silent", "--accept-package-agreements", "--accept-source-agreements") + $extra
    & winget @args2
    Refresh-Path
    return ($LASTEXITCODE -eq 0)
}

Write-Host "============================================================"
Write-Host ("  MCTM Tracker - one-time setup" + $(if ($Check) { " (check only)" } else { "" }))
Write-Host "  $AppDir"
Write-Host "============================================================"
$report = [ordered]@{}

# ---- 1. Python --------------------------------------------------------------
Step 1 "Python"
$Python = Find-Python
if ($Python) {
    Info "Found $Python"
} elseif ($Check) {
    Warn "Not installed. Setup would install Python 3.12 with winget."
} else {
    Winget-Install "Python.Python.3.12" "Python 3.12" @("--scope", "user") | Out-Null
    $Python = Find-Python
    if (-not $Python) {
        Warn "Python still isn't available. Install it from https://www.python.org/downloads/"
        Warn "(tick 'Add python.exe to PATH'), then run Setup MCTM.cmd again."
        Finish 1
    }
    Info "Installed: $Python"
}
$report.python = $Python

# ---- 2. Packages ------------------------------------------------------------
Step 2 "Python packages"
$pkgsOk = $false
if ($Python) {
    & $Python -c "import flask, openpyxl, PIL, pytesseract, requests, dotenv" 2>$null
    $pkgsOk = ($LASTEXITCODE -eq 0)
}
if ($pkgsOk) {
    Info "All installed."
} elseif ($Check) {
    Warn "Missing. Setup would run: pip install -r requirements.txt"
} else {
    & $Python -m pip install --user --disable-pip-version-check -r (Join-Path $AppDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { Warn "Installing packages failed. Check the internet connection and run Setup MCTM.cmd again."; Finish 1 }
    Info "Installed."
}

# ---- 3. Tesseract -----------------------------------------------------------
Step 3 "Tesseract (reads register photos offline)"
$Tess = Find-Tesseract
if ($Tess) {
    Info "Found $Tess"
} elseif ($Check) {
    Warn "Not installed. Setup would install it with winget (UB-Mannheim.TesseractOCR)."
} else {
    $ok = Winget-Install "UB-Mannheim.TesseractOCR" "Tesseract OCR" @()
    $Tess = Find-Tesseract
    if ($Tess) {
        Info "Installed: $Tess"
    } else {
        Warn "Tesseract isn't installed. Scanning still works with a Gemini key (next step),"
        Warn "or install it later from https://github.com/UB-Mannheim/tesseract/wiki"
    }
}
$report.tesseract = $Tess

# ---- 4. Gemini key (optional) -----------------------------------------------
Step 4 "Gemini key (optional, much better on handwriting)"
$EnvFile = Join-Path $AppDir ".env"
$hasKey = (Test-Path $EnvFile) -and ((Get-Content $EnvFile -Raw) -match "(?m)^\s*GEMINI_API_KEY\s*=\s*\S+")
if ($hasKey) {
    Info "A key is already set in .env."
} elseif ($Check) {
    Info "No key set. Setup would offer to add one (optional)."
} elseif ($Quiet) {
    Info "Skipped. Run SETUP_GEMINI_KEY.bat any time to add one."
} else {
    Info "Get a free key at https://aistudio.google.com/apikey"
    Info "Photos are then sent to Google to be read; leave blank to keep everything offline."
    $key = (Read-Host "  Paste the key, or press Enter to skip").Trim()
    if ($key) {
        # Same script as SETUP_GEMINI_KEY.bat: checks the key, saves it without a BOM.
        & (Join-Path $PSScriptRoot "set_gemini_key.ps1") -Key $key -NoRestart -Quiet
        $hasKey = ($LASTEXITCODE -eq 0) -and (Test-Path $EnvFile) -and ((Get-Content $EnvFile -Raw) -match "(?m)^\s*GEMINI_API_KEY\s*=\s*\S+")
    } else {
        Info "Skipped. Run SETUP_GEMINI_KEY.bat any time to add one."
    }
}
$report.gemini = $hasKey

if ($Check) {
    Write-Host ""
    Write-Host "Check only: nothing was installed." -ForegroundColor Green
    exit 0
}

# ---- 5. Database, shortcuts, Apps entry ------------------------------------
Step 5 "Database and shortcuts"
$installArgs = @{ SkipPackages = $true; Quiet = $true }
if ($ShortcutRoot) { $installArgs.ShortcutRoot = $ShortcutRoot }
if ($NoRegistry) { $installArgs.NoRegistry = $true }
& (Join-Path $PSScriptRoot "install.ps1") @installArgs
if ($LASTEXITCODE -ne 0) { Warn "Adding shortcuts failed (see above)."; Finish 1 }

New-Item -ItemType Directory -Force -Path $DataDir | Out-Null
$report.finished = (Get-Date).ToString("s")
$report | ConvertTo-Json | Set-Content -Path $DoneFile -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Setup is done. Open MCTM Tracker from the Desktop icon." -ForegroundColor Green
if (-not $Tess -and -not $hasKey) {
    Write-Host "  Note: no reading engine yet, so scans can't read names." -ForegroundColor Yellow
}
Write-Host "============================================================" -ForegroundColor Green
Finish 0
