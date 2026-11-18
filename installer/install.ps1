<#
  Installs MCTM Tracker for the current Windows user. No admin rights needed.

  What it does:
    1. finds Python 3.9+ (asks you to install it if missing)
    2. installs the Python packages from requirements.txt
    3. creates the Excel database if there isn't one yet (existing data is never touched)
    4. draws the app icon
    5. adds shortcuts: Desktop "MCTM Tracker", and a Start menu folder with
       MCTM Tracker / Stop MCTM server / Uninstall MCTM Tracker
    6. lists the app in Settings > Apps > Installed apps, so it can be uninstalled there

  Run it with "Install MCTM.bat". Flags (mainly for testing):
    -ShortcutRoot <dir>   put the Desktop and Start menu shortcuts under this folder instead
    -NoRegistry           don't add the Installed apps entry
    -SkipPackages         don't run pip
    -Quiet                no pause / no launch at the end
#>
param(
    [string]$ShortcutRoot = "",
    [switch]$NoRegistry,
    [switch]$SkipPackages,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$AppName = "MCTM Tracker"
$AppId = "MCTM-Tracker"
$AppDir = Split-Path -Parent $PSScriptRoot
$Version = "2.0.0"

function Step($n, $text) { Write-Host ""; Write-Host "[$n/6] $text" -ForegroundColor Cyan }
function Fail($text) {
    Write-Host ""
    Write-Host "  $text" -ForegroundColor Red
    if (-not $Quiet) { Read-Host "Press Enter to close" | Out-Null }
    exit 1
}

Write-Host "============================================================"
Write-Host "  Installing $AppName"
Write-Host "  from $AppDir"
Write-Host "============================================================"

# ---- 1. Python --------------------------------------------------------------
Step 1 "Looking for Python"
# A PC can have several Pythons (e.g. "py" picks 3.14 while "python" is 3.11).
# Collect every one we can see, then prefer one that already has the packages.
$candidates = @()
$probe = "import sys; print(sys.executable); print(sys.version_info >= (3, 9))"
foreach ($cmd in @(@("python"), @("py", "-3"))) {
    if (Get-Command $cmd[0] -ErrorAction SilentlyContinue) {
        try {
            $argList = @()
            if ($cmd.Count -gt 1) { $argList += $cmd[1] }
            $argList += @("-c", $probe)
            $out = & $cmd[0] @argList 2>$null
            if ($LASTEXITCODE -eq 0 -and $out.Count -ge 2 -and $out[1] -eq "True" -and $candidates -notcontains $out[0].Trim()) {
                $candidates += $out[0].Trim()
            }
        } catch { }
    }
}
if (Get-Command py -ErrorAction SilentlyContinue) {
    # Every Python the launcher knows about, e.g. " -V:3.11 *  C:\...\python.exe"
    foreach ($line in (& py -0p 2>$null)) {
        if ($line -match '([A-Za-z]:\\.*python\.exe)\s*$') {
            $p = $Matches[1].Trim()
            if ($candidates -notcontains $p) {
                $ok = & $p -c "import sys; print(sys.version_info >= (3, 9))" 2>$null
                if ($ok -eq "True") { $candidates += $p }
            }
        }
    }
}
$Python = $null
foreach ($p in $candidates) {
    & $p -c "import flask, openpyxl, PIL, requests, dotenv" 2>$null
    if ($LASTEXITCODE -eq 0) { $Python = $p; break }
}
if (-not $Python -and $candidates.Count -gt 0) { $Python = $candidates[0] }
if (-not $Python) {
    Fail "Python 3.9 or newer was not found. Install it from https://www.python.org/downloads/ (tick 'Add python.exe to PATH'), then run Install MCTM.bat again."
}
$PythonW = Join-Path (Split-Path $Python) "pythonw.exe"
if (-not (Test-Path $PythonW)) { $PythonW = $Python }
Write-Host "  Using $Python"

# ---- 2. Packages ------------------------------------------------------------
Step 2 "Installing the Python packages"
if ($SkipPackages) {
    Write-Host "  Skipped."
} else {
    & $Python -c "import flask, openpyxl, PIL, pytesseract, requests, dotenv" 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Already installed."
    } else {
        & $Python -m pip install --user --disable-pip-version-check -r (Join-Path $AppDir "requirements.txt")
        if ($LASTEXITCODE -ne 0) { Fail "Installing packages failed. Check the internet connection and run the installer again." }
    }
}

# ---- 3. Database ------------------------------------------------------------
Step 3 "Checking the Excel database"
$Db = Join-Path $AppDir "data\SchoolDisciplineSystem.xlsx"
if (Test-Path $Db) {
    Write-Host "  Found your existing database. Leaving it untouched."
} else {
    Push-Location $AppDir
    & $Python (Join-Path $AppDir "build_excel_template.py")
    Pop-Location
    if (-not (Test-Path $Db)) { Fail "Couldn't create the database." }
    Write-Host "  Created it with 5 example students. Remove them on the Master student list page before real use."
}

# ---- 4. Icon ----------------------------------------------------------------
Step 4 "Drawing the app icon"
$Icon = Join-Path $PSScriptRoot "mctm.ico"
& $Python (Join-Path $PSScriptRoot "make_icon.py") $Icon 2>$null | Out-Null
if (-not (Test-Path $Icon)) { $Icon = $PythonW; Write-Host "  Couldn't draw it; using the Python icon instead." }

# ---- 5. Shortcuts -----------------------------------------------------------
Step 5 "Adding shortcuts"
if ($ShortcutRoot) {
    $Desktop = Join-Path $ShortcutRoot "Desktop"
    $Programs = Join-Path $ShortcutRoot "Start Menu\Programs"
} else {
    $Desktop = [Environment]::GetFolderPath("Desktop")
    $Programs = [Environment]::GetFolderPath("Programs")
}
$MenuDir = Join-Path $Programs $AppName
New-Item -ItemType Directory -Force -Path $Desktop, $MenuDir | Out-Null

$Shell = New-Object -ComObject WScript.Shell
function New-Shortcut($path, $target, $arguments, $description, $icon) {
    $s = $Shell.CreateShortcut($path)
    $s.TargetPath = $target
    $s.Arguments = $arguments
    $s.WorkingDirectory = $AppDir
    $s.Description = $description
    $s.IconLocation = "$icon,0"
    $s.Save()
    return $path
}

$Uninstaller = Join-Path $PSScriptRoot "uninstall.ps1"
$PowerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$Launcher = Join-Path $AppDir "MCTM.pyw"
$StopScript = Join-Path $PSScriptRoot "stop_server.pyw"

$created = @()
$created += New-Shortcut (Join-Path $Desktop "$AppName.lnk") $PythonW "`"$Launcher`"" "Open the MCTM discipline tracker" $Icon
$created += New-Shortcut (Join-Path $MenuDir "$AppName.lnk") $PythonW "`"$Launcher`"" "Open the MCTM discipline tracker" $Icon
$created += New-Shortcut (Join-Path $MenuDir "Stop MCTM server.lnk") $PythonW "`"$StopScript`"" "Stop the tracker's background server" $Icon
$created += New-Shortcut (Join-Path $MenuDir "Uninstall $AppName.lnk") $PowerShellExe "-NoProfile -ExecutionPolicy Bypass -File `"$Uninstaller`"" "Remove MCTM Tracker from this PC" $Icon
foreach ($c in $created) { Write-Host "  $c" }

# ---- 6. Installed apps entry ------------------------------------------------
Step 6 "Listing it in Settings > Apps"
$RegKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppId"
if ($NoRegistry) {
    Write-Host "  Skipped."
    $RegKey = ""
} else {
    New-Item -Path $RegKey -Force | Out-Null
    $props = @{
        DisplayName     = $AppName
        DisplayVersion  = $Version
        Publisher       = "MCTM"
        DisplayIcon     = $Icon
        InstallLocation = $AppDir
        UninstallString = "`"$PowerShellExe`" -NoProfile -ExecutionPolicy Bypass -File `"$Uninstaller`""
        Comments        = "Discipline and grooming tracker (Excel database + web app)"
    }
    foreach ($k in $props.Keys) { New-ItemProperty -Path $RegKey -Name $k -Value $props[$k] -PropertyType String -Force | Out-Null }
    New-ItemProperty -Path $RegKey -Name NoModify -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $RegKey -Name NoRepair -Value 1 -PropertyType DWord -Force | Out-Null
    Write-Host "  Done."
}

# Remember exactly what was created, so the uninstaller removes only that.
$manifest = [ordered]@{
    version   = $Version
    installed = (Get-Date).ToString("s")
    appDir    = $AppDir
    python    = $Python
    shortcuts = $created
    menuDir   = $MenuDir
    registry  = $RegKey
}
$manifest | ConvertTo-Json | Set-Content -Path (Join-Path $PSScriptRoot "install.json") -Encoding UTF8

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  $AppName is installed." -ForegroundColor Green
Write-Host "  Open it from the '$AppName' icon on your Desktop."
Write-Host "  Other PCs on the network use http://<this-PC's-IP>:5000"
Write-Host "  (If Windows asks about the firewall on first launch, click Allow.)"
Write-Host "============================================================" -ForegroundColor Green

if (-not $Quiet) {
    $answer = Read-Host "Open MCTM Tracker now? [Y/n]"
    if ($answer -eq "" -or $answer -match "^[Yy]") { Start-Process -FilePath $PythonW -ArgumentList "`"$Launcher`"" -WorkingDirectory $AppDir }
}
exit 0
