<#
  Removes MCTM Tracker from this PC.

    1. stops the background server
    2. copies your data (the Excel database, .env and scanned photos) to
       Documents\MCTM Tracker backup <date>, unless you say not to
    3. removes the shortcuts and the Settings > Apps entry the installer made
    4. optionally moves the program folder itself to the Recycle Bin

  Nothing is permanently deleted: removed files go to the Recycle Bin.
  Python and its packages are left installed (other programs may use them).

  Run from Start menu > MCTM Tracker > Uninstall, from Settings > Apps, or with
  "Uninstall MCTM.bat". Flags (for scripted use / testing):
    -Quiet            no questions; keeps a data backup, keeps the program folder
    -NoBackup         with -Quiet: skip the data backup
    -RemoveAppFolder  with -Quiet: also move the program folder to the Recycle Bin
    -BackupRoot <dir> where to put the backup (default: Documents)
#>
param(
    [switch]$Quiet,
    [switch]$NoBackup,
    [switch]$RemoveAppFolder,
    [string]$BackupRoot = ""
)

$ErrorActionPreference = "Stop"
$AppName = "MCTM Tracker"
$AppDir = Split-Path -Parent $PSScriptRoot
$ManifestPath = Join-Path $PSScriptRoot "install.json"

Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms

function Send-ToRecycleBin($path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    $ui = [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs
    $bin = [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
    if (Test-Path -LiteralPath $path -PathType Container) {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($path, $ui, $bin)
    } else {
        [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($path, $ui, $bin)
    }
}

function Ask($text, $buttons = "YesNo", $default = "Button1") {
    if ($Quiet) { return $null }
    return [System.Windows.Forms.MessageBox]::Show($text, "Uninstall $AppName",
        [System.Windows.Forms.MessageBoxButtons]::$buttons,
        [System.Windows.Forms.MessageBoxIcon]::Question,
        [System.Windows.Forms.MessageBoxDefaultButton]::$default)
}

if (-not $Quiet) {
    $go = Ask "Remove $AppName from this PC?`n`nYour data is backed up first unless you choose otherwise." "OKCancel"
    if ($go -ne "OK") { exit 0 }
}

$manifest = $null
if (Test-Path $ManifestPath) { $manifest = Get-Content $ManifestPath -Raw | ConvertFrom-Json }

# ---- 1. Stop the server -----------------------------------------------------
Write-Host "Stopping the server..."
$python = $null
if ($manifest -and $manifest.python -and (Test-Path $manifest.python)) { $python = $manifest.python }
elseif (Get-Command python -ErrorAction SilentlyContinue) { $python = (Get-Command python).Source }
if ($python) {
    & $python (Join-Path $PSScriptRoot "stop_server.pyw") --quiet | Out-Host
}

# ---- 2. Back up the data ----------------------------------------------------
$dataDir = Join-Path $AppDir "data"
$backupDir = $null
$doBackup = -not $NoBackup
if (-not $Quiet) {
    $answer = Ask "Keep a copy of your data?`n`nThe Excel database (every student and entry), your Gemini key file and the scanned register photos will be copied to your Documents folder." "YesNo"
    $doBackup = ($answer -eq "Yes")
}
if ($doBackup -and (Test-Path $dataDir)) {
    if (-not $BackupRoot) { $BackupRoot = [Environment]::GetFolderPath("MyDocuments") }
    $backupDir = Join-Path $BackupRoot ("$AppName backup " + (Get-Date).ToString("yyyy-MM-dd HHmm"))
    New-Item -ItemType Directory -Force -Path $backupDir | Out-Null
    Copy-Item -Path (Join-Path $dataDir "*") -Destination $backupDir -Recurse -Force -Exclude "server.pid"
    $envFile = Join-Path $AppDir ".env"
    if (Test-Path $envFile) { Copy-Item $envFile -Destination $backupDir -Force }
    Write-Host "Data backed up to $backupDir"
}

# ---- 3. Shortcuts and the Installed apps entry ------------------------------
Write-Host "Removing shortcuts..."
$shortcuts = @()
$menuDir = $null
$regKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\MCTM-Tracker"
if ($manifest) {
    $shortcuts = @($manifest.shortcuts)
    $menuDir = $manifest.menuDir
    if ($manifest.PSObject.Properties.Name -contains "registry") { $regKey = $manifest.registry }
} else {
    # No manifest (installed by hand?): fall back to the standard locations.
    $shortcuts = @(Join-Path ([Environment]::GetFolderPath("Desktop")) "$AppName.lnk")
    $menuDir = Join-Path ([Environment]::GetFolderPath("Programs")) $AppName
}
foreach ($s in $shortcuts) { if ($s) { Send-ToRecycleBin $s } }
if ($menuDir -and (Test-Path $menuDir)) { Send-ToRecycleBin $menuDir }
if ($regKey -and (Test-Path $regKey)) { Remove-Item -Path $regKey -Recurse }
foreach ($f in @($ManifestPath, (Join-Path $PSScriptRoot "mctm.ico"))) { Send-ToRecycleBin $f }

# ---- 4. The program folder --------------------------------------------------
$removeFolder = [bool]$RemoveAppFolder
if (-not $Quiet) {
    $warn = if ($backupDir) { "Your data copy is safe in:`n$backupDir" } else { "You chose NOT to keep a copy of your data." }
    $answer = Ask "Also move the program folder to the Recycle Bin?`n`n$AppDir`n`n$warn`n`nChoose No to keep the folder (you can reinstall from it later)." "YesNo" "Button2"
    $removeFolder = ($answer -eq "Yes")
}
if ($removeFolder) {
    Set-Location $env:TEMP
    try {
        Send-ToRecycleBin $AppDir
        Write-Host "Program folder moved to the Recycle Bin."
    } catch {
        Write-Host "Couldn't move the program folder (a file in it is open). Close Excel and delete it by hand: $AppDir"
    }
}

$done = "$AppName has been removed."
if ($backupDir) { $done += "`n`nYour data is saved in:`n$backupDir" }
if (-not $removeFolder) { $done += "`n`nThe program folder was kept:`n$AppDir" }
Write-Host $done
if (-not $Quiet) {
    [System.Windows.Forms.MessageBox]::Show($done, "Uninstall $AppName", "OK", "Information") | Out-Null
}
exit 0
