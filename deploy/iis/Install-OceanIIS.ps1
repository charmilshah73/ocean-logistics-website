# Install Ocean Logistics Control Tower behind IIS (Admin PowerShell).
#
# What it does:
#   1. Enables IIS + URL Rewrite prerequisites (ARR must be installed separately if missing)
#   2. Copies this repo to C:\inetpub\ocean\logistics_website (avoids OneDrive locks)
#   3. Creates venv, installs requirements + waitress
#   4. Creates IIS site + reverse-proxy web.config
#   5. Registers a Scheduled Task to keep Waitress running
#
# Usage (Run as Administrator):
#   cd <path-to-logistics_website>
#   powershell -ExecutionPolicy Bypass -File .\deploy\iis\Install-OceanIIS.ps1
#
# Optional:
#   .\deploy\iis\Install-OceanIIS.ps1 -SitePort 8080 -SkipCopy

[CmdletBinding()]
param(
    [string]$SourceRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path,
    [string]$InstallRoot = "C:\inetpub\ocean\logistics_website",
    [string]$SiteName = "OceanLogistics",
    [int]$SitePort = 8080,
    [int]$AppPort = 8050,
    [switch]$SkipCopy
)

$ErrorActionPreference = "Stop"

function Assert-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Run this script in an elevated (Administrator) PowerShell."
    }
}

function Ensure-WindowsFeature([string[]]$Names) {
    foreach ($name in $Names) {
        $feature = Get-WindowsOptionalFeature -Online -FeatureName $name -ErrorAction SilentlyContinue
        if ($null -eq $feature) { continue }
        if ($feature.State -ne "Enabled") {
            Write-Host "Enabling Windows feature: $name"
            Enable-WindowsOptionalFeature -Online -FeatureName $name -All -NoRestart | Out-Null
        }
    }
}

Assert-Admin

Write-Host "==> Source: $SourceRoot"
Write-Host "==> Install: $InstallRoot"
Write-Host "==> IIS site: $SiteName on port $SitePort -> 127.0.0.1:$AppPort"

# --- IIS role ---
Ensure-WindowsFeature @(
    "IIS-WebServerRole",
    "IIS-WebServer",
    "IIS-CommonHttpFeatures",
    "IIS-StaticContent",
    "IIS-DefaultDocument",
    "IIS-HttpErrors",
    "IIS-ApplicationDevelopment",
    "IIS-NetFxExtensibility45",
    "IIS-HealthAndDiagnostics",
    "IIS-HttpLogging",
    "IIS-Security",
    "IIS-RequestFiltering",
    "IIS-Performance",
    "IIS-HttpCompressionStatic",
    "IIS-WebServerManagementTools",
    "IIS-ManagementConsole",
    "IIS-WebSockets"
)

Import-Module WebAdministration -ErrorAction Stop

# --- Copy app off OneDrive ---
$backupData = $null
if (-not $SkipCopy) {
    New-Item -ItemType Directory -Force -Path (Split-Path $InstallRoot) | Out-Null
    Write-Host "Copying application (excluding .git, __pycache__, .venv)..."
    if (Test-Path $InstallRoot) {
        # Keep existing data\ if present
        if (Test-Path (Join-Path $InstallRoot "data")) {
            $backupData = Join-Path $env:TEMP ("ocean_data_backup_" + (Get-Date -Format "yyyyMMddHHmmss"))
            Copy-Item (Join-Path $InstallRoot "data") $backupData -Recurse -Force
        }
    }
    else {
        New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
    }

    $robolog = Join-Path $env:TEMP "ocean_robocopy.log"
    $args = @(
        $SourceRoot, $InstallRoot,
        "/MIR", "/XD", ".git", "__pycache__", ".venv", "node_modules", ".cursor",
        "/XF", "*.pyc",
        "/NFL", "/NDL", "/NJH", "/NJS", "/NP",
        "/R:2", "/W:2",
        "/LOG:$robolog"
    )
    & robocopy @args | Out-Null
    $rc = $LASTEXITCODE
    if ($rc -ge 8) { throw "robocopy failed with code $rc (see $robolog)" }

    if ($backupData -and (Test-Path $backupData)) {
        Write-Host "Restoring previous data\ folder..."
        Remove-Item (Join-Path $InstallRoot "data") -Recurse -Force -ErrorAction SilentlyContinue
        Copy-Item $backupData (Join-Path $InstallRoot "data") -Recurse -Force
        Remove-Item $backupData -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# --- Python venv ---
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) { throw "python not found on PATH. Install Python 3.11+ and re-run." }
$python = $pythonCmd.Source

$venvPython = Join-Path $InstallRoot ".venv\Scripts\python.exe"
$venvPip = Join-Path $InstallRoot ".venv\Scripts\pip.exe"
if (-not (Test-Path $venvPython)) {
    Write-Host "Creating virtualenv..."
    & $python -m venv (Join-Path $InstallRoot ".venv")
}
Write-Host "Installing Python packages..."
& $venvPip install --upgrade pip
& $venvPip install -r (Join-Path $InstallRoot "requirements.txt")

# Ensure data folders exist + Modify for app-pool / LOCAL SERVICE
$dataRoot = Join-Path $InstallRoot "data"
foreach ($sub in @("input", "staging", "archive", "current", "output", "logs")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $dataRoot $sub) | Out-Null
}
icacls $dataRoot /grant "IIS_IUSRS:(OI)(CI)M" /T | Out-Null
icacls $dataRoot /grant "IUSR:(OI)(CI)M" /T | Out-Null
icacls $InstallRoot /grant "IIS_IUSRS:(OI)(CI)RX" /T | Out-Null

# --- web.config into site physical path ---
$siteRoot = Join-Path (Split-Path $InstallRoot) "iis_site"
New-Item -ItemType Directory -Force -Path $siteRoot | Out-Null
Copy-Item (Join-Path $PSScriptRoot "web.config") (Join-Path $siteRoot "web.config") -Force

# Patch AppPort in web.config if not 8050
if ($AppPort -ne 8050) {
    $cfg = Get-Content (Join-Path $siteRoot "web.config") -Raw
    $cfg = $cfg -replace "127\.0\.0\.1:8050", "127.0.0.1:$AppPort"
    Set-Content -Path (Join-Path $siteRoot "web.config") -Value $cfg -Encoding UTF8
}

# --- App pool + site ---
$poolName = "$SiteName`AppPool"
if (-not (Test-Path "IIS:\AppPools\$poolName")) {
    New-WebAppPool -Name $poolName | Out-Null
}
Set-ItemProperty "IIS:\AppPools\$poolName" -Name managedRuntimeVersion -Value ""
Set-ItemProperty "IIS:\AppPools\$poolName" -Name startMode -Value "AlwaysRunning"

if (Get-Website -Name $SiteName -ErrorAction SilentlyContinue) {
    Remove-Website -Name $SiteName
}
New-Website -Name $SiteName -Port $SitePort -PhysicalPath $siteRoot -ApplicationPool $poolName | Out-Null

# Raise request limits at site level too
$filter = "system.webServer/security/requestFiltering/requestLimits"
Set-WebConfigurationProperty -PSPath "IIS:\Sites\$SiteName" -Filter $filter -Name maxAllowedContentLength -Value 104857600

# --- ARR proxy (if installed) ---
$arrProxy = "system.webServer/proxy"
try {
    Set-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST" -Filter $arrProxy -Name enabled -Value $true
    Write-Host "ARR proxy enabled."
}
catch {
    Write-Warning @"
ARR (Application Request Routing) proxy setting failed.
Install ARR + URL Rewrite from Microsoft, then re-run or enable proxy in IIS Manager:
  Server node -> Application Request Routing Cache -> Server Proxy Settings -> Enable proxy
Download:
  URL Rewrite: https://www.iis.net/downloads/microsoft/url-rewrite
  ARR:         https://www.iis.net/downloads/microsoft/application-request-routing
"@
}

# Allow server variables for forwarded headers (best-effort)
try {
    Add-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST" `
        -Filter "system.webServer/rewrite/allowedServerVariables" `
        -Name "." -Value @{name = "HTTP_X_FORWARDED_PROTO" } -ErrorAction SilentlyContinue
    Add-WebConfigurationProperty -PSPath "MACHINE/WEBROOT/APPHOST" `
        -Filter "system.webServer/rewrite/allowedServerVariables" `
        -Name "." -Value @{name = "HTTP_X_FORWARDED_HOST" } -ErrorAction SilentlyContinue
}
catch { }

# --- Scheduled task: keep Waitress alive ---
$taskName = "OceanLogisticsWaitress"
$deployDir = Join-Path $InstallRoot "deploy\iis"
New-Item -ItemType Directory -Force -Path $deployDir | Out-Null
Copy-Item (Join-Path $PSScriptRoot "Start-OceanWaitress.ps1") (Join-Path $deployDir "Start-OceanWaitress.ps1") -Force
Copy-Item (Join-Path $InstallRoot "serve_production.py") (Join-Path $InstallRoot "serve_production.py") -Force -ErrorAction SilentlyContinue
$runner = Join-Path $deployDir "Start-OceanWaitress.ps1"

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$runner`" -InstallRoot `"$InstallRoot`" -Port $AppPort"
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
Start-ScheduledTask -TaskName $taskName

Start-Sleep -Seconds 3

Write-Host ""
Write-Host "Done."
Write-Host "  IIS URL:        http://localhost:$SitePort/"
Write-Host "  Admin:          http://localhost:$SitePort/admin"
Write-Host "  Python direct:  http://127.0.0.1:$AppPort/"
Write-Host "  App folder:     $InstallRoot"
Write-Host "  Waitress task:  $taskName"
Write-Host ""
Write-Host "Open the IIS URL, then test upload on /admin."
Write-Host "If reverse proxy returns 502, install URL Rewrite + ARR and re-enable proxy."
