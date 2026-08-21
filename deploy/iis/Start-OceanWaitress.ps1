# Starts Waitress for Ocean Logistics (called by Scheduled Task / manually).
param(
    [string]$InstallRoot = "C:\inetpub\ocean\logistics_website",
    [int]$Port = 8050
)

$ErrorActionPreference = "Stop"
Set-Location $InstallRoot

$python = Join-Path $InstallRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Missing venv python: $python. Run Install-OceanIIS.ps1 first."
}

# Kill stale listener on this port (best-effort)
Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue |
    ForEach-Object {
        try { Stop-Process -Id $_.OwningProcess -Force -ErrorAction SilentlyContinue } catch { }
    }

$env:OCEAN_HOST = "127.0.0.1"
$env:OCEAN_PORT = "$Port"

& $python (Join-Path $InstallRoot "serve_production.py")
