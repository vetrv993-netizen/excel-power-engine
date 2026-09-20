param([switch]$Full)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
Set-Location $Root
if(-not (Test-Path '.venv\\Scripts\\python.exe')){ py -3.11 -m venv .venv }
$Py=Join-Path $Root '.venv\\Scripts\\python.exe'
& $Py -m pip install --upgrade pip
if($Full){ & $Py -m pip install -e '.[full,dev]' } else { & $Py -m pip install -e '.[dev]' }
& $Py -m pytest -q
Write-Host 'Excel Power Engine v0.6.0 is ready.' -ForegroundColor Green
