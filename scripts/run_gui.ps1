$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
$Py=Join-Path $Root '.venv\Scripts\python.exe'
if(-not(Test-Path $Py)){ throw "Virtual environment not found. Run .\scripts\bootstrap_windows.ps1 -Full first." }
& $Py -m excel_power_engine.gui
exit $LASTEXITCODE
