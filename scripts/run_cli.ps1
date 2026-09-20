param(
  [Parameter(Mandatory=$true)][ValidateSet('gui','inspect','inspect-deep','plan-edit','read','diff','safe-edit','smart-edit','smart-recalc','open-excel','cell-info','formula-scan','bulk-preview','bulk-edit')][string]$Command,
  [string]$Path,
  [string]$Before,
  [string]$After,
  [string]$Sheet,
  [string[]]$Edit,
  [string]$Output,
  [string]$Engine='auto',
  [string]$Operation='cell-edit',
  [string]$Cell,
  [string]$StartCell,
  [string]$Ranges,
  [string]$DataFile,
  [string]$DataSheet,
  [switch]$Recalculate,
  [switch]$NoFullRebuild,
  [switch]$PreferNative,
  [switch]$NoBackup,
  [switch]$NoFormulas,
  [switch]$ClearEmpty,
  [switch]$Trace,
  [switch]$Issues
)
$ErrorActionPreference='Stop'
$Root=Split-Path -Parent $PSScriptRoot
$Py=Join-Path $Root '.venv\Scripts\python.exe'
if(-not(Test-Path $Py)){ throw "Virtual environment not found. Run .\scripts\bootstrap_windows.ps1 first." }
$args=@('-m','excel_power_engine.cli',$Command)
switch($Command){
  'gui' { }
  'inspect' { $args += $Path }
  'inspect-deep' { $args += $Path }
  'plan-edit' { $args += $Path; if($Operation){$args += @('--operation',$Operation)}; if($Recalculate){$args += '--recalculate'}; if($PreferNative){$args += '--prefer-native'} }
  'read' { $args += $Path; $args += @('--sheet',$Sheet); $args += @('--engine',$Engine) }
  'diff' { $args += $Before; $args += $After; if($Sheet){$args += @('--sheet',$Sheet)} }
  'safe-edit' { $args += $Path; $args += @('--sheet',$Sheet); foreach($e in $Edit){$args += @('--edit',$e)}; if($Output){$args += @('--output',$Output)}; if($NoBackup){$args += '--no-backup'} }
  'smart-edit' { $args += $Path; $args += @('--sheet',$Sheet); foreach($e in $Edit){$args += @('--edit',$e)}; if($Output){$args += @('--output',$Output)}; if($NoBackup){$args += '--no-backup'}; if($Recalculate){$args += '--recalculate'}; if($PreferNative){$args += '--prefer-native'} }
  'smart-recalc' { $args += $Path; if($Output){$args += @('--output',$Output)}; if($NoBackup){$args += '--no-backup'}; if($NoFullRebuild){$args += '--no-full-rebuild'} }
  'open-excel' { $args += $Path; if($Sheet){$args += @('--sheet',$Sheet)}; if($Cell){$args += @('--cell',$Cell)} }
  'cell-info' { $args += $Path; $args += @('--sheet',$Sheet,'--cell',$Cell); if($Trace){$args += '--trace'}; if($Issues){$args += '--issues'} }
  'formula-scan' { $args += $Path }
  'bulk-preview' { $args += $Path; $args += @('--sheet',$Sheet); if($StartCell){$args += @('--start-cell',$StartCell)}; if($Ranges){$args += @('--ranges',$Ranges)}; $args += @('--data-file',$DataFile); if($DataSheet){$args += @('--data-sheet',$DataSheet)}; if($NoFormulas){$args += '--no-formulas'}; if($ClearEmpty){$args += '--clear-empty'} }
  'bulk-edit' { $args += $Path; $args += @('--sheet',$Sheet); if($StartCell){$args += @('--start-cell',$StartCell)}; if($Ranges){$args += @('--ranges',$Ranges)}; $args += @('--data-file',$DataFile); if($DataSheet){$args += @('--data-sheet',$DataSheet)}; if($Output){$args += @('--output',$Output)}; if($NoBackup){$args += '--no-backup'}; if($NoFormulas){$args += '--no-formulas'}; if($ClearEmpty){$args += '--clear-empty'} }
}
& $Py @args
exit $LASTEXITCODE
