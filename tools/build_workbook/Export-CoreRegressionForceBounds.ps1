# ДЛЯ ТЕСТОВ: экспортирует геометрию и диаграммы, не запуская НДС.
param([string]$Directory='docs/regression/Performance/CoreRegression_2026-10-15')
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
if(-not $directoryPath.StartsWith((Join-Path $root 'docs/regression/Performance/'),[StringComparison]::OrdinalIgnoreCase)){throw 'Performance evidence directory required.'}
$run=Get-Content -LiteralPath (Join-Path $directoryPath 'Run.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $run.AllJobsExecuted){throw 'Wait for the timed main run first.'}
$temporary=Join-Path $root 'tmp/CoreRegressionForceBounds'
New-Item -ItemType Directory -Path $temporary -Force | Out-Null
$combined=Join-Path $temporary 'Combined.bas'
if(Test-Path -LiteralPath $combined){throw 'Existing bounds preparation; choose another evidence run.'}
$wrapper=@'

' ДЛЯ ТЕСТОВ: те же сохранённые элементы и СП35, что в нагрузочном прогоне.
Public Function DumpCoreRegressionForceBounds() As String
    Dim oldSystem As Variant, output As String
    Dim settings As CSystemSettingsReader, units As CUnitSystem
    On Error GoTo Failed
    oldSystem = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula
    SetSystemSetting "Geometry.Source", "AutoCAD"
    SetSystemSetting "SLS.Crack.Code", "SP35"
    output = DumpFocusedForceBounds("SavedWorkbook")
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    output = output & "ACCEPTANCE_TOLERANCE|N=" & FocusedBoundNumber(Abs(units.InputForceToInternal(settings.GetDouble("Solver.ToleranceN")))) & _
        "|Mx=" & FocusedBoundNumber(Abs(units.InputMomentMxToInternal(settings.GetDouble("Solver.ToleranceMx")))) & _
        "|My=" & FocusedBoundNumber(Abs(units.InputMomentMyToInternal(settings.GetDouble("Solver.ToleranceMy")))) & vbCrLf
    GoTo Restore
Failed:
    output = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
Restore:
    If IsArray(oldSystem) Then ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange.Formula = oldSystem
    DumpCoreRegressionForceBounds = output
End Function
'@
$helper=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'FocusedForceBounds.bas'),[Text.Encoding]::UTF8)
[IO.File]::WriteAllText($combined,$helper+"`r`n"+$wrapper,[Text.UTF8Encoding]::new($true))
& (Join-Path $PSScriptRoot 'Prepare-FocusedAcceptance.ps1') -OutputDirectory ($Directory+'/ForceBounds') -ExtraHelperFile '../../tmp/CoreRegressionForceBounds/Combined.bas'
$preparation=Get-Content -LiteralPath (Join-Path $directoryPath 'ForceBounds/Preparation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$inputs=Get-Content -LiteralPath (Join-Path $directoryPath 'InputsBefore.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if($preparation.SourceSHA256 -cne $inputs.CurrentSHA256.ToUpperInvariant()){throw 'Main workbook changed since comparison.'}
Copy-Item -LiteralPath $combined -Destination (Join-Path $directoryPath 'ForceBounds/Combined.bas')
& (Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1') -SourceWorkbook ($Directory+'/ForceBounds/RC_Section_NDM.xlsm') -ReportPath ($Directory+'/ForceBounds/Model.txt') -Macro 'modTestBatchCalculation.DumpCoreRegressionForceBounds' -Mode No -TimeoutSeconds 120
if($LASTEXITCODE -ne 0){throw 'Force model export failed.'}
$report=Join-Path $directoryPath 'ForceBounds/Model.txt'
$body=Get-Content -LiteralPath $report -Raw -Encoding UTF8
if(@([regex]::Matches($body,'(?m)^FIBER\|')).Count -ne 65 -or $body -match 'RUNTIME ERROR'){throw 'Unexpected saved force model.'}
[ordered]@{SourceSHA256=$preparation.SourceSHA256;DataSHA256=(Get-FileHash -LiteralPath $report).Hash;ConcreteFibers=45;Rebars=20;EquilibriumSolverExecuted=$false;CoreChanged=$false} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directoryPath 'ForceBounds/ModelManifest.json') -Encoding UTF8
