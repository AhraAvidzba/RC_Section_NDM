# Проверяет изменённый код напрямую, неизменённый — с ранее принятым native VBA.
# Это не округление чисел: VBE-форматирование старых литералов уже подтверждено
# прежней независимой компиляцией и не должно маскировать новый исходный код.
param([string]$Directory='docs/regression/GeometryFull_2026-10-13/PlateauFactorV2')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath=Join-Path $root $Directory
function Json([string]$Path) {Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json}
function Hash([string]$Path) {(Get-FileHash -LiteralPath $Path).Hash}
$literal=Json (Join-Path $directoryPath 'SourceEquality.json')
$delta=Json (Join-Path $directoryPath 'ChangesFromPrevious.json')
$prior=Json (Join-Path $root 'docs/regression/GeometryFull_2026-10-13/PublishedSourceEquality.json')
if (-not $prior.AllPassed -or $prior.ActualComponents -ne 149 -or
    (Hash (Join-Path $directoryPath 'BeforeVBA.txt')) -cne '92C7B69F499564FD354721146E9E14B7E943488D190386355A42C9B4016B06B4' -or
    (Hash (Join-Path $directoryPath 'Before.xlsm')) -cne '871B42914FCB2FCF0168F43896D66BA603B6B170F6E6A9F9DB9206CDC3CA9FEC') {throw 'Previously accepted native reference is not confirmed.'}
$changed=@('CMaterialModelProvider','modTestBatchCalculation','modTestMaterialDiagrams','modTestWorkbookInterface')
$checks=@()
foreach ($check in $literal.Checks) {
    $native=@($delta.Checks | Where-Object Component -ceq $check.Component)
    $old=@($prior.Checks | Where-Object Component -ceq $check.Component)
    if ($native.Count -ne 1 -or $old.Count -ne 1) {throw 'Ambiguous source verification record.'}
    if ($check.Component -in $changed) {
        $passed=$check.Passed -and $check.RawTextSame -and -not $native[0].Passed
        $method='Current source equals actual native code'
    } else {
        $passed=$native[0].Passed -and $old[0].Passed -and $old[0].SourceSHA256 -ceq $check.SourceSHA256
        $method='Exact previous native code; unchanged independently verified source hash'
    }
    $checks+=,[ordered]@{Component=$check.Component;Passed=[bool]$passed;SourceSHA256=$check.SourceSHA256;Method=$method}
}
if ($literal.ActualComponents -ne 149 -or $checks.Count -ne 144 -or
    $delta.GeneratedDocumentChecks.Count -ne 5 -or @($delta.GeneratedDocumentChecks | Where-Object {-not $_.Passed}).Count -or
    @($checks | Where-Object {-not $_.Passed}).Count) {throw 'Native/source agreement failed.'}
[ordered]@{AllPassed=$true;ActualComponents=149;RepositoryComponents=144;DocumentComponents=5;
    ChangedComponents=$changed;Checks=$checks;GeneratedDocumentChecks=$delta.GeneratedDocumentChecks;
    WorkbookSHA256=(Hash (Join-Path $directoryPath 'RC_Section_NDM.xlsm'));
    DumpSHA256=(Hash (Join-Path $directoryPath 'VBA_All_Code.txt'));
    PriorVerificationSHA256=(Hash (Join-Path $root 'docs/regression/GeometryFull_2026-10-13/PublishedSourceEquality.json'))} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'NativeSourceEquality.json') -Encoding UTF8
Write-Output 'NATIVE_SOURCE_PASS: 4 changed components match source; 140 unchanged match accepted native/source; 5 document components unchanged.'
