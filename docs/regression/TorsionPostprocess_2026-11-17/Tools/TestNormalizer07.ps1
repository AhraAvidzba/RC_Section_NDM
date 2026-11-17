$ErrorActionPreference='Stop'
$p='C:\Users\avidzba\.codex\visualizations\2026\10\12\01a12ece-8192-7141-89b3-f83f1176184c\TorsionPostprocess_2026-11-17'
$script=[IO.File]::ReadAllText((Join-Path $p 'TestNativeCandidate.ps1'))
$match=[regex]::Match($script,"(?s)Add-Type -TypeDefinition @'\r?\n(.*?)\r?\n'@")
if(-not $match.Success){throw 'No actual source comparer'}
Add-Type -TypeDefinition $match.Groups[1].Value
$expected=[IO.File]::ReadAllText((Join-Path $p 'Native06/CLoadPathMath_Expected.txt'))
$native=[IO.File]::ReadAllText((Join-Path $p 'Native06/CLoadPathMath_Native.txt'))
$checks=[ordered]@{}
$checks.CLoadPathActualNativeEqual=[TestTorsionPostNativeIdentity]::Normalize($expected) -ceq [TestTorsionPostNativeIdentity]::Normalize($native)
function Test-Normalize([string]$t){return [TestTorsionPostNativeIdentity]::Normalize("Option Explicit`r`n"+$t)}
$checks.IdentifierCaseEqual=(Test-Normalize 'Dim Radius As Double') -ceq (Test-Normalize 'dim radius as double')
$checks.ExponentEqual=(Test-Normalize 'x=0.00000000000000177635683940025') -ceq (Test-Normalize 'x=1.77635683940025E-15')
$checks.DifferentNumbersRejected=(Test-Normalize 'x=1.1') -cne (Test-Normalize 'x=1.2')
$checks.StringCaseRejected=(Test-Normalize 'x="VALUE"') -cne (Test-Normalize 'x="value"')
$checks.CommentCaseRejected=(Test-Normalize "x=1 'Норма") -cne (Test-Normalize "x=1 'норма")
$checks.InlineCommentAlignmentEqual=(Test-Normalize "x=1 'Норма") -ceq (Test-Normalize "x=1        'Норма")
$checks.VbeLongLiteralCanonical=(Test-Normalize 'x=1.797693134862315E+308') -ceq (Test-Normalize 'x=1.79769313486231E+308')
$checks.DifferentFifteenthDigitRejected=(Test-Normalize 'x=1.23456789012345') -cne (Test-Normalize 'x=1.23456789012346')
$checks.FractionDoubleSuffixEqual=(Test-Normalize 'x=0.45#') -ceq (Test-Normalize 'x=0.45')
$checks.IntegerDoubleTypeRetained=(Test-Normalize 'x=1#') -cne (Test-Normalize 'x=1')
$pass=@($checks.Values|Where-Object{-not $_}).Count -eq 0
$result=[ordered]@{Passed=$pass;Checks=$checks}
$result|ConvertTo-Json -Depth 4|Set-Content -LiteralPath (Join-Path $p 'Normalizer07.json') -Encoding UTF8
$result|ConvertTo-Json -Depth 4
if(-not $pass){throw 'Source comparer regression failed'}

$aa=[TestTorsionPostNativeIdentity]::Normalize([IO.File]::ReadAllText((Join-Path $p 'Native11/CTorsionPostCalculator_Expected.txt'))).Split("`n")
$bb=[TestTorsionPostNativeIdentity]::Normalize([IO.File]::ReadAllText((Join-Path $p 'Native11/CTorsionPostCalculator_Native.txt'))).Split("`n")
$delta=@();for($i=0;$i -lt [Math]::Max($aa.Count,$bb.Count);$i++){if($aa[$i] -cne $bb[$i]){$delta+=,[pscustomobject]@{Line=$i+1;Expected=$aa[$i];Native=$bb[$i]};if($delta.Count -ge 5){break}}};$delta|ConvertTo-Json -Depth 3