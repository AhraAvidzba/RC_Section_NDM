# Создает воспроизводимую covering-array для действующих расчетных selectors.
# Проверяет все пары допустимых значений и явный high-risk Cartesian subset.
# Это план runtime cases, не PASS расчетов; общий Search/State не подменяется.
param(
    [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputPath))
if (-not $output.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Matrix must stay in docs/regression/Audit03.' }
if (Test-Path -LiteralPath $output) { throw 'Matrix evidence already exists; choose a new path.' }
$factors = @(
    @{ Name='Shape'; Values=@('CircleUneven','RectL','RoundedMixed','HollowOffset') },
    @{ Name='Extension'; Values=@('No','Yes') },
    @{ Name='CapacityStrategy'; Values=@('Auto','UltimateStrain','LoadMultiplier') },
    @{ Name='FormationStrategy'; Values=@('Auto','UltimateStrain','LoadMultiplier') },
    @{ Name='SearchMethod'; Values=@('Bisection','Brent','Secant') },
    @{ Name='SolverMethod'; Values=@('Newton','Secant') },
    @{ Name='ConcreteDiagram'; Values=@('TwoLine','ThreeLine') },
    @{ Name='SteelDiagram'; Values=@('TwoLine','ThreeLine') },
    @{ Name='StrengthTension'; Values=@('Ignore','UseDiagram') },
    @{ Name='PsiMode'; Values=@('User','Auto','AlwaysCalc') }
)
$required = [Collections.Generic.HashSet[string]]::new()
$covered = [Collections.Generic.HashSet[string]]::new()
$pairs = [Collections.Generic.List[object]]::new()
$cases = [Collections.Generic.List[object]]::new()
$caseKeys = @{}
for ($a=0; $a -lt $factors.Count; $a++) {
    for ($b=$a+1; $b -lt $factors.Count; $b++) {
        for ($av=0; $av -lt $factors[$a].Values.Count; $av++) {
            for ($bv=0; $bv -lt $factors[$b].Values.Count; $bv++) {
                $key = "$a/$av/$b/$bv"
                [void]$required.Add($key)
                $pairs.Add(@{ A=$a; AV=$av; B=$b; BV=$bv; Key=$key })
            }
        }
    }
}

# Однозначно кодирует индексы факторов. Генератор работает с конфигурацией,
# а не с физическими формулами; каждое созданное значение входит в каталог.
function Get-PairKeys([int[]]$Vector) {
    $result = [Collections.Generic.List[string]]::new()
    for ($a=0; $a -lt $Vector.Count; $a++) {
        for ($b=$a+1; $b -lt $Vector.Count; $b++) {
            $result.Add("$a/$($Vector[$a])/$b/$($Vector[$b])")
        }
    }
    return ,$result.ToArray()
}

# Хранит high-risk происхождение даже если тот же tuple уже выбран pairwise.
# При дедупликации сохраняется один runtime case с двумя coverage labels.
function Add-MatrixCase([int[]]$Vector, [string]$Scope) {
    $key = $Vector -join '/'
    if ($caseKeys.ContainsKey($key)) {
        if (-not $caseKeys[$key].Scopes.Contains($Scope)) { $caseKeys[$key].Scopes.Add($Scope) }
        return
    }
    $values = [ordered]@{}
    for ($i=0; $i -lt $factors.Count; $i++) { $values[$factors[$i].Name] = $factors[$i].Values[$Vector[$i]] }
    $scopes = [Collections.Generic.List[string]]::new()
    $scopes.Add($Scope)
    $item = [pscustomobject]@{ Id=('Selector_' + ($cases.Count + 1).ToString('D3')); Indices=$Vector.Clone(); Values=$values; Scopes=$scopes; RuntimeStatus='NotRun' }
    $cases.Add($item)
    $caseKeys[$key] = $item
    foreach ($pairKey in (Get-PairKeys $Vector)) { [void]$covered.Add($pairKey) }
}

# Явные четверки: форма x extension x capacity strategy x tensile concrete.
# Остальные selectors меняются детерминированно, а не всегда остаются default.
$serial = 0
for ($shape=0; $shape -lt 4; $shape++) {
    for ($extension=0; $extension -lt 2; $extension++) {
        for ($strategy=0; $strategy -lt 3; $strategy++) {
            for ($tension=0; $tension -lt 2; $tension++) {
                $vector = [int[]]@(0,0,0,0,0,0,0,0,0,0)
                for ($i=0; $i -lt $factors.Count; $i++) { $vector[$i] = ($serial + $i) % $factors[$i].Values.Count }
                $vector[0]=$shape; $vector[1]=$extension; $vector[2]=$strategy; $vector[8]=$tension
                Add-MatrixCase $vector 'HighRisk:Shape-Extension-CapacityStrategy-StrengthTension'
                $serial++
            }
        }
    }
}

# Из каждой непокрытой пары строим admissible candidate и выбираем тот,
# который закроет больше всего оставшихся пар. Равенство разрешается стабильным
# порядком исходного списка; случайный seed и полный многомерный перебор не нужны.
while ($covered.Count -lt $required.Count) {
    $best = $null
    $bestScore = 0
    foreach ($pair in $pairs) {
        if ($covered.Contains($pair.Key)) { continue }
        for ($rotation=0; $rotation -lt 3; $rotation++) {
            $vector = New-Object 'int[]' $factors.Count
            for ($i=0; $i -lt $factors.Count; $i++) { $vector[$i] = ($serial + $i + $rotation) % $factors[$i].Values.Count }
            $vector[$pair.A]=$pair.AV; $vector[$pair.B]=$pair.BV
            $score = 0
            foreach ($key in (Get-PairKeys $vector)) { if (-not $covered.Contains($key)) { $score++ } }
            if ($score -gt $bestScore) { $bestScore=$score; $best=$vector.Clone() }
        }
    }
    if ($null -eq $best -or $bestScore -eq 0) { throw 'Covering-array generation stagnated before full pair coverage.' }
    Add-MatrixCase $best 'Pairwise'
    $serial++
}

# Независимо повторно проверяем полноту по опубликованным cases, не по счетчику
# основного greedy-цикла. Runtime acceptance остается отдельным последним gate.
$verified = [Collections.Generic.HashSet[string]]::new()
foreach ($case in $cases) { foreach ($key in (Get-PairKeys $case.Indices)) { [void]$verified.Add($key) } }
$missing = @($required | Where-Object { -not $verified.Contains($_) })
if ($missing.Count) { throw 'Published cases do not cover all required pairs.' }
$result = [ordered]@{
    Factors=$factors; RequiredPairs=$required.Count; CoveredPairs=$verified.Count
    HighRiskQuads=48; CaseCount=$cases.Count; RuntimeAcceptance='NotRun'
    Cases=$cases
    FixedContracts=@('CrackInitiation.ConcreteTension=UseDiagram','CrackedState.ConcreteTension=Ignore','Each case must execute all six Capacity and six Formation load paths','This plan does not replace Light/Stress 13-shape matrix or near-limit gates')
}
$result | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath $output -Encoding UTF8
Write-Output "SELECTOR_MATRIX_PLAN: cases=$($cases.Count); requiredPairs=$($required.Count); coveredPairs=$($verified.Count); highRiskQuads=48; runtime=NotRun"
