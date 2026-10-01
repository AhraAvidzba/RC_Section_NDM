# Снимает воспроизводимые структурные метрики baseline/current без правки исходников.
# Метрики не заменяют ручной аудит ответственности и не являются целевыми квотами.
param(
    [Parameter(Mandatory=$true)][string]$BaselineRoot,
    [string]$CurrentRoot = '.',
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$roots = @((Resolve-Path -LiteralPath $BaselineRoot).Path, (Resolve-Path -LiteralPath $CurrentRoot).Path)
$names = @('CBatchSectionCalculator', 'CCrackWidthCalculator', 'CCapacitySolver', 'CCombinationResult',
    'CCrackFormationCalculator', 'CExecutionReport', 'CLimitSearchCoordinator', 'CLoadMultiplierSearch',
    'CUltimateStrainSearch', 'CLimitSearchResult', 'ILimitSearchProblem', 'CDirectStateResult', 'CStateProvider', 'CStateRepository')
$rows = @()
for ($version = 0; $version -lt 2; $version++) {
    $files = @(Get-ChildItem -LiteralPath (Join-Path $roots[$version] 'src') -Recurse -File -Filter '*.cls')
    foreach ($name in $names) {
        $file = @($files | Where-Object BaseName -eq $name)
        if ($file.Count -ne 1) { throw "Неоднозначный файл $name." }
        $code = [IO.File]::ReadAllText($file[0].FullName)
        $methods = [regex]::Matches($code, '(?m)^(?:Public|Private|Friend) (?:Sub|Function|Property (?:Get|Let|Set))\b')
        $indexedScalar = [regex]::Matches($code, '(?m)^Public (?:Function|Property Get) \w+\(ByVal index As Long\) As (?:Double|Long|Boolean|String)\b')
        $scalarFields = [regex]::Matches($code, '(?m)^Public \w+ As (?:Double|Long|Boolean|String)\b')
        $assignments = [regex]::Matches($code, '(?m)^\s*(?:Set )?(?:m\w+|[A-Z]\w+)(?:\([^\r\n]*?\))?(?:\.\w+)? = (?:[a-z]\w+\.(?:\w+\.)?\w+)')
        $dependencies = @([regex]::Matches($code, '\bAs (?:New )?(C[A-Z]\w+|I[A-Z]\w+)\b') |
            ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        $rows += [pscustomobject]@{
            Version = @('baseline', 'current')[$version]
            Class = $name
            Lines = ($code -split '\r?\n').Count
            Methods = $methods.Count
            IndexedScalarAPI = $indexedScalar.Count
            PublicScalarFields = $scalarFields.Count
            ObjectReadAssignments = $assignments.Count
            Dependencies = $dependencies.Count
            DependencyNames = $dependencies -join ','
        }
    }
}
$report = [IO.Path]::GetFullPath($ReportPath)
$rows | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $report -Encoding UTF8
$rows | Format-Table Version,Class,Lines,Methods,IndexedScalarAPI,PublicScalarFields,ObjectReadAssignments,Dependencies -AutoSize
$baseline = @(Get-ChildItem -LiteralPath (Join-Path $roots[0] 'src') -Recurse -File -Filter '*.cls' | ForEach-Object BaseName)
$current = @(Get-ChildItem -LiteralPath (Join-Path $roots[1] 'src') -Recurse -File -Filter '*.cls' | ForEach-Object BaseName)
Compare-Object $baseline $current
