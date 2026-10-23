# Сравнивает неизмененные baseline-сценарии на отдельных копиях двух книг.
# Счетчики внедряются только в временные VBA-проекты и не меняют исходники.
param(
    [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
    [Parameter(Mandatory=$true)][string]$CurrentWorkbook,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [int]$Repeats = 3
)
$ErrorActionPreference = "Stop"
$sources = @((Resolve-Path -LiteralPath $BaselineWorkbook).Path, (Resolve-Path -LiteralPath $CurrentWorkbook).Path)
$hashes = @($sources | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash })
$folder = Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_Performance_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $folder | Out-Null
$report = [IO.Path]::GetFullPath($ReportPath)
$lines = New-Object System.Collections.Generic.List[string]
$failed = $false
$excel = $null
$book = $null

# Добавляет счетчик перед единственным определенным оператором, сохраняя математику.
function Insert-Counter {
    param([object]$Module, [string]$Pattern, [string]$Call, [switch]$After)
    $code = $Module.Lines(1, $Module.CountOfLines)
    $matches = [regex]::Matches($code, $Pattern, [Text.RegularExpressions.RegexOptions]::Multiline)
    if ($matches.Count -ne 1) { throw "Неоднозначная точка счетчика: $Pattern ($($matches.Count))." }
    $position = $matches[0].Index
    if ($After) { $position += $matches[0].Length }
    $replacement = $code.Insert($position, "    $Call`r`n")
    $Module.DeleteLines(1, $Module.CountOfLines)
    $Module.AddFromString($replacement)
}

# Вызывает существующие private-тесты, сохраняя их физические expected/tolerance.
function Add-BenchmarkEntry {
    param([object]$Module, [string]$StatsType, [hashtable]$Cases)
    $branches = ($Cases.Keys | Sort-Object | ForEach-Object { "        Case `"$_`": $($Cases[$_]) stats" }) -join "`r`n"
    $entry = @"
' ДЛЯ ТЕСТОВ: временный benchmark использует штатные baseline assertions.
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет историческую проверку вместе со штатными assertions и счётчиками частной benchmark-книги.
Public Function TestRunAudit02PerfCase(ByVal caseName As String) As String
    On Error GoTo Failed
    Dim stats As $StatsType
    Dim started As Double
    ResetSectionEquilibriumSolveCount
    TestAudit02ResetCounters
    started = Timer
    Select Case caseName
$branches
        Case Else: Err.Raise 5, "Audit02Perf", "Неизвестный сценарий"
    End Select
    TestRunAudit02PerfCase = "case=" & caseName & "; passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; elapsedSec=" & CStr(Timer - started) & _
        "; solves=" & CStr(SectionEquilibriumSolveCount()) & "; probes=" & CStr(Audit02ProbeCount) & _
        "; planeProbes=" & CStr(Audit02PlaneCount) & "; probeReuse=" & CStr(Audit02ProbeReuseCount) & _
        "; stateReuse=" & CStr(Audit02StateReuseCount) & "; unconvergedProbes=" & CStr(Audit02UnconvergedCount) & _
        "; unconfirmedPhysical=" & CStr(Audit02UnconfirmedLimitCount)
    If stats.Failed > 0 Then TestRunAudit02PerfCase = TestRunAudit02PerfCase & vbCrLf & stats.Report
    Exit Function
Failed:
    TestRunAudit02PerfCase = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
"@
    $Module.AddFromString($entry)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    for ($version = 0; $version -lt 2; $version++) {
        $label = @("baseline", "current")[$version]
        $fixture = Join-Path $folder ("$label.xlsm")
        Copy-Item -LiteralPath $sources[$version] -Destination $fixture
        $book = $excel.Workbooks.Open($fixture)
        $settings = $book.Names.Item("rngSystemSettings").RefersToRange
        $data = $settings.Value2
        $key = @("Solver.DirectState.DiagramExtension", "General.DiagramExtension")[$version]
        $found = $false
        for ($row = 2; $row -le $settings.Rows.Count; $row++) {
            if ([string]$data[$row, 1] -eq $key) { $settings.Cells.Item($row, 2).Value2 = "No"; $found = $true }
        }
        if (-not $found) { throw "Отсутствует настройка $key." }
        $counter = $book.VBProject.VBComponents.Add(1)
        $counter.Name = "modTestAudit02PerfCounters"
        $counter.CodeModule.AddFromString(@'
Option Explicit
' ==========================================================================
' ДЛЯ ТЕСТОВ — модуль только частной диагностической книги.
' ==========================================================================
' ДЛЯ ТЕСТОВ: счетчики только этой временной benchmark-книги.
Public Audit02ProbeCount As Long
Public Audit02ProbeReuseCount As Long
Public Audit02StateReuseCount As Long
Public Audit02PlaneCount As Long
Public Audit02UnconvergedCount As Long
Public Audit02UnconfirmedLimitCount As Long
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сбрасывает диагностические счётчики частного benchmark перед независимым сценарием.
Public Sub TestAudit02ResetCounters()
    Audit02ProbeCount = 0
    Audit02ProbeReuseCount = 0
    Audit02StateReuseCount = 0
    Audit02PlaneCount = 0
    Audit02UnconvergedCount = 0
    Audit02UnconfirmedLimitCount = 0
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает физические probe только инструментированной benchmark-копии.
Public Sub TestAudit02RecordProbe()
    Audit02ProbeCount = Audit02ProbeCount + 1
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает восстановленные probe отдельно от новых вычислений.
Public Sub TestAudit02RecordProbeReuse()
    Audit02ProbeReuseCount = Audit02ProbeReuseCount + 1
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает восстановленные состояния отдельно от новых решений.
Public Sub TestAudit02RecordStateReuse()
    Audit02StateReuseCount = Audit02StateReuseCount + 1
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает вычисления пробной плоскости деформаций в инструментированной копии.
Public Sub TestAudit02RecordPlane()
    Audit02PlaneCount = Audit02PlaneCount + 1
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разделяет несходимость probe и неподтверждённое равновесием достижение предельной деформации.
Public Sub TestAudit02RecordOutcome(ByVal solver As CSectionSolver, ByVal state As String)
    If solver Is Nothing Then Exit Sub
    If Not solver.Converged Then
        Audit02UnconvergedCount = Audit02UnconvergedCount + 1
        If state = "ConcreteStrainLimit" Or state = "ConcreteTensionStrainLimit" Or state = "SteelStrainLimit" Then
            Audit02UnconfirmedLimitCount = Audit02UnconfirmedLimitCount + 1
        End If
    End If
End Sub
'@
        )
        $capacity = $book.VBProject.VBComponents.Item("CCapacitySolver").CodeModule
        Insert-Counter $capacity '^    If TryRestoreProbeCache\(' 'TestAudit02RecordProbe'
        Insert-Counter $capacity '^            TryRestoreProbeCache = True' 'TestAudit02RecordProbeReuse'
        Insert-Counter $capacity '^    AppendDiagnostic solver, mx, my, lambdaValue, attempt, ProbeOnce' 'TestAudit02RecordOutcome solver, ProbeOnce'
        Insert-Counter $capacity '(?m)^Public Function LimitSearchEvaluateUltimateResidual\([\s\S]*?\) As Boolean\r?\n' 'TestAudit02RecordPlane' -After
        $formationName = @("CCrackWidthCalculator", "CCrackFormationCalculator")[$version]
        $formation = $book.VBProject.VBComponents.Item($formationName).CodeModule
        Insert-Counter $formation '(?m)^Public Function LimitSearchEvaluateCrackFormationLoadMultiplier\([\s\S]*?\) As Boolean\r?\n' 'TestAudit02RecordProbe' -After
        $provider = $book.VBProject.VBComponents.Item("CStateProvider").CodeModule
        Insert-Counter $provider '^            mLastStateWasReused = True' 'TestAudit02RecordStateReuse'
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestCapacitySolver").CodeModule "TCapacityTestStats" @{
            bending = "TestMxPositiveAndNegative"
            biaxial = "TestMxySignedCombinations"
            asymmetric = "TestAsymmetricCoupledCurvatures"
            axial = "TestAxialLoadMultiplierFindsNult"
            methods = "TestLoadMultiplierSearchMethods"
        }
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestCrackWidth").CodeModule "TCrackTestStats" @{
            crack_bending = "TestAutoMcrcPureBendingConverges"
            crack_paths = "TestCrackInitiationLoadPaths"
            crack_cache = "TestCrackFormationCacheHitWithoutLastRunner"
        }
        $lines.Add("SOURCE|$label|$($sources[$version])|sha256=$($hashes[$version])")
        for ($run = 1; $run -le $Repeats; $run++) {
            foreach ($case in @("bending", "biaxial", "asymmetric", "axial", "methods", "crack_bending", "crack_paths", "crack_cache")) {
                $module = "modTestCapacitySolver"
                if ($case.StartsWith("crack_")) { $module = "modTestCrackWidth" }
                $lines.Add("START|$label|run=$run|case=$case")
                $lines | Set-Content -LiteralPath $report -Encoding UTF8
                $result = [string]$excel.Run("'$($book.Name)'!$module.TestRunAudit02PerfCase", $case)
                $lines.Add("PERF|$label|run=$run|$result")
                if ($result -match 'failed=([1-9][0-9]*)|RUNTIME ERROR') { $failed = $true }
            }
        }
        $book.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
        $book = $null
    }
    for ($i = 0; $i -lt 2; $i++) {
        if ((Get-FileHash -LiteralPath $sources[$i] -Algorithm SHA256).Hash -ne $hashes[$i]) { throw "Изменилась исходная книга." }
    }
    $lines.Add("SOURCES_UNCHANGED: True")
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
} finally {
    if ($book) { try { $book.Close($false) } catch {} }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $lines | Set-Content -LiteralPath $report -Encoding UTF8
    $lines
}
if ($failed) { exit 1 }
