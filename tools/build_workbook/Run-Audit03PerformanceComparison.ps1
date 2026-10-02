# Сравнивает корректный baseline Audit03 и текущий код на отдельных копиях.
# Пять повторов минимум; прежний быстрый ошибочный pre-Audit02 код не используется.
# Счетчики внедряются только в временные VBA-проекты и не меняют исходники.
param(
    [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
    [Parameter(Mandatory=$true)][string]$CurrentWorkbook,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [ValidateRange(5, 50)][int]$Repeats = 5,
    [switch]$InternalWorker,
    [string[]]$Cases = @()
)
$ErrorActionPreference = "Stop"
# Внешний процесс ограничивает COM-вызовы: compile dialog или зависание не
# оставляет бесконечный benchmark. Закрываются только новые test Excel PID.
if (-not $InternalWorker) {
    $root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
    $beforeExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $report = [IO.Path]::GetFullPath($ReportPath)
    $quoted = @($PSCommandPath, $BaselineWorkbook, $CurrentWorkbook, $ReportPath)
    if (@($quoted | Where-Object { $_.Contains('"') }).Count) { throw 'Unsupported quote in benchmark path.' }
    $arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $PSCommandPath + '" -BaselineWorkbook "' + $BaselineWorkbook +
        '" -CurrentWorkbook "' + $CurrentWorkbook + '" -ReportPath "' + $ReportPath + '" -Repeats ' + $Repeats + ' -InternalWorker'
    if ($Cases.Count) {
        $Cases = @($Cases | ForEach-Object { $_ -split ',' })
        if (@($Cases | Where-Object { $_ -notmatch '^[a-z_]+$' }).Count) { throw 'Invalid benchmark case name.' }
        $arguments += ' -Cases ' + ($Cases -join ',')
    }
    $process = Start-Process -FilePath (Join-Path $PSHOME 'powershell.exe') -ArgumentList $arguments `
        -WorkingDirectory $root -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput ($report + '.stdout.log') -RedirectStandardError ($report + '.stderr.log')
    $null = $process.Handle
    if (-not $process.WaitForExit(1800000)) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $beforeExcel -notcontains $_.Id } | Stop-Process -Force
        Add-Content -LiteralPath $report -Encoding UTF8 -Value 'WATCHDOG_FAILURE: benchmark exceeded 1800 seconds.'
        exit 1
    }
    $process.WaitForExit()
    Write-Output "BENCHMARK_COMPLETED: exit=$($process.ExitCode); report=$report"
    exit $process.ExitCode
}
$sources = @((Resolve-Path -LiteralPath $BaselineWorkbook).Path, (Resolve-Path -LiteralPath $CurrentWorkbook).Path)
$hashes = @($sources | ForEach-Object { (Get-FileHash -LiteralPath $_ -Algorithm SHA256).Hash })
$folder = Join-Path ([IO.Path]::GetTempPath()) ("RC_NDM_Performance_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $folder | Out-Null
$report = [IO.Path]::GetFullPath($ReportPath)
$lines = New-Object System.Collections.Generic.List[string]
$failed = $false
$excel = $null
$book = $null
$lines.Add("ENVIRONMENT|date=$([DateTime]::Now.ToString('s'))|powershell=$($PSVersionTable.PSVersion)|os=$([Environment]::OSVersion)|repeats=$Repeats")
$lines.Add("GIT|$(& git rev-parse HEAD)|dirty=$([bool](& git status --porcelain))")

# Короткий совместимый доступ к журналу не должен прерывать расчет при чтении
# прогресса. Повторяется только конфликт sharing/lock; другие I/O ошибки видимы.
function Save-PerformanceReport {
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $stream = $null
        $writer = $null
        try {
            $stream = [IO.File]::Open($report, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
            $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($true)))
            foreach ($line in $lines) { $writer.WriteLine($line) }
            $writer.Flush()
            return
        } catch [IO.IOException] {
            $code = $_.Exception.HResult -band 65535
            if ($code -notin @(32, 33) -or $attempt -eq 19) { throw }
            Start-Sleep -Milliseconds 50
        } finally {
            if ($writer) { $writer.Dispose() }
            elseif ($stream) { $stream.Dispose() }
        }
    }
}

# Добавляет счетчик перед единственным определенным оператором, сохраняя математику.
function Insert-Counter {
    param([object]$Module, [string]$Pattern, [string]$Call, [switch]$After, [int]$ExpectedMatches = 1)
    $code = $Module.Lines(1, $Module.CountOfLines)
    $matches = [regex]::Matches($code, $Pattern, [Text.RegularExpressions.RegexOptions]::Multiline -bor [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($matches.Count -ne $ExpectedMatches) { throw "Неоднозначная точка счетчика: $Pattern ($($matches.Count))." }
    $replacement = $code
    for ($index = $matches.Count - 1; $index -ge 0; $index--) {
        $position = $matches[$index].Index
        if ($After) { $position += $matches[$index].Length }
        $replacement = $replacement.Insert($position, "    $Call`r`n")
    }
    $Module.DeleteLines(1, $Module.CountOfLines)
    $Module.AddFromString($replacement)
}

# Измеряет весь вызов метода, включая ранний выход. Только временная копия VBA;
# численный порядок и содержимое условия не меняются, учет не дублируется.
function Add-DurationCounter {
    param([object]$Module, [string]$Method, [string]$Category)
    $code = $Module.Lines(1, $Module.CountOfLines)
    $pattern = '(?ims)^Public Sub ' + [regex]::Escape($Method) + '\([\s\S]*?^End Sub\r?$'
    $matches = [regex]::Matches($code, $pattern)
    if ($matches.Count -ne 1) { throw "Duration method not unique: $Method" }
    $methodCode = $matches[0].Value
    $header = [regex]::Match($methodCode, '(?is)^Public Sub [\s\S]*?\)\r?\n')
    if (-not $header.Success) { throw "Duration header missing: $Method" }
    $methodCode = $methodCode.Insert($header.Length, "    Dim audit03Started As Double`r`n    audit03Started = Timer`r`n")
    $record = "Audit03RecordDuration `"$Category`", Timer - audit03Started"
    $methodCode = $methodCode.Replace('Exit Sub', "$record`: Exit Sub")
    $methodCode = [regex]::Replace($methodCode, '(?m)^End Sub', "    $record`r`nEnd Sub")
    $replacement = $code.Remove($matches[0].Index, $matches[0].Length).Insert($matches[0].Index, $methodCode)
    $Module.DeleteLines(1, $Module.CountOfLines)
    $Module.AddFromString($replacement)
}

# Вызывает существующие private-тесты, сохраняя их физические expected/tolerance.
function Add-BenchmarkEntry {
    param([object]$Module, [string]$StatsType, [hashtable]$Cases)
    $branches = ($Cases.Keys | Sort-Object | ForEach-Object { "        Case `"$_`": $($Cases[$_]) stats" }) -join "`r`n"
    $entry = @"
' ДЛЯ ТЕСТОВ: временный benchmark использует штатные baseline assertions.
Public Function RunAudit03PerfCase(ByVal caseName As String) As String
    On Error GoTo Failed
    Dim stats As $StatsType
    Dim started As Double
    ResetSectionEquilibriumSolveCount
    Audit03ResetCounters
    Audit03DisableDiagnostics = (caseName = "diagnostics_off")
    started = Timer
    Dim cycle As Long, cycles As Long
    cycles = 1
    If caseName = "geometry_hollow" Or caseName = "geometry_tapered" Then cycles = 25
    For cycle = 1 To cycles
    Select Case caseName
$branches
        Case Else: Err.Raise 5, "Audit03Perf", "Неизвестный сценарий"
    End Select
    Next cycle
    RunAudit03PerfCase = "case=" & caseName & "; passed=" & CStr(stats.Passed) & _
        "; failed=" & CStr(stats.Failed) & "; cycles=" & CStr(cycles) & "; elapsedSec=" & CStr(Timer - started) & _
        "; solves=" & CStr(SectionEquilibriumSolveCount()) & "; probes=" & CStr(Audit03ProbeCount) & _
        "; planeProbes=" & CStr(Audit03PlaneCount) & "; probeReuse=" & CStr(Audit03ProbeReuseCount) & _
        "; stateReuse=" & CStr(Audit03StateReuseCount) & "; unconvergedProbes=" & CStr(Audit03UnconvergedCount) & _
        "; unconfirmedPhysical=" & CStr(Audit03UnconfirmedLimitCount) & _
        "; newtonCalls=" & CStr(Audit03NewtonCallCount) & "; integrations=" & CStr(Audit03IntegrationCount) & _
        "; iterations=" & CStr(Audit03IterationCount) & "; retries=" & CStr(Audit03RetryCount) & _
        "; duplicateSolveAttempts=" & CStr(Audit03DuplicateSolveAttempts) & _
        "; retryFormatCalls=" & CStr(Audit03RetryFormatCalls) & _
        "; geometrySec=" & CStr(Audit03GeometrySeconds) & "; coreSec=" & CStr(Audit03CoreSeconds) & _
        "; packagingSec=" & CStr(Audit03PackagingSeconds) & "; writeSec=" & CStr(Audit03WriteSeconds) & _
        "; modelSizes=" & Audit03ModelSizes
    RunAudit03PerfCase = RunAudit03PerfCase & vbCrLf & stats.Report & Audit03DuplicateLog
    Exit Function
Failed:
    RunAudit03PerfCase = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
"@
    $Module.AddFromString($entry)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $lines.Add("EXCEL|version=$($excel.Version)|build=$($excel.Build)|operatingSystem=$($excel.OperatingSystem)")
    for ($version = 0; $version -lt 2; $version++) {
        $label = @("baseline", "current")[$version]
        $fixture = Join-Path $folder ("$label.xlsm")
        Copy-Item -LiteralPath $sources[$version] -Destination $fixture
        $book = $excel.Workbooks.Open($fixture)
        $settings = $book.Names.Item("rngSystemSettings").RefersToRange
        $data = $settings.Value2
        $key = "General.DiagramExtension"
        $found = $false
        for ($row = 2; $row -le $settings.Rows.Count; $row++) {
            if ([string]$data[$row, 1] -eq $key) { $settings.Cells.Item($row, 2).Value2 = "No"; $found = $true }
        }
        if (-not $found) { throw "Отсутствует настройка $key." }
        $counter = $book.VBProject.VBComponents.Add(1)
        $counter.Name = "modAudit03PerfCounters"
        $counter.CodeModule.AddFromString(@'
Option Explicit
' ДЛЯ ТЕСТОВ: счетчики только этой временной benchmark-книги.
Public Audit03ProbeCount As Long
Public Audit03ProbeReuseCount As Long
Public Audit03StateReuseCount As Long
Public Audit03PlaneCount As Long
Public Audit03UnconvergedCount As Long
Public Audit03UnconfirmedLimitCount As Long
Public Audit03NewtonCallCount As Long
Public Audit03IntegrationCount As Long
Public Audit03IterationCount As Long
Public Audit03RetryCount As Long
Public Audit03GeometrySeconds As Double
Public Audit03CoreSeconds As Double
Public Audit03PackagingSeconds As Double
Public Audit03WriteSeconds As Double
Public Audit03ModelSizes As String
Public Audit03DuplicateSolveAttempts As Long
Private Audit03SolveAttempts As Collection
Private Audit03SolveAttemptNumbers As Collection
Private Audit03SolveAttemptSequence As Long
Public Audit03DuplicateLog As String
Public Audit03DisableDiagnostics As Boolean
Public Audit03RetryFormatCalls As Long
Public Sub Audit03ResetCounters()
    Audit03ProbeCount = 0
    Audit03ProbeReuseCount = 0
    Audit03StateReuseCount = 0
    Audit03PlaneCount = 0
    Audit03UnconvergedCount = 0
    Audit03UnconfirmedLimitCount = 0
    Audit03NewtonCallCount = 0
    Audit03IntegrationCount = 0
    Audit03IterationCount = 0
    Audit03RetryCount = 0
    Audit03GeometrySeconds = 0
    Audit03CoreSeconds = 0
    Audit03PackagingSeconds = 0
    Audit03WriteSeconds = 0
    Audit03ModelSizes = vbNullString
    Audit03DuplicateSolveAttempts = 0
    Set Audit03SolveAttempts = New Collection
    Set Audit03SolveAttemptNumbers = New Collection
    Audit03SolveAttemptSequence = 0
    Audit03DuplicateLog = vbNullString
    Audit03DisableDiagnostics = False
    Audit03RetryFormatCalls = 0
End Sub
' ДЛЯ ТЕСТОВ: сравнение scalar options использует точные Double, без string-key.
Public Sub Audit03RecordSolveAttempt(ByVal attempt As Variant)
    Dim previous As Variant, index As Long, equal As Boolean
    Dim previousIndex As Long
    Audit03SolveAttemptSequence = Audit03SolveAttemptSequence + 1
    For Each previous In Audit03SolveAttempts
        previousIndex = previousIndex + 1
        equal = True
        For index = LBound(attempt) To UBound(attempt)
            If previous(index) <> attempt(index) Then equal = False: Exit For
        Next index
        If equal Then
            Audit03DuplicateSolveAttempts = Audit03DuplicateSolveAttempts + 1
            Audit03DuplicateLog = Audit03DuplicateLog & "DUPLICATE: previous=" & CStr(Audit03SolveAttemptNumbers(previousIndex)) & _
                "; current=" & CStr(Audit03SolveAttemptSequence) & "; N=" & CStr(attempt(3)) & _
                "; Mx=" & CStr(attempt(4)) & "; My=" & CStr(attempt(5)) & "; loadSteps=" & CStr(attempt(7)) & _
                "; lineSearch=" & CStr(attempt(11)) & "; maxDeltaKappa=" & CStr(attempt(15)) & _
                "; eps0=" & CStr(attempt(20)) & "; kx=" & CStr(attempt(21)) & "; ky=" & CStr(attempt(22)) & _
                "; initialN=" & CStr(attempt(24)) & "; initialMx=" & CStr(attempt(25)) & "; initialMy=" & CStr(attempt(26)) & vbCrLf
            Exit Sub
        End If
    Next previous
    Audit03SolveAttempts.Add attempt
    Audit03SolveAttemptNumbers.Add Audit03SolveAttemptSequence
End Sub
Public Sub Audit03RecordIteration()
    Audit03IterationCount = Audit03IterationCount + 1
End Sub
Public Sub Audit03RecordRetryFormat()
    Audit03RetryFormatCalls = Audit03RetryFormatCalls + 1
End Sub
Public Sub Audit03RecordRetry()
    Audit03RetryCount = Audit03RetryCount + 1
End Sub
Public Sub Audit03RecordDuration(ByVal category As String, ByVal seconds As Double)
    If seconds < 0 Then seconds = seconds + 86400#
    Select Case category
        Case "geometry": Audit03GeometrySeconds = Audit03GeometrySeconds + seconds
        Case "core": Audit03CoreSeconds = Audit03CoreSeconds + seconds
        Case "packaging": Audit03PackagingSeconds = Audit03PackagingSeconds + seconds
        Case "write": Audit03WriteSeconds = Audit03WriteSeconds + seconds
    End Select
End Sub
Public Sub Audit03RecordModel(ByVal mesh As CFiberMeshBuilder, ByVal rebars As CRebarLayout)
    Dim sizeText As String
    sizeText = CStr(mesh.FiberCount) & "f/"
    If rebars Is Nothing Then
        sizeText = sizeText & "0r"
    Else
        sizeText = sizeText & CStr(rebars.Count) & "r"
    End If
    If InStr(1, "|" & Audit03ModelSizes & "|", "|" & sizeText & "|", vbBinaryCompare) = 0 Then
        If Len(Audit03ModelSizes) > 0 Then Audit03ModelSizes = Audit03ModelSizes & "|"
        Audit03ModelSizes = Audit03ModelSizes & sizeText
    End If
End Sub
Public Sub Audit03RecordNewtonCall()
    Audit03NewtonCallCount = Audit03NewtonCallCount + 1
End Sub
Public Sub Audit03RecordIntegration()
    Audit03IntegrationCount = Audit03IntegrationCount + 1
End Sub
Public Sub Audit03RecordProbe()
    Audit03ProbeCount = Audit03ProbeCount + 1
End Sub
Public Sub Audit03RecordProbeReuse()
    Audit03ProbeReuseCount = Audit03ProbeReuseCount + 1
End Sub
Public Sub Audit03RecordStateReuse()
    Audit03StateReuseCount = Audit03StateReuseCount + 1
End Sub
Public Sub Audit03RecordPlane()
    Audit03PlaneCount = Audit03PlaneCount + 1
End Sub
Public Sub Audit03RecordOutcome(ByVal solver As CSectionSolver, ByVal state As String)
    If solver Is Nothing Then Exit Sub
    If Not solver.Converged Then
        Audit03UnconvergedCount = Audit03UnconvergedCount + 1
        If state = "ConcreteStrainLimit" Or state = "ConcreteTensionStrainLimit" Or state = "SteelStrainLimit" Then
            Audit03UnconfirmedLimitCount = Audit03UnconfirmedLimitCount + 1
        End If
    End If
End Sub
'@
        )
        $solver = $book.VBProject.VBComponents.Item("CSectionSolver").CodeModule
        Insert-Counter $solver '(?m)^Private Function SolveNewtonLoadStep\([\s\S]*?\) As Boolean\r?\n' 'Audit03RecordNewtonCall' -After
        Insert-Counter $solver '(?m)^Private Sub EvaluateStateInto\([\s\S]*?\)\r?\n' 'Audit03RecordIntegration' -After
        Insert-Counter $solver '^        mIterations = mIterations \+ 1' 'Audit03RecordIteration' -ExpectedMatches 2
        Insert-Counter $solver '^    RecordSectionEquilibriumSolve\r?$' @'
Audit03RecordSolveAttempt Array(ObjPtr(section), ObjPtr(concreteMaterial), ObjPtr(steelMaterial), _
    targetN, targetMx, targetMy, mMaxIterations, mLoadSteps, mToleranceN, mToleranceMx, mToleranceMy, _
    mLineSearchEnabled, mDampingInitial, mMinLineSearchAlpha, mMaxDeltaEpsilon0, mMaxDeltaKappa, _
    mSolverMethod, mSecantMaxRestarts, mSecantMinStepNorm, mUseInitialState, _
    mInitialEpsilon0, mInitialKappaX, mInitialKappaY, mUseInitialLoads, _
    mInitialTargetN, mInitialTargetMx, mInitialTargetMy)
'@
        Add-DurationCounter $solver 'Solve' 'core'
        Add-DurationCounter $solver 'EvaluateStrainPlane' 'core'
        Add-DurationCounter $book.VBProject.VBComponents.Item("CFiberMeshBuilder").CodeModule 'BuildMesh' 'geometry'
        Add-DurationCounter $book.VBProject.VBComponents.Item("CSectionStateResult").CodeModule 'InitializeFromSolver' 'packaging'
        Add-DurationCounter $book.VBProject.VBComponents.Item("CCrackWidthResult").CodeModule 'InitializeFromCalculator' 'packaging'
        Insert-Counter $book.VBProject.VBComponents.Item("modGeometryTypes").CodeModule '(?m)^Public Function BuildGeneratedSectionModel\([\s\S]*?\) As CSectionModel\r?\n' 'Audit03RecordModel mesh, rebars' -After
        $capacity = $book.VBProject.VBComponents.Item("CCapacitySolver").CodeModule
        Insert-Counter $capacity '^    If TryRestoreProbeCache\(' 'Audit03RecordProbe'
        Insert-Counter $capacity '^            TryRestoreProbeCache = True' 'Audit03RecordProbeReuse'
        Insert-Counter $capacity '^    AppendDiagnostic solver, mx, my, lambdaValue, attempt, ProbeOnce' 'Audit03RecordOutcome solver, ProbeOnce'
        Insert-Counter $capacity '^        If attempt < mMaxRetries Then mRetryCount = mRetryCount \+ 1' 'If attempt < mMaxRetries Then Audit03RecordRetry'
        Insert-Counter $capacity '(?m)^Public Function LimitSearchEvaluateUltimateResidual\([\s\S]*?\) As Boolean\r?\n' 'Audit03RecordPlane' -After
        $formationName = "CCrackFormationCalculator"
        $formation = $book.VBProject.VBComponents.Item($formationName).CodeModule
        Insert-Counter $formation '(?m)^Public Function LimitSearchEvaluateCrackFormationLoadMultiplier\([\s\S]*?\) As Boolean\r?\n' 'Audit03RecordProbe' -After
        $provider = $book.VBProject.VBComponents.Item("CStateProvider").CodeModule
        Insert-Counter $provider '^            mLastStateWasReused = True' 'Audit03RecordStateReuse'
        Insert-Counter $book.VBProject.VBComponents.Item('CStateSolutionRunner').CodeModule '(?m)^Private Function FormatNumberInvariant\(ByVal value As Double\) As String\r?\n' 'Audit03RecordRetryFormat' -After
        Insert-Counter $book.VBProject.VBComponents.Item('modTestCapacitySolver').CodeModule '^    cap.SolverMaxIterations = 60\r?$' 'If Audit03DisableDiagnostics Then cap.DiagnosticsEnabled = False' -After
        # Batch WriteSummary includes its detailed writers; count only the root.
        Add-DurationCounter $book.VBProject.VBComponents.Item('CBatchResultWriter').CodeModule 'WriteSummary' 'write'
        Add-DurationCounter $book.VBProject.VBComponents.Item('CNDMResultsWriter').CodeModule 'WriteResults' 'write'
        $book.VBProject.VBComponents.Item('modTestGeometry').CodeModule.AddFromString(@'
' ДЛЯ ТЕСТОВ: та же Hollow-сетка в обеих версиях, без уменьшения дискретизации.
Private Sub TestAudit03HollowMeshPerf(ByRef stats As TTestStats)
    Dim geom As CGeometryHollowRectangle
    Set geom = New CGeometryHollowRectangle
    geom.Initialize 500#, 800#, 50#, 200#, 500#, 20#, 40#, -30#
    Dim mesh As CFiberMeshBuilder
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geom, 10#, 10#, 1, 2
    Dim section As CSectionModel
    Set section = BuildGeneratedSectionModel(mesh, Nothing, "HollowPerf")
    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateConcrete section
    Dim available As Boolean
    Dim area As Double
    area = geom.AnalyticalArea(available)
    AssertTrue stats, "perf.hollow.areaAvailable", available
    AssertTrue stats, "perf.hollow.gridArea", Abs(props.Area - area) < area * 0.01
    AssertTrue stats, "perf.hollow.empty", Not geom.ContainsPoint(40#, -30#)
    AppendLine stats, "NUMERIC: hollow; fibers=" & CStr(mesh.FiberCount) & "; area=" & CStr(props.Area) & _
        "; x=" & CStr(props.CentroidX) & "; y=" & CStr(props.CentroidY) & _
        "; Ix=" & CStr(props.Ixc) & "; Iy=" & CStr(props.Iyc) & "; Ixy=" & CStr(props.Ixyc)
End Sub
'@
        )
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestCapacitySolver").CodeModule "TCapacityTestStats" @{
            bending = "TestMxPositiveAndNegative"
            biaxial = "TestMxySignedCombinations"
            asymmetric = "TestAsymmetricCoupledCurvatures"
            diagnostics_off = "TestAsymmetricCoupledCurvatures"
            axial = "TestAxialLoadMultiplierFindsNult"
            methods = "TestLoadMultiplierSearchMethods"
            capacity_strategies = "TestCapacitySolutionStrategyComparisons"
        }
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestGeometry").CodeModule "TTestStats" @{
            geometry_rounded = "TestSymmetricRoundedRectangle"
            geometry_tapered = "TestTaperedRoundedRectangle"
            geometry_hollow = "TestHollowRectangleGeometry"
            geometry_hollow_rebars = "TestHollowRectangleRebarLayout"
            geometry_hollow_mesh = "TestAudit03HollowMeshPerf"
        }
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestCrackWidth").CodeModule "TCrackTestStats" @{
            crack_bending = "TestAutoMcrcPureBendingConverges"
            crack_paths = "TestCrackInitiationLoadPaths"
            crack_cache = "TestCrackFormationCacheHitWithoutLastRunner"
        }
        Add-BenchmarkEntry $book.VBProject.VBComponents.Item("modTestBatchCalculation").CodeModule "TBatchTestStats" @{
            batch_write = "TestBatchSummaryWriter"
        }
        $lines.Add("SOURCE|$label|$($sources[$version])|sha256=$($hashes[$version])")
        $caseNames = @("bending", "biaxial", "asymmetric", "diagnostics_off", "axial", "methods", "capacity_strategies", "crack_bending", "crack_paths", "crack_cache", "geometry_rounded", "geometry_tapered", "geometry_hollow", "geometry_hollow_rebars", "geometry_hollow_mesh", "batch_write")
        if ($Cases.Count) {
            $requested = @($Cases | ForEach-Object { $_ -split ',' })
            if (@($requested | Where-Object { $_ -notin $caseNames }).Count) { throw 'Unknown benchmark case.' }
            $caseNames = $requested
        }
        for ($run = 1; $run -le $Repeats; $run++) {
            foreach ($case in $caseNames) {
                $module = "modTestCapacitySolver"
                if ($case.StartsWith("crack_")) { $module = "modTestCrackWidth" }
                if ($case.StartsWith("geometry_")) { $module = "modTestGeometry" }
                if ($case.StartsWith("batch_")) { $module = "modTestBatchCalculation" }
                $lines.Add("START|$label|run=$run|case=$case")
                Save-PerformanceReport
                $result = [string]$excel.Run("'$($book.Name)'!$module.RunAudit03PerfCase", $case)
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
    Save-PerformanceReport
    $lines | Where-Object { $_ -match '^SOURCE|^PERF|^SCRIPT ERROR' }
}
if ($failed) { exit 1 }
