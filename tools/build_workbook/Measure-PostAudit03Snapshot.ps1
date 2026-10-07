# Сравнивает инженерные числа baseline/current в собственных Excel-копиях.
# Fixture отличается только согласованным API подготовки снимка; математика,
# нагрузка, профиль и параметры материалов одинаковы. Config пользователя не меняется.
param(
    [string]$CurrentWorkbook = 'docs/regression/PostAudit03/IsolatedFinalDirected/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/PostAudit03/SnapshotComparison'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory = Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$fixture = @'
Option Explicit
Public Function MeasureSnapshot(ByVal scenario As Long) As String
    Dim table As Object, row As Long, key As String
    Set table = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    For row = 2 To table.Rows.Count
        key = CStr(table.Cells(row, 2).Value2)
        Select Case key
            Case "Calculation.Strength.DirectState": table.Cells(row, 3).Value2 = "Yes"
            Case "Calculation.Strength.Capacity", "Calculation.Crack.Width", "Calculation.Stability.Enabled": table.Cells(row, 3).Value2 = "No"
        End Select
    Next row
    Dim section As CSectionModel, x As Long, y As Long
    Set section = New CSectionModel: section.SourceType = "AutoCADImport"
    For x = -10 To 9
        For y = -10 To 9
            section.AddConcreteElement x * 20# + 10#, y * 20# + 10#, 400#, 1, , , "Rectangle", 20#, 20#
        Next y
    Next x
    For x = -1 To 1 Step 2
        For y = -1 To 1 Step 2
            section.AddRebarElement x * 160#, y * 160#, 32#, 0#, "A400"
        Next y
    Next x
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters, provider As CMaterialModelProvider
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider: provider.InitializeFromParameters concrete, steel, diagramExtensionEnabled:=(scenario Mod 2 = 1)
    Dim profiles As CCalculationProfileCatalog, batch As CBatchSectionCalculator, writer As CNDMResultsWriter
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider: Set batch.ProfileCatalog = profiles
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook: batch.ApplySettings settings
    Dim force As Double, mx As Double, my As Double
    force = -500000#: mx = 10000000#: my = -5000000#
    If scenario >= 2 And scenario <= 3 Then force = -7000000#: mx = 0#: my = 0#
    If scenario >= 4 Then force = 100000#: mx = 25000000#: my = 10000000#
    batch.AddCombination "SNAP", force, mx, my, "PR1", "Численная сверка", "Auto"
    batch.Execute
    Dim solves As Long, before As Double, prepared As Double, elapsed As Double
    solves = SectionEquilibriumSolveCount(): before = Timer
    Set writer = New CNDMResultsWriter
    __PREPARE__
    prepared = Timer - before
    __WRITE__
    elapsed = Timer - before
    MeasureSnapshot = "status=" & batch.ResultAt(1).Status & "; extension=" & CStr(batch.ResultAt(1).ExtensionUsed) & _
        "; noSolve=" & CStr(SectionEquilibriumSolveCount() = solves) & "; prepareSeconds=" & CStr(prepared) & "; totalSeconds=" & CStr(elapsed)
End Function
'@
$runs = @(); $baseline = @{}; $differences = @(); $compared = 0
foreach ($mode in 'Baseline', 'Current') {
    $source = Join-Path $root 'docs/regression/PostAudit03/Baseline/RC_Section_NDM.xlsm'
    if ($mode -eq 'Current') { $source = Join-Path $root $CurrentWorkbook }
    $copy = Join-Path $directory ($mode + '.xlsm')
    Copy-Item -LiteralPath $source -Destination $copy -Force
    $excel = $null; $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
        $book = $excel.Workbooks.Open($copy)
        $body = $fixture.Replace('__PREPARE__', '').Replace('__WRITE__', 'writer.WriteResults ThisWorkbook, section, provider, batch')
        if ($mode -eq 'Current') {
            $body = $fixture.Replace('__PREPARE__', 'Dim snapshot As CSectionPropertiesCalculator: Set snapshot = PrepareSectionSnapshot(section, provider)').Replace('__WRITE__', 'writer.WriteResults ThisWorkbook, section, snapshot, batch')
        }
        $component = $book.VBProject.VBComponents.Add(1); $component.Name = 'modPostAudit03MeasureFixture'; $component.CodeModule.AddFromString($body)
        for ($scenario = 0; $scenario -lt 6; $scenario++) {
            $result = [string]$excel.Run("'$($book.Name)'!modPostAudit03MeasureFixture.MeasureSnapshot", $scenario)
            if ($result -notmatch 'noSolve=True') { throw "Snapshot caused solve: $result" }
            $runs += [ordered]@{Mode=$mode; Scenario=$scenario; Result=$result}
            foreach ($name in 'rngNDMSectionGeometry', 'rngNDMElementResults', 'rngNDMSectionProperties', 'rngNDMMaterialDiagrams', 'rngNDMSectionAnnotations') {
                $anchor = $book.Names.Item($name).RefersToRange
                $rows = 1
                while ($anchor.Offset($rows, 0).Value2 -ne $null -and [string]$anchor.Offset($rows, 0).Value2 -ne '') { $rows++ }
                $cols = 0
                while ($anchor.Offset(0, $cols).Value2 -ne $null -and [string]$anchor.Offset(0, $cols).Value2 -ne '') { $cols++ }
                if ($name -eq 'rngNDMSectionGeometry') { $cols = 15 }
                if (-not $cols) { continue }
                $values = $anchor.Resize($rows, $cols).Value2
                # RunID - время записи, оно намеренно не участвует в численной сверке.
                for ($r=2; $r -le $rows; $r++) {
                    for ($c=2; $c -le $cols; $c++) {
                        if ($values[$r,$c] -isnot [double] -and $values[$r,$c] -isnot [int]) { continue }
                        $key = "$scenario|$name|$r|$c"
                        $currentValue = [double]$values[$r,$c]
                        if ($mode -eq 'Baseline') { $baseline[$key] = $currentValue }
                        elseif ($baseline.ContainsKey($key)) {
                            $compared++
                            if ($currentValue -ne $baseline[$key]) {
                                $differences += [ordered]@{Key=$key; Baseline=$baseline[$key]; Current=$currentValue; AbsoluteDifference=[Math]::Abs($currentValue - $baseline[$key])}
                            }
                            $baseline.Remove($key)
                        } else { $differences += [ordered]@{Key=$key; Reason='New numeric cell'} }
                    }
                }
            }
            Write-Output "$mode $scenario $result"
        }
    } finally {
        if ($book) { try {$book.Close($false)} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
        if ($excel) { try {$excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
[ordered]@{Compared=$compared; Missing=@($baseline.Keys); Differences=$differences; Runs=$runs; Acceptance='Exact Double comparison; RunID excluded; no tolerance altered'} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directory 'Comparison.json') -Encoding UTF8
Write-Output "SNAPSHOT_COMPARISON: compared=$compared; missing=$($baseline.Count); differences=$($differences.Count)"
if ($baseline.Count -or $differences.Count) {exit 1}
