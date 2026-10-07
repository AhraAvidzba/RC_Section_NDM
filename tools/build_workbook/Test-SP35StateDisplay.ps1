# Адресный native Excel-тест цепочки meta -> status-policy -> detailed writer.
# Тестовый доступ к неопубликованному LC добавляется только в собственную копию,
# не попадает в src, основную книгу или публикационный VBA_All_Code.txt.
param(
    [string]$WorkbookPath = 'docs/regression/SP35/AreaBoundFinal/RC_Section_NDM.xlsm',
    [string]$ReportDirectory = 'docs/regression/SP35/StateDisplayFinal'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory = Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$copy = Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $root $WorkbookPath) -Destination $copy -Force
$excel = $null; $book = $null

# Fixture задает только typed исходы для проверки presentation-контракта.
# Это не подтверждение физического равновесия искусственно созданного state.
$fixture = @'
Option Explicit
Private passed As Long
Private failed As Long
Private report As String
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        passed = passed + 1
        report = report & "OK: " & name & vbCrLf
    Else
        failed = failed + 1
        report = report & "FAIL: " & name & vbCrLf
    End If
End Sub
Public Function RunStateDisplay() As String
    On Error GoTo FailedRun
    Dim section As CSectionModel, provider As CMaterialModelProvider, batch As CBatchSectionCalculator
    Dim profiles As CCalculationProfileCatalog
    Dim concrete As CConcreteMaterialParameters, steel As CSteelMaterialParameters, spec As CMaterialModelSpec
    Dim result As CCombinationResult, aggregate As CCrackResult, formation As CCrackFormationResult
    Dim meta As CResultMeta, formationMeta As CResultMeta, blocked As CResultMeta
    Dim width As CCrackWidthResult, longitudinal As CLongitudinalCrackResult
    Dim request As CStateRequest, state As CSectionStateResult, stateType As Variant
    Dim preState As CSectionStateResult, postState As CSectionStateResult
    Dim statuses As Variant, codes As Variant, i As Long, expected As String
    Dim policy As CResultStatusPolicy, writer As CCrackSummaryWriter, units As CUnitSystem
    Dim sheet As Object, anchor As Object, original As String, stage As String
    passed = 0: failed = 0: report = vbNullString
    original = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo
    Set sheet = ThisWorkbook.Worksheets.Add
    ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo = "='" & sheet.Name & "'!$C$20"
    Set anchor = ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Set section = New CSectionModel
    section.AddConcreteElement 0#, 0#, 800#, 1, , , "Rectangle", 20#, 40#
    Set concrete = New CConcreteMaterialParameters: concrete.Initialize 15.5, 1.1, 22#, 1.8, 32500#, 32500#, rbMc2:=14.6
    Set steel = New CSteelMaterialParameters: steel.Initialize 350#, 350#, 390#, 390#, 200000#, 200000#
    Set provider = New CMaterialModelProvider: provider.InitializeFromParameters concrete, steel
    Set batch = New CBatchSectionCalculator: batch.Initialize section, provider
    Set profiles = New CCalculationProfileCatalog: profiles.LoadFromWorkbook ThisWorkbook
    Set batch.ProfileCatalog = profiles
    Set spec = New CMaterialModelSpec: spec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    Set policy = New CResultStatusPolicy
    statuses = Array(rsSuccess, rsCheckFailed, rsNumericalFailure, rsInvalidInput, rsInternalError, rsNotApplicable, rsSuccessWithWarning)
    codes = Array(rcCheckPassed, rcCheckFailed, rcNumericalFailure, rcInvalidInput, rcInternalError, rcNotApplicable, rcCheckPassed)
    For i = 0 To UBound(statuses)
        stage = "combination " & CStr(i)
        batch.AddCombination "STATUS_" & CStr(i), -1000#, 0#, 0#, "PR2", vbNullString, "Auto"
        Set result = batch.SP35TestUnpublishedResult(i + 1)
        stage = "states " & CStr(i)
        Set meta = New CResultMeta
        meta.SetResult CLng(statuses(i)), CLng(codes(i)), rkDirectState, "Контрольная причина состояния.", , True, False
        For Each stateType In Array(sstPreCrackState, sstPostCrackState, sstCrackedState)
            Set request = New CStateRequest: request.Initialize CLng(stateType), cpCrackedNDS, spec, -1000#, 0#, 0#, True
            Set state = New CSectionStateResult: state.InitializeFailure request, meta
            If stateType = sstPreCrackState Then Set preState = state
            If stateType = sstPostCrackState Then Set postState = state
            result.AddState state
        Next stateType
        stage = "aggregate " & CStr(i)
        Set formationMeta = New CResultMeta
        formationMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCrackFormation, "Контрольный поиск без точки.", , True, False
        Set formation = New CCrackFormationResult: formation.InitializeUncalculated formationMeta
        formation.SP35TestAttachDiagnosticStates preState, postState
        Set blocked = New CResultMeta: blocked.SetBlockedByDependency rkCrackWidth, "текущего НДС"
        Set width = New CCrackWidthResult: width.InitializeFromCalculator Nothing, blocked
        blocked.SetBlockedByDependency rkLongitudinalCrack, "текущего НДС"
        Set longitudinal = New CLongitudinalCrackResult: longitudinal.Initialize blocked, 0#, 0#, 0#
        Set aggregate = New CCrackResult: aggregate.Initialize formation, meta, width, longitudinal
        result.StoreCrackAggregateResult aggregate
        result.SetDiagramExtensionEnabled (i Mod 2 = 0)
    Next i
    stage = "writer"
    Set units = New CUnitSystem: units.InitializeDefaults
    Set writer = New CCrackSummaryWriter: writer.WriteSummary ThisWorkbook, batch, units
    For i = 0 To UBound(statuses)
        Set result = batch.ResultAt(i + 1)
        expected = policy.ExternalStatus(result.CrackCurrentStateMeta)
        Check "stateDisplay.pre." & CStr(i), CStr(anchor.Offset(i, 20).Value2) = expected
        Check "stateDisplay.post." & CStr(i), CStr(anchor.Offset(i, 23).Value2) = expected
        Check "stateDisplay.current." & CStr(i), CStr(anchor.Offset(i, 28).Value2) = expected
        Check "stateDisplay.internalUnchanged." & CStr(i), result.CrackCurrentStateMeta.InternalStatus = CLng(statuses(i))
        Check "stateDisplay.codeUnchanged." & CStr(i), result.CrackCurrentStateMeta.ResultCode = CLng(codes(i))
        Check "stateDisplay.comment." & CStr(i), CStr(anchor.Offset(i, 1).Value2) = result.CrackSummaryMeta.ResultComment
        If i Mod 2 = 0 Then expected = "no" Else expected = "-"
        Check "stateDisplay.ExtUsed." & CStr(i), CStr(anchor.Offset(i, 27).Value2) = expected
        Check "stateDisplay.noFakeStrains." & CStr(i), IsEmpty(anchor.Offset(i, 18).Value2) And IsEmpty(anchor.Offset(i, 25).Value2)
    Next i
    Check "stateDisplay.noSolve", batch.SolverCallCount = 0
    GoTo Restore
FailedRun:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Source & "; " & stage & "; " & Err.Description, False
Restore:
    On Error Resume Next
    ThisWorkbook.Names.Item("rngCrackSummaryAnchor").RefersTo = original
    Application.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    On Error GoTo 0
    RunStateDisplay = report & "TOTAL_SP35_STATE_DISPLAY: passed=" & CStr(passed) & "; failed=" & CStr(failed)
End Function
'@
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($copy)
    $module = $book.VBProject.VBComponents.Item('CBatchSectionCalculator').CodeModule
    $module.AddFromString(@'
Friend Function SP35TestUnpublishedResult(ByVal index As Long) As CCombinationResult
    If mResults(index) Is Nothing Then Set mResults(index) = New CCombinationResult
    Set SP35TestUnpublishedResult = mResults(index)
End Function
'@)
    $module = $book.VBProject.VBComponents.Item('CCrackFormationResult').CodeModule
    $module.AddFromString(@'
Friend Sub SP35TestAttachDiagnosticStates(ByVal pre As CSectionStateResult, ByVal post As CSectionStateResult)
    Set mPreCrackState = pre
    Set mPostCrackState = post
End Sub
'@)
    $component = $book.VBProject.VBComponents.Add(1)
    $component.Name = 'modSP35StateDisplayFixture'
    $component.CodeModule.AddFromString($fixture)
    $result = [string]$excel.Run("'$($book.Name)'!modSP35StateDisplayFixture.RunStateDisplay")
    $result | Set-Content -LiteralPath (Join-Path $directory 'StateDisplay.txt') -Encoding UTF8
    ($result -split "`r?`n" | Where-Object { $_ -match '^FAIL|^TOTAL' }) -join "`n"
    if ($result -notmatch 'failed=0\b') { throw 'State display assertions failed.' }
} finally {
    if ($null -ne $book) { $book.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($null -ne $excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
