Attribute VB_Name = "modTestCalculationSnapshot"
Option Explicit

' ==================== ДЛЯ ТЕСТОВ ====================
' Полный цикл снимка: точные числа, оба хранилища, смена сечения и ошибки файла.
' Выполняется только runner-ом в его частной копии книги; внешние файлы принадлежат тесту.
Private Type TTestSnapshotDouble
    Value As Double ' Число, сравниваемое без округления.
End Type
Private Type TTestSnapshotBits
    Low As Long ' Первые четыре байта Double.
    High As Long ' Последние четыре байта Double.
End Type
Private mTestReport As String, mTestPassed As Long, mTestFailed As Long ' Итог и компактные адреса расхождений.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальные вызовы расчёта и кнопки, полную замену файла и защиту старых Results.
Public Function RunCalculationSnapshotTests() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    TestSnapshotBinaryBuffer
    TestSnapshotUTF16Strings
    TestSnapshotAdaptiveExtent
    TestSnapshotStatusSpans
    TestSnapshotOutputGuards
    TestSnapshotSeparatedProperties
    TestSnapshotOriginalLoadMemory
    Dim initial As CSystemSettingsReader: Set initial = New CSystemSettingsReader: initial.LoadFromWorkbook ThisWorkbook
    Dim extension As String: extension = initial.GetRequiredString("General.DiagramExtension")
    modTestPerformance.TestConfigurePerformanceFixture "MIXED", 6, "Newton", "Circle", "Normal", extension
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim token As String: token = "TestSnapshot_" & Format$(Now, "yyyymmdd_hhnnss")
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    TestSnapshotSetting "Plot.LoadCase", "PERF_01"
    Dim loads As Object: Set loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange
    loads.Cells(2, 1).Value2 = "Сочетание_" & ChrW$(&H3B1) & "_01": TestSnapshotSetting "Plot.LoadCase", "Сочетание_" & ChrW$(&H3B1) & "_01"
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    TestSnapshotCheck "readScope.afterCalculation", TestCalculationSnapshotReadReleased()
    Dim original As CCalculationSnapshot: Set original = store.LoadFile(ThisWorkbook)
    TestSnapshotSummaryLoads original
    TestSnapshotLegacySummary original
    TestSnapshotOneUseSource
    TestSnapshotNativeLayout original
    TestSnapshotValidationReuse original
    TestSnapshotElementIndex
    Dim originalPlot As String: originalPlot = TestSnapshotPlot()
    TestSnapshotSetting "AutoCAD.Export.CombinationID", "Сочетание_" & ChrW$(&H3B1) & "_01"
    Dim originalExport As String: originalExport = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestSnapshotCheck "native.calculate.fileReference", store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "native.calculate.noSheetArchive", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotCheck "native.calculate.states", UBound(original.Table("rngNDMElementResults"), 1) > 1
    TestSnapshotCheck "native.calculate.elapsed", original.ElapsedSeconds > 0#
    TestSnapshotCheck "native.common.onlyALL", TestSnapshotCommonOnly(original.Table("rngNDMSectionProperties"))
    TestSnapshotCheck "native.lc.separate", UBound(original.Table("LoadCase.Properties"), 1) > 1
    TestSnapshotUpper original
    TestSnapshotCorruptUpper
    message = vbNullString
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper original
    TestSnapshotCheck "file.restored.plot.exact", TestSnapshotPlot() = originalPlot
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    TestSnapshotCheck "name.extensionAutomatic", path = ThisWorkbook.Path & Application.PathSeparator & token & "_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotCheck "name.explicitSuffixCompatible", store.SnapshotPath(ThisWorkbook) = path
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndms"
    TestSnapshotCheck "name.previousSuffixCompatible", store.SnapshotPath(ThisWorkbook) = path
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1)
    TestSnapshotLegacyFile original, store
    store.SaveFile ThisWorkbook, original
    Dim fileSnapshot As CCalculationSnapshot: Set fileSnapshot = store.LoadFile(ThisWorkbook)
    TestSnapshotEqual "file.allTables", original, fileSnapshot
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "file.linked", store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "file.noSheetArchive", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotUpper original
    TestSnapshotCheck "file.plot.exact", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "file.export.exact", Audit02ReadExportSnapshotForTests(ThisWorkbook) = originalExport
    TestSnapshotCheck "readScope.afterExport", TestCalculationSnapshotReadReleased()
    TestSnapshotEqual "file.linkedExact", original, store.LoadLinkedFile(ThisWorkbook)
    Dim geometry As Variant: geometry = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry")
    TestSnapshotCheck "file.mesh.alwaysShown", TestSnapshotSameMatrix(geometry, original.Table("rngNDMSectionGeometry"))
    TestSnapshotCheck "file.diagrams.alwaysShown", TestSnapshotSameMatrix(ReadPhysicalResultTable(ThisWorkbook, "rngNDMMaterialDiagrams"), original.Table("rngNDMMaterialDiagrams"))
    Dim profiles As Object: Set profiles = ThisWorkbook.Names.Item("rngCalculationProfiles").RefersToRange
    Dim header As Object: Set header = profiles.Find("PR1", LookAt:=1)
    If header Is Nothing Then Err.Raise 5, , "В тестовой таблице отсутствует PR1."
    header.Value2 = "PR91"
    TestSnapshotCheck "file.savedProfile.fallback", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "file.savedProfile.export", Audit02ReadExportSnapshotForTests(ThisWorkbook) = originalExport
    header.Value2 = "PR1"
    ' Новый выбранный файл не должен менять показанное НДС до нажатия кнопки.
    TestSnapshotSetting "Results.SnapshotName", token & "_other.ndm"
    TestSnapshotCheck "selection.keepsDisplayed", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "missing.rejected", TestSnapshotRefreshRejected()
    TestSnapshotCheck "readScope.afterFailure", TestCalculationSnapshotReadReleased()
    TestSnapshotCheck "missing.keepsGeometry", TestSnapshotSameMatrix(geometry, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim value As Byte: file.OpenUpdate path: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotCheck "corrupt.rejected", TestSnapshotRefreshRejected()
    TestSnapshotCheck "corrupt.keepsGeometry", TestSnapshotSameMatrix(geometry, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionGeometry"))
    store.SaveFile ThisWorkbook, original
    UpdateSectionResultsForWorkbook ThisWorkbook
    ' Другая геометрия и единицы, тот же файл: старые строки должны исчезнуть полностью.
    TestSnapshotSetting "Geometry.Source", "Generated": TestSnapshotSetting "Geometry.Type", "Rectangles"
    TestSnapshotSetting "Units.Length.Output", "cm": TestSnapshotSetting "Units.Area.Output", "cm2"
    TestSnapshotSetting "Units.Stress.Output", "kgf/cm2"
    message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim second As CCalculationSnapshot: Set second = store.LoadFile(ThisWorkbook)
    TestSnapshotCheck "overwrite.newGeometry", Not TestSnapshotSameMatrix(original.Table("rngNDMSectionGeometry"), second.Table("rngNDMSectionGeometry"))
    TestSnapshotCheck "overwrite.units", second.OutputLengthUnit = "cm"
    TestSnapshotCheck "overwrite.oneFile", Not CreateObject("Scripting.FileSystemObject").FileExists(path & ".part") And CreateObject("Scripting.FileSystemObject").FileExists(path)
    TestSnapshotUpper second
    Dim secondPlot As String: secondPlot = TestSnapshotPlot()
    TestSnapshotSetting "Results.SnapshotName", token & "_other.ndm": store.SaveFile ThisWorkbook, original
    Dim inputs As Variant: inputs = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Formula
    TestSnapshotCorruptUpper
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper original
    TestSnapshotCheck "switch.otherSection.plot", TestSnapshotPlot() = originalPlot
    TestSnapshotCheck "switch.inputPreserved", TestSnapshotSameMatrix(inputs, ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Formula)
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm": UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "switch.back.plot", TestSnapshotPlot() = secondPlot
    ' Никакого обращения к решателю при обновлении: даже ошибочный solver Config допустим.
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotUpper second
    TestSnapshotSetting "Solver.MaxIterations", "100"
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    Dim detached As CCalculationSnapshot: Set detached = CurrentCalculationSnapshot(ThisWorkbook)
    TestSnapshotCheck "excel.ignoresPreviousFileLink", detached Is Nothing
    SaveAndDisplayCalculationSnapshot ThisWorkbook, second
    TestSnapshotExcelCurrent second
    Dim invalidName As Variant
    For Each invalidName In Array("../escape", "CON.ndm", "a:b", "", "NUL", "LPT1.ndm")
        TestSnapshotSetting "Results.SnapshotName", CStr(invalidName)
        TestSnapshotCheck "name.rejected." & CStr(invalidName), TestSnapshotPathRejected(store)
    Next invalidName
    TestSnapshotSetting "Results.SnapshotName", token & "_" & ChrW$(&H3B1) & ".ndm"
    Dim preview As String: preview = FormSectionForWorkbook(ThisWorkbook)
    TestSnapshotCheck "preview.clearsSnapshot", Not store.HasFileReference(ThisWorkbook)
    TestSnapshotCheck "readScope.afterPreview", TestCalculationSnapshotReadReleased()
    TestSnapshotCheck "layout.lowerOrder", TestSnapshotLowerOrder()
    TestSnapshotRepeatedImportedAnnotations extension
    TestSnapshotColumnCodec second, store
    TestSnapshotRepeatedElementIDs second, store
    TestSnapshotMinimalFormatting
    GoTo Finished
Failed:
    TestSnapshotCheck "runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    On Error Resume Next
    If Not store Is Nothing Then store.ClearReference ThisWorkbook
    If Len(path) > 0 Then If CreateObject("Scripting.FileSystemObject").FileExists(path) Then CreateObject("Scripting.FileSystemObject").DeleteFile path
    Dim otherPath As String: otherPath = ThisWorkbook.Path & Application.PathSeparator & token & "_other.ndm"
    If Len(token) > 0 Then If CreateObject("Scripting.FileSystemObject").FileExists(otherPath) Then CreateObject("Scripting.FileSystemObject").DeleteFile otherPath
    On Error GoTo 0
    RunCalculationSnapshotTests = mTestReport & "TOTAL CALCULATION SNAPSHOT: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет границы старого и нового листа, большие оценки и точную классификацию ошибки объединения.
Private Sub TestSnapshotOutputGuards()
    Dim limit As Variant
    For Each limit In Array(65536, ThisWorkbook.Worksheets("Results").Rows.Count)
        TestSnapshotCheck "rows.exact." & CStr(limit), ResultRowsFitSheet(156, CDbl(limit) - 155#, CLng(limit))
        TestSnapshotCheck "rows.offsetOverflow." & CStr(limit), Not ResultRowsFitSheet(156, CDbl(limit) - 154#, CLng(limit))
        TestSnapshotCheck "rows.totalOverflow." & CStr(limit), Not ResultRowsFitSheet(1, CDbl(limit) + 1#, CLng(limit))
        TestSnapshotCheck "rows.largeDouble." & CStr(limit), Not ResultRowsFitSheet(156, 5000000000#, CLng(limit))
    Next limit
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    TestSnapshotCheck "rows.estimateNoLongOverflow", writer.EstimatedElementResultRows(1000000, 5000) = 5000000001#
    Dim russian As String, english As String
    russian = ResultOutputErrorDescription("Мы не можем это сделать в объединенной ячейке.")
    english = ResultOutputErrorDescription("We can't do that to a merged cell.")
    TestSnapshotCheck "merged.russian", InStr(russian, "Не удалось вывести результаты на лист Results") > 0 And InStr(russian, "Снимите объединение") > 0
    TestSnapshotCheck "merged.english", english = russian
    TestSnapshotCheck "merged.other1004Unchanged", ResultOutputErrorDescription("Application-defined or object-defined error") = "Application-defined or object-defined error"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ до solver у последней строки, частичное объединение и оформление перенесённой шапки.
Private Sub TestSnapshotNativeLayout(ByVal original As CCalculationSnapshot)
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    Dim name As Object: Set name = ThisWorkbook.Names.Item("rngNDMElementResults")
    Dim previous As String: previous = name.RefersTo
    Dim anchor As Object: Set anchor = name.RefersToRange
    Dim limit As Long: limit = anchor.Worksheet.Rows.Count
    name.RefersTo = "=" & anchor.Worksheet.Cells(limit - 2, anchor.Column).Address(External:=True)
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    Dim reason As String: reason = TestSnapshotCalculationError()
    TestSnapshotCheck "rows.native.preflight", InStr(reason, "Results.SnapshotStorage = File") > 0 And InStr(reason, CStr(limit)) > 0 And InStr(reason, "1)") > 0
    TestSnapshotCheck "rows.native.noSolver", SectionEquilibriumSolveCount() = solves
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "rows.native.restoreAdvice", InStr(reason, "1) На Config установите Results.SnapshotStorage = File") > 0
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotUpper original
    Set name = ThisWorkbook.Names.Item("rngNDMSectionProperties")
    previous = name.RefersTo: Set anchor = name.RefersToRange
    Dim common As Variant: common = original.Table("rngNDMSectionProperties")
    Dim cases As Variant: cases = original.Table("LoadCase.Properties")
    Dim staged() As Variant, row As Long, column As Long, nextRow As Long
    ReDim staged(1 To UBound(common, 1) + 2, 1 To 6)
    For row = 1 To UBound(common, 1)
        For column = 1 To 6: staged(row, column) = common(row, column): Next column
    Next row
    nextRow = UBound(common, 1) + 1
    For row = 2 To UBound(cases, 1)
        If CStr(cases(row, 3)) = "LoadReferenceX" Or CStr(cases(row, 3)) = "LoadReferenceY" Then
            For column = 1 To 6: staged(nextRow, column) = cases(row, column): Next column
            nextRow = nextRow + 1
        End If
    Next row
    name.RefersTo = "=" & anchor.Worksheet.Cells(limit - UBound(staged, 1) + 1, anchor.Column).Address(External:=True)
    name.RefersToRange.Resize(UBound(staged, 1), 6).Value2 = staged ' Сохраняем метаданные сформированного сечения: проверяется переполнение вывода, а не пустой источник геометрии.
    solves = SectionEquilibriumSolveCount()
    reason = TestSnapshotCalculationError()
    TestSnapshotCheck "rows.readableLC.preflight", InStr(reason, "rngNDMSectionProperties") > 0 And InStr(reason, CStr(limit)) > 0
    TestSnapshotCheck "rows.readableLC.noSolver", SectionEquilibriumSolveCount() = solves
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "rows.readableLC.prepared", InStr(reason, "rngNDMSectionProperties") > 0 And InStr(reason, CStr(limit)) > 0
    name.RefersToRange.Resize(UBound(staged, 1), 6).Clear
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    Set name = ThisWorkbook.Names.Item("rngBatchSummary")
    previous = name.RefersTo
    Set anchor = ThisWorkbook.Worksheets("Results").Cells(name.RefersToRange.Row, 500)
    anchor.Offset(0, -1).Resize(1, 2).Merge
    name.RefersTo = "=" & anchor.Address(External:=True)
    reason = TestSnapshotPreparedRefreshError(original)
    TestSnapshotCheck "merged.native.refresh", InStr(reason, "область вывода попадает в объединённые ячейки") > 0
    reason = TestSnapshotCalculationError()
    TestSnapshotCheck "merged.native.calculate", InStr(reason, "область вывода попадает в объединённые ячейки") > 0
    anchor.Offset(0, -1).Resize(1, 2).UnMerge
    name.RefersTo = previous
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    Set name = ThisWorkbook.Names.Item("rngBatchSummary")
    previous = name.RefersTo
    Set anchor = ThisWorkbook.Worksheets("Results").Cells(name.RefersToRange.Row, 500)
    name.RefersTo = "=" & anchor.Address(External:=True)
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotUpper original
    TestSnapshotCheck "minimal.movedHeaderCreated", anchor.Cells(6, 1).Font.Bold
    anchor.Cells(6, 1).Font.Color = RGB(12, 34, 56)
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    TestSnapshotCheck "minimal.recognizedHeaderPreserved", CLng(anchor.Cells(6, 1).Font.Color) = RGB(12, 34, 56)
    name.RefersTo = previous
    TestSnapshotSetting "General.ReformatResultsEnabled", "Yes"
    TestSnapshotSetting "Results.SnapshotStorage", "File"
    UpdateSectionResultsForWorkbook ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает реальную причину отказа расчёта в частной книге без изменения ожиданий.
Private Function TestSnapshotCalculationError() As String
    On Error GoTo Rejected
    Dim message As String: message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotCalculationError = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт готовый снимок, чтобы проверить размещение без повторного чтения уже перенесённого якоря.
Private Function TestSnapshotPreparedRefreshError(ByVal original As CCalculationSnapshot) As String
    On Error GoTo Rejected
    UpdateSectionResultsForWorkbook ThisWorkbook, original
    Exit Function
Rejected:
    TestSnapshotPreparedRefreshError = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Старый суффикс читается только при отсутствии нового; повреждение нового не маскируется старым.
Private Sub TestSnapshotLegacyFile(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Dim previous As String: previous = Left$(path, Len(path) - 4) & ".ndms"
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim checksum As CFileChecksum: Set checksum = New CFileChecksum
    store.SaveFile ThisWorkbook, original
    fileSystem.MoveFile path, previous ' FSO сохраняет Unicode имени; оператор VBA Name зависит от системной кодовой страницы.
    Dim hash As String: hash = checksum.FileSHA256(previous)
    TestSnapshotEqual "legacy.load", original, store.LoadFile(ThisWorkbook)
    store.LinkFile ThisWorkbook
    TestSnapshotEqual "legacy.link", original, store.LoadLinkedFile(ThisWorkbook)
    store.SaveFile ThisWorkbook, original
    TestSnapshotCheck "legacy.saveNewSuffix", fileSystem.FileExists(path) And checksum.FileSHA256(previous) = hash
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    Dim value As Byte: file.OpenUpdate path: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotCheck "legacy.corruptPreferredRejected", TestSnapshotFileRejected(store)
    TestSnapshotCheck "legacy.oldFileUntouched", checksum.FileSHA256(previous) = hash
    store.SaveFile ThisWorkbook, original
    file.OpenUpdate previous: file.Position = 95: value = file.ReadByte
    file.Position = 95: file.WriteByte value Xor 1: file.CloseFile
    TestSnapshotEqual "legacy.preferredWins", original, store.LoadFile(ThisWorkbook)
    fileSystem.DeleteFile previous
    store.ClearReference ThisWorkbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неудачное чтение выбранного файла обязано вернуть ошибку, не переключаясь на старый снимок.
Private Function TestSnapshotFileRejected(ByVal store As CCalculationSnapshotStore) As Boolean
    On Error GoTo Rejected
    Dim snapshot As CCalculationSnapshot: Set snapshot = store.LoadFile(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotFileRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо заданные поля SP35 и точки нагрузки принадлежат LC-блоку, а не геометрии.
Private Sub TestSnapshotSeparatedProperties()
    Dim snapshot As CCalculationSnapshot: Set snapshot = New CCalculationSnapshot
    Dim data(1 To 7, 1 To 6) As Variant, keys As Variant, row As Long
    data(1, 1) = "RunID": data(1, 2) = "LoadCase": data(1, 3) = "Parameter"
    data(1, 4) = "Value": data(1, 5) = "Unit": data(1, 6) = "Comment"
    keys = Array("Gross.Area", "SP35.Slenderness1", "SP35.Slenderness2", "LoadReferenceX", "LoadReferenceY", "StateLoadN")
    For row = 2 To 7
        data(row, 1) = 1#: data(row, 2) = "ALL": data(row, 3) = keys(row - 2): data(row, 4) = CDbl(row)
    Next row
    data(7, 2) = "LC_01"
    snapshot.AddSeparatedTable "rngNDMSectionProperties", data
    Dim common As Variant, cases As Variant
    common = snapshot.Table("rngNDMSectionProperties"): cases = snapshot.Table("LoadCase.Properties")
    TestSnapshotCheck "separated.commonOnlyArea", UBound(common, 1) = 2 And common(2, 3) = "Gross.Area"
    TestSnapshotCheck "separated.slendernessInCases", UBound(cases, 1) = 6 And cases(2, 3) = "SP35.Slenderness1" And cases(3, 3) = "SP35.Slenderness2"
    TestSnapshotCheck "separated.logicalTableExact", TestSnapshotSameMatrix(data, snapshot.ResultTable("rngNDMSectionProperties"))
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Оба хранилища обязаны соблюдать выключенное переоформление общих таблиц Results.
Private Sub TestSnapshotMinimalFormatting()
    Dim mode As Variant, target As Object, message As String
    TestSnapshotSetting "General.ReformatResultsEnabled", "No"
    For Each mode In Array("Excel", "File")
        TestSnapshotSetting "Results.SnapshotStorage", CStr(mode)
        Set target = ThisWorkbook.Names.Item("rngNDMSectionGeometry").RefersToRange.Cells(2, 4)
        target.Font.Color = RGB(12, 34, 56): target.NumberFormat = "0.0000000000"
        message = RunSectionCalculationForWorkbook(ThisWorkbook)
        TestSnapshotCheck "minimal." & CStr(mode) & ".font", CLng(target.Font.Color) = RGB(12, 34, 56)
        TestSnapshotCheck "minimal." & CStr(mode) & ".numberFormat", CStr(target.NumberFormat) = "0.0000000000"
    Next mode
    TestSnapshotSetting "General.ReformatResultsEnabled", "Yes"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимый байтовый эталон проверяет границу буфера, прямой большой блок и seek/read.
Private Sub TestSnapshotBinaryBuffer()
    Dim path As String: path = ThisWorkbook.Path & Application.PathSeparator & "TestSnapshotBuffer.bin"
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    Dim expected(0 To 131088) As Byte, actual() As Byte, row As Long, same As Boolean
    For row = 0 To UBound(expected): expected(row) = row Mod 251: Next row
    file.OpenWrite path: file.WriteASCII "NDMBUF01": file.WriteBytes expected
    For row = 1 To 20000: file.WriteLong row: Next row
    file.WriteDouble 1E-250
    TestSnapshotCheck "buffer.logicalSize", file.Size = 8 + 131089 + 80000 + 8 And file.Position = file.Size
    file.Position = 8 + 65535: file.WriteASCII "BUFFER": file.CloseFile
    For row = 0 To 5: expected(65535 + row) = Asc(Mid$("BUFFER", row + 1, 1)): Next row
    file.OpenRead path
    TestSnapshotCheck "buffer.magic", file.ReadASCII(8) = "NDMBUF01"
    actual = file.ReadBytes(131089): same = True
    For row = 0 To UBound(expected)
        If actual(row) <> expected(row) Then same = False: Exit For
    Next row
    TestSnapshotCheck "buffer.largeBlockAndSeek", same
    same = True
    For row = 1 To 20000
        If file.ReadLong <> row Then same = False
    Next row
    TestSnapshotCheck "buffer.smallWrites", same
    TestSnapshotCheck "buffer.doubleExact", TestSnapshotDouble(file.ReadDouble) = TestSnapshotDouble(1E-250)
    TestSnapshotCheck "buffer.eofRejected", TestSnapshotBinaryEOFRejected(file)
    file.CloseFile: file.OpenUpdate path: file.Position = 0: file.WriteByte 42: file.Position = 0
    TestSnapshotCheck "buffer.flushBeforeRead", file.ReadByte = 42
    file.WriteByte 99 ' После короткого чтения Windows-указатель находится дальше Position из-за предварительного чтения.
    TestSnapshotCheck "buffer.writeAtLogicalReadPosition", file.ReadByte = Asc("M") And file.Position = 3
    file.Position = 0
    TestSnapshotCheck "buffer.readAfterMixedWrite", file.ReadByte = 42 And file.ReadByte = 99
    Dim chunk As Variant
    For Each chunk In Array(127&, 32767&, 65535, 65536, 65537, 131089)
        file.Position = 8: actual = file.ReadBytes(CLng(chunk)): same = True
        For row = 0 To UBound(actual)
            If actual(row) <> expected(row) Then same = False: Exit For
        Next row
        TestSnapshotCheck "buffer.readChunk." & CStr(chunk), same And file.Position = 8 + CLng(chunk)
    Next chunk
    file.Position = 8 + 65534: actual = file.ReadBytes(19): same = True
    For row = 0 To UBound(actual)
        If actual(row) <> expected(65534 + row) Then same = False: Exit For
    Next row
    TestSnapshotCheck "buffer.backwardSeekAndBoundary", same
    file.CloseFile
    CreateObject("Scripting.FileSystemObject").DeleteFile path
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    file.CloseFile
    If CreateObject("Scripting.FileSystemObject").FileExists(path) Then CreateObject("Scripting.FileSystemObject").DeleteFile path
    On Error GoTo 0
    Err.Raise number, "TestSnapshotBinaryBuffer", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Чтение за физическим концом файла обязано завершиться ошибкой.
Private Function TestSnapshotBinaryEOFRejected(ByVal file As CBinarySnapshotFile) As Boolean
    On Error GoTo Rejected
    Dim value As Byte: value = file.ReadByte
    Exit Function
Rejected:
    TestSnapshotBinaryEOFRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ускоренное чтение границ сохраняет первую пустую ячейку, пробелы и точный maxCount.
Private Sub TestSnapshotAdaptiveExtent()
    Dim sheet As Object, app As Object, alerts As Boolean
    Set app = ThisWorkbook.Application: alerts = app.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    Dim values() As Variant, row As Long
    ReDim values(1 To 100001, 1 To 1)
    For row = 1 To UBound(values, 1): values(row, 1) = row: Next row
    sheet.Cells(1, 1).Resize(UBound(values, 1), 1).Value2 = values
    sheet.Cells(100003, 1).Value2 = CVErr(2015) ' Посторонняя ошибка за первой пустой ячейкой не входит в таблицу.
    TestSnapshotCheck "extent.large", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 100001
    TestSnapshotCheck "extent.maxCount", SnapshotAnchoredExtent(sheet.Cells(1, 1), False, 65537) = 65537
    sheet.Cells(4097, 1).Value2 = " "
    TestSnapshotCheck "extent.spaceAtGrowingBoundary", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 4096
    sheet.Cells(1024, 1).ClearContents
    TestSnapshotCheck "extent.emptyAtInitialBoundary", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 1023
    sheet.Cells(1, 1).ClearContents
    TestSnapshotCheck "extent.emptyAnchor", SnapshotAnchoredExtent(sheet.Cells(1, 1), False) = 0
    ReDim values(1 To 1, 1 To 4098)
    For row = 1 To UBound(values, 2): values(1, row) = "Header" & CStr(row): Next row
    sheet.Cells(2, 2).Resize(1, UBound(values, 2)).Value2 = values
    sheet.Cells(2, 2 + 3072).ClearContents
    TestSnapshotCheck "extent.columns", SnapshotAnchoredExtent(sheet.Cells(2, 2), True) = 3072
    app.DisplayAlerts = False: sheet.Delete: app.DisplayAlerts = alerts
    Exit Sub
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    On Error Resume Next
    app.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    app.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise number, "TestSnapshotAdaptiveExtent", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Объединение цветных интервалов сохраняет пропуски и работает без чтения заполненных ячеек.
Private Sub TestSnapshotStatusSpans()
    Dim sheet As Object, app As Object, alerts As Boolean
    Set app = ThisWorkbook.Application: alerts = app.DisplayAlerts
    On Error GoTo Failed
    Set sheet = ThisWorkbook.Worksheets.Add
    Dim target As Object: Set target = sheet.Cells(4, 5).Resize(13, 3)
    Dim values(1 To 13, 1 To 3) As Variant, body(1 To 11, 1 To 3) As Variant
    Dim rows(1 To 8) As Long, ids As Variant, row As Long, column As Long, testCase As Long, i As Long, selected As Boolean, same As Boolean
    ids = Array(3&, 4&, 10&, 6&, 7&, 8&, 13&, 12&)
    For i = 1 To 8: rows(i) = CLng(ids(i - 1)): Next i
    For row = 1 To 13
        values(row, 1) = "OK": values(row, 2) = "FAIL": values(row, 3) = "N/A"
    Next row
    values(4, 3) = "Cracked": values(10, 3) = "NotCracked": values(6, 3) = ""
    For row = 1 To 11
        For column = 1 To 3: body(row, column) = values(row + 2, column): Next column
    Next row
    For testCase = 0 To 2
        target.Interior.Color = RGB(1, 2, 3)
        If testCase = 0 Then
            target.Value2 = values: ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3
        ElseIf testCase = 1 Then
            target.Value2 = CVErr(2015): ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3, values
        Else
            target.Value2 = CVErr(2015): ApplyTableStatusFills target, rows, Array(1, 2, 3), True, 3, body, 2
        End If
        same = True
        For row = 1 To 13
            selected = False
            For i = 1 To 8: If rows(i) = row Then selected = True
            Next i
            For column = 1 To 3
                Dim color As Long: color = RGB(1, 2, 3)
                If selected Then
                    If Not (column = 3 And (row = 4 Or row = 6 Or row = 10)) Then color = StatusFillColor(CStr(values(row, column)))
                End If
                If CLng(target.Cells(row, column).Interior.Color) <> color Then same = False
            Next column
        Next row
        TestSnapshotCheck "statusSpans.gapsAndPreparedValues." & CStr(testCase), same
    Next testCase
    app.DisplayAlerts = False: sheet.Delete: app.DisplayAlerts = alerts
    Exit Sub
Failed:
    Dim number As Long, description As String: number = Err.Number: description = Err.Description
    On Error Resume Next
    app.DisplayAlerts = False
    If Not sheet Is Nothing Then sheet.Delete
    app.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise number, "TestSnapshotStatusSpans", description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет прямой строковый IO независимо от кодека: старые байтовые операции служат эталоном.
Private Sub TestSnapshotUTF16Strings()
    Dim path As String: path = ThisWorkbook.Path & Application.PathSeparator & "TestUTF16_" & Format$(Now, "yyyymmdd_hhnnss") & ".bin"
    If Len(Dir$(path)) > 0 Then Err.Raise 5, , "Тестовый файл UTF16 уже существует."
    Dim file As CBinarySnapshotFile: Set file = New CBinarySnapshotFile
    On Error GoTo Failed
    Dim samples As Variant, sample As Variant, bytes() As Byte, actual As String, count As Long, index As Long
    samples = Array(vbNullString, "ASCII", ChrW$(&H416) & ChrW$(&H3B1), "left" & Chr$(0) & "right", _
        ChrW$(&HD83D) & ChrW$(&HDE00), ChrW$(&HD800), String$(32767, "x"), String$(65536, "z"))
    file.OpenWrite path
    For Each sample In samples: file.WriteSizedUTF16 CStr(sample): Next sample
    file.CloseFile: file.OpenRead path
    For Each sample In samples
        count = file.ReadLong: actual = vbNullString
        If count > 0 Then bytes = file.ReadBytes(count): actual = bytes
        TestSnapshotCheck "utf16.write.matchesBytes." & CStr(index), count = LenB(CStr(sample)) And StrComp(actual, CStr(sample), vbBinaryCompare) = 0
        index = index + 1
    Next sample
    file.CloseFile: file.OpenWrite path
    For Each sample In samples
        file.WriteLong LenB(CStr(sample))
        If LenB(CStr(sample)) > 0 Then bytes = CStr(sample): file.WriteBytes bytes
    Next sample
    file.CloseFile: file.OpenRead path: index = 0
    For Each sample In samples
        actual = file.ReadSizedUTF16(131072)
        TestSnapshotCheck "utf16.read.matchesBytes." & CStr(index), LenB(actual) = LenB(CStr(sample)) And StrComp(actual, CStr(sample), vbBinaryCompare) = 0
        index = index + 1
    Next sample
    file.CloseFile
    Dim testCase As Long
    For testCase = 0 To 3
        file.OpenWrite path
        Select Case testCase
            Case 0: file.WriteLong -1
            Case 1: file.WriteLong 3: file.WriteASCII "abc"
            Case 2: file.WriteSizedUTF16 String$(65537, "z")
            Case 3: file.WriteLong 20: file.WriteByte 7
        End Select
        file.CloseFile: file.OpenRead path
        TestSnapshotCheck "utf16.invalidLength." & CStr(testCase), TestSnapshotUTF16Rejected(file)
        file.CloseFile
    Next testCase
    Kill path
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    file.CloseFile: Kill path
    On Error GoTo 0
    Err.Raise number, "TestSnapshotUTF16Strings", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Перехватывает только ожидаемую ошибку чтения некорректной длины строки.
Private Function TestSnapshotUTF16Rejected(ByVal file As CBinarySnapshotFile) As Boolean
    On Error GoTo Rejected
    Dim text As String: text = file.ReadSizedUTF16(131072)
    Exit Function
Rejected:
    TestSnapshotUTF16Rejected = (Err.Number = 5)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторное использование writer-а читает новые входные строки, а не массив предыдущей операции.
Private Sub TestSnapshotOneUseSource()
    Dim target As Object, saved As Variant
    Set target = ThisWorkbook.Names("rngNDMSectionGeometry").RefersToRange.Cells(2, 15)
    saved = target.Formula
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = ReadFormedSectionFromResults(ThisWorkbook)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    Dim provider As CMaterialModelProvider: Set provider = New CMaterialModelProvider: provider.Initialize settings, units
    Dim props As CSectionPropertiesCalculator: Set props = PrepareSectionSnapshot(section, provider)
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim batch As CBatchSectionCalculator: Set batch = New CBatchSectionCalculator
    Dim first As CCalculationSnapshot: Set first = New CCalculationSnapshot
    target.Value2 = "TestSourceBefore"
    writer.PreparePreservedSnapshotSource ThisWorkbook
    Set writer.CaptureSnapshot = first
    writer.WriteResults ThisWorkbook, section, props, batch, units, True
    Dim values As Variant: values = first.Table("rngNDMSectionGeometry")
    TestSnapshotCheck "capture.oneUse.initial", CStr(values(2, 15)) = "TestSourceBefore"
    target.Value2 = "TestSourceAfter"
    Dim second As CCalculationSnapshot: Set second = New CCalculationSnapshot
    Set writer.CaptureSnapshot = second
    writer.WriteResults ThisWorkbook, section, props, batch, units, True
    values = second.Table("rngNDMSectionGeometry")
    TestSnapshotCheck "capture.oneUse.newSource", CStr(values(2, 15)) = "TestSourceAfter"
    target.Formula = saved
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    target.Formula = saved
    On Error GoTo 0
    Err.Raise number, "TestSnapshotOneUseSource", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторная проверка сохраняет защиту при изменении даты и добавлении новой части.
Private Sub TestSnapshotValidationReuse(ByVal snapshot As CCalculationSnapshot)
    snapshot.Validate: snapshot.Validate
    Dim run As Double: run = snapshot.RunID
    snapshot.RunID = run + 1#
    TestSnapshotCheck "validate.changedRunRejected", TestSnapshotValidationRejected(snapshot)
    snapshot.RunID = run: snapshot.Validate
    Dim copy As Variant: copy = snapshot.Table("rngNDMElementResults")
    copy(2, 1) = run + 2#
    snapshot.Validate
    Dim actual As Variant: actual = snapshot.Table("rngNDMElementResults")
    TestSnapshotCheck "validate.returnedArrayIndependent", CDbl(actual(2, 1)) = run
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает отказ действительного валидатора, не подменяя ожидаемые статусы.
Private Function TestSnapshotValidationRejected(ByVal snapshot As CCalculationSnapshot) As Boolean
    On Error GoTo Rejected
    snapshot.Validate
    Exit Function
Rejected:
    TestSnapshotValidationRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Индекс обязан сохранять несмежные группы, регистр, порядок и произвольные позиции колонок.
Private Sub TestSnapshotElementIndex()
    Dim snapshot As CCalculationSnapshot: Set snapshot = New CCalculationSnapshot
    Dim data(1 To 7, 1 To 9) As Variant, row As Long
    data(1, 1) = "RunID": data(1, 5) = "LoadCase": data(1, 3) = "StateType"
    For row = 2 To 7
        data(row, 1) = 1#: data(row, 2) = "Element" & CStr(row)
        data(row, 5) = IIf(row = 4 Or row = 7, "B", "A")
        data(row, 3) = IIf(row = 3 Or row = 7, "CrackedState", "StrengthState")
    Next row
    snapshot.AddTable "rngNDMElementResults", data
    Dim actual As Variant: actual = snapshot.SelectElementResults("a", "strengthstate", False)
    TestSnapshotCheck "index.nonAdjacentOrder", UBound(actual, 1) = 4 And actual(2, 2) = "Element2" And actual(3, 2) = "Element5" And actual(4, 2) = "Element6"
    actual = snapshot.SelectElementResults("", "", True)
    TestSnapshotCheck "index.firstRows", UBound(actual, 1) = 5 And actual(2, 2) = "Element2" And actual(3, 2) = "Element3" And actual(4, 2) = "Element4" And actual(5, 2) = "Element7"
    actual = snapshot.SelectElementResults("missing", "", False)
    TestSnapshotCheck "index.missingHeaderOnly", UBound(actual, 1) = 1
    actual = snapshot.SelectElementResults("", "", False)
    TestSnapshotCheck "index.fullExact", TestSnapshotSameMatrix(data, actual)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет плотные колонки и серии независимо заданных чисел, Unicode и типов.
Private Sub TestSnapshotColumnCodec(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim sample As CCalculationSnapshot: Set sample = New CCalculationSnapshot
    sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
    Dim key As Variant
    For Each key In original.Keys(): sample.AddTable CStr(key), original.Table(CStr(key)): Next key
    Dim data(1 To 1024, 1 To 8) As Variant, row As Long, negative As TTestSnapshotDouble, bits As TTestSnapshotBits
    bits.High = &H80000000: LSet negative = bits
    For row = 1 To 1024
        data(row, 1) = "Сечение_" & ChrW$(&H3B1)
        data(row, 2) = CDbl(row) / 7#
        data(row, 3) = (row < 513)
        data(row, 4) = 0#: If row > 512 Then data(row, 4) = negative.Value
        data(row, 5) = "": If row > 512 Then data(row, 5) = Empty
        data(row, 6) = CDbl(row Mod 23) * 1E-250
        data(row, 7) = IIf(row Mod 2 = 0, "one", "two")
        If row = 1 Then data(row, 8) = "header" Else data(row, 8) = original.RunID
    Next row
    sample.AddTable "Test.ColumnCodec", data
    store.SaveFile ThisWorkbook, sample
    Dim loaded As CCalculationSnapshot: Set loaded = store.LoadFile(ThisWorkbook)
    Dim actual As Variant: actual = loaded.Table("Test.ColumnCodec")
    TestSnapshotCheck "codec.mixed.bitExact", TestSnapshotSameMatrix(data, actual)
    TestSnapshotCheck "codec.emptyKinds", VarType(actual(1, 5)) = vbString And VarType(actual(1024, 5)) = vbEmpty
    TestSnapshotCheck "codec.booleanKind", VarType(actual(1, 3)) = vbBoolean And VarType(actual(1024, 3)) = vbBoolean
    store.SaveFile ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяющиеся ElementID сохраняются точно; любое нарушение шаблона включает общий кодек.
Private Sub TestSnapshotRepeatedElementIDs(ByVal original As CCalculationSnapshot, ByVal store As CCalculationSnapshotStore)
    Dim count As Long: count = original.TableRowCount("rngNDMSectionGeometry") - 1
    Dim groups As Long: groups = 2 + 4096 \ count
    Dim expected As Variant: expected = original.Table("rngNDMElementResults")
    Dim data() As Variant, row As Long, column As Long, sourceRow As Long
    ReDim data(1 To groups * count + 1, 1 To 9)
    For column = 1 To 9: data(1, column) = expected(1, column): Next column
    For row = 2 To UBound(data, 1)
        sourceRow = 2 + (row - 2) Mod count
        For column = 1 To 9: data(row, column) = expected(sourceRow, column): Next column
    Next row
    Dim testCase As Long, sample As CCalculationSnapshot, loaded As CCalculationSnapshot, key As Variant
    Dim changed As Long, savedID As String
    For testCase = 0 To 3
        changed = 0
        Select Case testCase
            Case 1: changed = UBound(data, 1): savedID = CStr(data(changed, 5)): data(changed, 5) = "TestMismatch_" & savedID
            Case 2: changed = 3: savedID = CStr(data(changed, 5)): data(changed, 5) = data(2, 5)
            Case 3: changed = 2: savedID = CStr(data(changed, 5)): data(changed, 5) = "ElementID"
        End Select
        Set sample = New CCalculationSnapshot: sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
        For Each key In original.Keys()
            If CStr(key) <> "rngNDMElementResults" Then sample.AddTable CStr(key), original.Table(CStr(key))
        Next key
        sample.AddTable "rngNDMElementResults", data
        store.SaveFile ThisWorkbook, sample: Set loaded = store.LoadFile(ThisWorkbook)
        TestSnapshotCheck "codec.repeatedIDs." & CStr(testCase) & ".exact", TestSnapshotSameMatrix(data, loaded.Table("rngNDMElementResults"))
        TestSnapshotCheck "codec.repeatedIDs." & CStr(testCase) & ".sourceUnchanged", TestSnapshotSameMatrix(data, sample.Table("rngNDMElementResults"))
        If changed > 0 Then data(changed, 5) = savedID
    Next testCase
    store.SaveFile ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторный расчёт сохранённого импорта не должен умножать автоматические подписи.
Private Sub TestSnapshotRepeatedImportedAnnotations(ByVal extension As String)
    modTestPerformance.TestConfigurePerformanceFixture "IMPORTED_CRACK_SP35", 3, "Newton", "Saved", "Normal", extension
    TestSnapshotSetting "Results.SnapshotStorage", "Excel"
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    TestSnapshotSetting "Plot.LoadCase", "PERF_01"
    Dim message As String: message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim first As Variant: first = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    Dim expected As Variant: expected = ReadPhysicalResultTable(ThisWorkbook, "rngNDMElementResults")
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    Dim actual As Variant: actual = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionAnnotations")
    TestSnapshotCheck "repeat.imported.annotationCount", UBound(first, 1) = UBound(actual, 1)
    Dim i As Long
    For i = 2 To UBound(first, 1): first(i, 1) = actual(i, 1): Next i
    TestSnapshotCheck "repeat.imported.annotationsExact", TestSnapshotSameMatrix(first, actual)
    actual = ReadElementResultTable(ThisWorkbook)
    For i = 2 To UBound(expected, 1): expected(i, 1) = actual(i, 1): Next i
    TestSnapshotCheck "repeat.imported.statesExact", TestSnapshotSameMatrix(expected, actual)
    TestSnapshotCheck "repeat.imported.contextReleased", TestCalculationSnapshotReadReleased()
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет текущие читаемые Excel-таблицы, отсутствие архива и неизменность верхнего вывода.
' Внешний runner вызывает метод только в своей частной рассчитанной копии.
Public Function TestExcelCurrentResultsRefresh() As String
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    On Error GoTo Failed
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    TestSnapshotCheck "excel.storage", store.StorageMode(ThisWorkbook) = "Excel"
    TestSnapshotCheck "excel.noArchiveName", Not TestSnapshotNameExists("rngNDMLoadCaseResults")
    TestSnapshotCheck "excel.noFileReference", Not store.HasFileReference(ThisWorkbook)
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    TestSnapshotCheck "excel.noPersistentFullSnapshot", snapshot Is Nothing
    Dim properties As Variant: properties = ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionProperties")
    TestSnapshotCheck "excel.readableLoadCaseProperties", Not TestSnapshotCommonOnly(properties)
    Dim names As Variant: names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    Dim widths As Variant: widths = Array(33, 49, 72, 84)
    Dim before(0 To 3) As Variant, i As Long
    Dim plot As String: plot = TestSnapshotPlot()
    Dim exported As String: exported = Audit02ReadExportSnapshotForTests(ThisWorkbook)
    TestSnapshotCorruptUpper ' Удалённые значения не должны восстанавливаться из скрытого архива.
    For i = 0 To 3: before(i) = ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Resize(32, CLng(widths(i))).Value2: Next i
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim iterations As String: iterations = settings.GetRequiredString("Solver.MaxIterations")
    Dim selectedName As String: selectedName = settings.GetRequiredString("Results.SnapshotName")
    TestSnapshotSetting "Results.SnapshotName", "../this_file_is_not_used_in_Excel"
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    UpdateSectionResultsForWorkbook ThisWorkbook
    TestSnapshotCheck "excel.refresh.noSolver", SectionEquilibriumSolveCount() = solves
    TestSnapshotCheck "excel.refresh.plotExact", TestSnapshotPlot() = plot
    TestSnapshotCheck "excel.refresh.exportExact", Audit02ReadExportSnapshotForTests(ThisWorkbook) = exported
    For i = 0 To 3
        TestSnapshotCheck "excel.refresh.keepsUpper." & CStr(names(i)), TestSnapshotSameMatrix(before(i), ThisWorkbook.Names.Item(CStr(names(i))).RefersToRange.Resize(32, CLng(widths(i))).Value2)
    Next i
    TestSnapshotCheck "excel.refresh.keepsProperties", TestSnapshotSameMatrix(properties, ReadPhysicalResultTable(ThisWorkbook, "rngNDMSectionProperties"))
    TestSnapshotCheck "excel.refresh.contextReleased", TestCalculationSnapshotReadReleased()
    TestSnapshotSetting "Solver.MaxIterations", iterations
    TestSnapshotSetting "Results.SnapshotName", selectedName
    GoTo Finished
Failed:
    TestSnapshotCheck "excel.runtime " & CStr(Err.Number) & ": " & Err.Description, False
Finished:
    TestExcelCurrentResultsRefresh = mTestReport & "TOTAL EXCEL CURRENT RESULTS: passed=" & CStr(mTestPassed) & "; failed=" & CStr(mTestFailed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет все шесть технических таблиц Excel с теми же готовыми массивами расчёта.
Private Sub TestSnapshotExcelCurrent(ByVal original As CCalculationSnapshot)
    Dim key As Variant
    TestSnapshotUpper original
    For Each key In original.BaseKeys()
        TestSnapshotCheck "excel.physical." & CStr(key), TestSnapshotSameMatrix(original.ResultTable(CStr(key)), ReadPhysicalResultTable(ThisWorkbook, CStr(key)))
    Next key
    Dim previousReport As String: previousReport = mTestReport
    Dim previousPassed As Long: previousPassed = mTestPassed
    Dim previousFailed As Long: previousFailed = mTestFailed
    Dim result As String: result = TestExcelCurrentResultsRefresh()
    mTestReport = previousReport & mTestReport: mTestPassed = previousPassed + mTestPassed: mTestFailed = previousFailed + mTestFailed
    If InStr(result, "failed=0") = 0 Then Err.Raise 5, , result
    SaveAndDisplayCalculationSnapshot ThisWorkbook, original
    TestSnapshotUpper original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отсутствие удалённого имени без создания диапазона или изменения книги.
Private Function TestSnapshotNameExists(ByVal name As String) As Boolean
    Dim target As Object
    On Error Resume Next
    Set target = ThisWorkbook.Names.Item(name)
    On Error GoTo 0
    TestSnapshotNameExists = Not target Is Nothing
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит реальный расчёт для внешней проверки Save/Close/Open и Save As.
' Runner вызывает его только в своей записываемой копии; пользовательские книги не участвуют.
Public Function TestPrepareSnapshotReopen(ByVal mode As String) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    modTestPerformance.TestConfigurePerformanceFixture "MIXED", 6, "Newton", "Circle", "Normal", settings.GetRequiredString("General.DiagramExtension")
    TestSnapshotSetting "Results.SnapshotStorage", mode
    TestSnapshotSetting "Results.SnapshotName", "TestReopen_" & ChrW$(&H3B1) & ".ndm"
    TestSnapshotSetting "General.ExecutionReportEnabled", "No"
    TestSnapshotSetting "Plot.LoadCase", "PERF_01"
    TestSnapshotSetting "AutoCAD.Export.CombinationID", "PERF_01"
    Dim message As String: message = FormSectionForWorkbook(ThisWorkbook)
    message = RunSectionCalculationForWorkbook(ThisWorkbook)
    TestPrepareSnapshotReopen = TestSnapshotReopenFingerprint()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' File восстанавливает повреждённый вывод; Excel обновляет схему без изменения сводок.
Public Function TestRestoreSnapshotAfterReopen() As String
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    If store.StorageMode(ThisWorkbook) = "File" Then TestSnapshotCorruptUpper
    TestSnapshotSetting "Solver.MaxIterations", "not_a_number"
    Dim solves As Long: solves = SectionEquilibriumSolveCount()
    UpdateSectionResultsForWorkbook ThisWorkbook
    If SectionEquilibriumSolveCount() <> solves Then Err.Raise 5, , "Восстановление повторно вызвало solver."
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    mTestReport = "": mTestPassed = 0: mTestFailed = 0
    If Not snapshot Is Nothing Then TestSnapshotUpper snapshot
    If mTestFailed <> 0 Then Err.Raise 5, , mTestReport
    If Not TestCalculationSnapshotReadReleased() Then Err.Raise 5, , "Контекст чтения не освобождён."
    TestRestoreSnapshotAfterReopen = TestSnapshotReopenFingerprint()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает все части снимка, данные схемы и экспорта для побитного сравнения после открытия.
Public Function TestSnapshotReopenFingerprint() As String
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(ThisWorkbook)
    Dim keys As Variant
    If snapshot Is Nothing Then
        keys = Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMElementResults", "rngNDMSectionContours", "rngNDMSectionAnnotations")
    Else
        keys = snapshot.Keys()
    End If
    Dim key As Variant, values As Variant, row As Long, column As Long, text As String, value As Variant
    Dim lines() As String, line As String
    If Not snapshot Is Nothing Then text = TestSnapshotDouble(snapshot.RunID) & "|" & TestSnapshotDouble(snapshot.ElapsedSeconds)
    For Each key In keys
        If snapshot Is Nothing Then values = ReadPhysicalResultTable(ThisWorkbook, CStr(key)) Else values = snapshot.Table(CStr(key))
        ReDim lines(0 To UBound(values, 1))
        lines(0) = CStr(key) & "|" & CStr(UBound(values, 1)) & "|" & CStr(UBound(values, 2))
        For row = 1 To UBound(values, 1)
            line = vbNullString
            For column = 1 To UBound(values, 2)
                value = values(row, column)
                If VarType(value) = vbString Or IsEmpty(value) Then
                    line = line & "S" & CStr(Len(CStr(value))) & ":" & CStr(value) & "|"
                Else
                    line = line & "D" & TestSnapshotDouble(CDbl(value)) & "|"
                End If
            Next column
            lines(row) = line
        Next row
        text = text & vbLf & Join(lines, vbLf)
    Next key
    TestSnapshotReopenFingerprint = text & vbLf & TestSnapshotPlot() & vbLf & Audit02ReadExportSnapshotForTests(ThisWorkbook)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все части одного логического снимка, включая каждое число IEEE 754.
Private Sub TestSnapshotEqual(ByVal label As String, ByVal expected As CCalculationSnapshot, ByVal actual As CCalculationSnapshot)
    Dim key As Variant
    TestSnapshotCheck label & ".run", TestSnapshotDouble(expected.RunID) = TestSnapshotDouble(actual.RunID)
    TestSnapshotCheck label & ".elapsed", TestSnapshotDouble(expected.ElapsedSeconds) = TestSnapshotDouble(actual.ElapsedSeconds)
    For Each key In expected.Keys()
        TestSnapshotCheck label & "." & CStr(key), actual.HasTable(CStr(key))
        If actual.HasTable(CStr(key)) Then TestSnapshotCheck label & ".values." & CStr(key), TestSnapshotSameMatrix(expected.Table(CStr(key)), actual.Table(CStr(key)))
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет физические верхние таблицы с сохранёнными данными, не решая НДС.
Private Sub TestSnapshotUpper(ByVal snapshot As CCalculationSnapshot)
    Dim key As Variant, name As String, values As Variant
    For Each key In Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
        Select Case CStr(key)
            Case "Summary": name = "rngBatchSummary"
            Case "Strength.Data": name = "rngStrengthSummaryAnchor"
            Case "Cracks.Data": name = "rngCrackSummaryAnchor"
            Case "Stability.Data": name = "rngStabilitySummaryAnchor"
        End Select
        values = snapshot.Table(CStr(key))
        TestSnapshotCheck "upper." & CStr(key), TestSnapshotSameMatrix(values, ThisWorkbook.Names.Item(name).RefersToRange.Resize(UBound(values, 1), UBound(values, 2)).Value2)
    Next key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Удаляет значения, чтобы обновление не могло пройти за счёт старого вывода.
Private Sub TestSnapshotCorruptUpper()
    Dim name As Variant
    For Each name In Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
        ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Resize(30, 3).ClearContents
    Next name
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Числа сравнивает побитно, Empty и пустую строку считает эквивалентным пустым Excel-полем.
Private Function TestSnapshotSameMatrix(ByVal expected As Variant, ByVal actual As Variant) As Boolean
    On Error GoTo Failed
    If UBound(expected, 1) <> UBound(actual, 1) Or UBound(expected, 2) <> UBound(actual, 2) Then Exit Function
    Dim i As Long, j As Long, a As Variant, b As Variant
    For i = 1 To UBound(expected, 1)
        For j = 1 To UBound(expected, 2)
            a = expected(i, j): b = actual(i, j)
            If VarType(a) = vbString Or VarType(b) = vbString Or IsEmpty(a) Or IsEmpty(b) Then
                If CStr(a) <> CStr(b) Then GoTo Different
            ElseIf TestSnapshotDouble(CDbl(a)) <> TestSnapshotDouble(CDbl(b)) Then
                GoTo Different
            End If
        Next j
    Next i
    TestSnapshotSameMatrix = True
    Exit Function
Different:
    mTestReport = mTestReport & "DIFF|row=" & CStr(i) & "|column=" & CStr(j) & "|expected=" & CStr(a) & "|actual=" & CStr(b) & vbCrLf
Failed:
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Формирует точное представление Double без зависимости от локали и десятичных цифр.
Private Function TestSnapshotDouble(ByVal value As Double) As String
    Dim number As TTestSnapshotDouble, bits As TTestSnapshotBits
    number.Value = value: LSet bits = number
    TestSnapshotDouble = Right$("00000000" & Hex$(bits.High), 8) & Right$("00000000" & Hex$(bits.Low), 8)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает реальные данные потребителя схемы, а не только наличие ChartObject.
Private Function TestSnapshotPlot() As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim reader As CSectionPlotDataReader: Set reader = New CSectionPlotDataReader: reader.LoadFromWorkbook ThisWorkbook, settings
    Dim text As String, i As Long
    text = reader.LoadCase & "|" & reader.StateType & "|" & reader.LengthUnit & "|" & reader.ResultUnit
    For i = 1 To reader.Count
        text = text & vbLf & reader.ElementID(i) & "|" & TestSnapshotDouble(reader.X(i)) & "|" & TestSnapshotDouble(reader.Y(i)) & "|" & _
            TestSnapshotDouble(reader.Area(i)) & "|" & TestSnapshotDouble(reader.ResultValue(i)) & "|" & reader.PhysicalState(i)
    Next i
    TestSnapshotPlot = text
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что LC-зависимые поля не остались в общей геометрической таблице.
Private Function TestSnapshotCommonOnly(ByVal values As Variant) As Boolean
    Dim i As Long
    For i = 2 To UBound(values, 1)
        If CStr(values(i, 2)) <> "ALL" Then Exit Function
        If CStr(values(i, 3)) = "LoadReferenceX" Or CStr(values(i, 3)) = "LoadReferenceY" Or InStr(1, CStr(values(i, 3)), "SP35.Slenderness", vbBinaryCompare) = 1 Then Exit Function
    Next i
    TestSnapshotCommonOnly = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Неудачная загрузка обязана завершиться явной ошибкой до перезаписи Results.
Private Function TestSnapshotRefreshRejected() As Boolean
    On Error GoTo Rejected
    UpdateSectionResultsForWorkbook ThisWorkbook
    Exit Function
Rejected:
    TestSnapshotRefreshRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет имя через тот же рабочий валидатор, что сохранение файла.
Private Function TestSnapshotPathRejected(ByVal store As CCalculationSnapshotStore) As Boolean
    On Error GoTo Rejected
    Dim path As String: path = store.SnapshotPath(ThisWorkbook)
    Exit Function
Rejected:
    TestSnapshotPathRejected = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Устанавливает настройку в частном fixture по ключу и реальному именованному диапазону.
Private Sub TestSnapshotSetting(ByVal key As String, ByVal value As String)
    Dim table As Object, field As String, column As Long: field = key: column = 2
    If Left$(key, 6) = "Units." Then
        Set table = ThisWorkbook.Names.Item("rngUnitSettings").RefersToRange
        field = Split(key, ".")(1): column = 4
    Else
        Set table = ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange
    End If
    Dim values As Variant: values = table.Value2
    Dim i As Long
    For i = 1 To UBound(values, 1)
        If CStr(values(i, 1)) = field Then table.Cells(i, column).Value2 = value: Exit Sub
    Next i
    Err.Raise 5, , "Не найдена тестовая настройка: " & key
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неизменённые исходные числа, включая малые моменты и ошибочную строку.
Private Sub TestSnapshotOriginalLoadMemory()
    Dim batch As CBatchSectionCalculator: Set batch = New CBatchSectionCalculator
    batch.AddCombination "RAW", -123.456#, 1E-20, -2E-20, "PR1", "Исходные числа"
    batch.AddInvalidCombination "BAD", "PR1", "Ошибка", "Некорректная нагрузка"
    TestSnapshotCheck "loads.original.available", batch.OriginalLoadsAvailable(1)
    TestSnapshotCheck "loads.original.force", TestSnapshotDouble(batch.OriginalN(1)) = TestSnapshotDouble(-123.456#)
    TestSnapshotCheck "loads.original.tinyMx", TestSnapshotDouble(batch.OriginalMx(1)) = TestSnapshotDouble(1E-20)
    TestSnapshotCheck "loads.original.tinyMy", TestSnapshotDouble(batch.OriginalMy(1)) = TestSnapshotDouble(-2E-20)
    TestSnapshotCheck "loads.invalid.notInvented", Not batch.OriginalLoadsAvailable(2)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает новые исходные числа с фактическим вводом, а рабочие — с таблицей устойчивости.
' Строка worst обязана принадлежать тому же LC, что указан в самом верху снимка.
Private Sub TestSnapshotSummaryLoads(ByVal snapshot As CCalculationSnapshot)
    Dim data As Variant: data = snapshot.Table("Summary")
    TestSnapshotCheck "loads.summary.width", UBound(data, 2) = 33
    TestSnapshotCheck "loads.summary.title", CStr(data(6, 28)) = "Нагрузки в точке приложения"
    TestSnapshotCheck "loads.summary.lineBreak", InStr(CStr(data(8, 31)), vbLf & "(если применимо)") > 0
    TestSnapshotCheck "loads.summary.originalMxComment", CStr(data(11, 29)) = "Исходный изгибающий момент Mx"
    TestSnapshotCheck "loads.summary.originalMyComment", CStr(data(11, 30)) = "Исходный изгибающий момент My"
    TestSnapshotCheck "loads.summary.forceNoEta", InStr(CStr(data(11, 31)), ChrW$(&H3B7)) = 0
    TestSnapshotCheck "loads.summary.momentsEta", InStr(CStr(data(11, 32)), ChrW$(&H3B7)) > 0
    TestSnapshotCheck "loads.summary.workingMxComment", CStr(data(11, 32)) = "Изгибающий момент Mx с учётом " & ChrW$(&H3B7) & " от продольного изгиба (если применимо)"
    TestSnapshotCheck "loads.summary.workingMyComment", CStr(data(11, 33)) = "Изгибающий момент My с учётом " & ChrW$(&H3B7) & " от продольного изгиба (если применимо)"
    Dim anchor As Object: Set anchor = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1)
    TestSnapshotCheck "loads.summary.workingMxSeparate", Not CBool(anchor.Cells(11, 32).MergeCells)
    TestSnapshotCheck "loads.summary.workingMySeparate", Not CBool(anchor.Cells(11, 33).MergeCells)
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem: Set units = New CUnitSystem: units.LoadFromSettings settings
    TestSnapshotCheck "loads.summary.forceUnit", CStr(data(10, 28)) = "N, " & units.OutputForceUnit
    TestSnapshotCheck "loads.summary.momentUnit", CStr(data(10, 29)) = "Mx, " & units.OutputMomentUnit
    Dim loads As Variant: loads = ThisWorkbook.Names.Item("rngLoadCombinations").RefersToRange.Value2
    Dim stable As Variant: stable = snapshot.Table("Stability.Data")
    Dim row As Long, inputRow As Long, stableRow As Long, column As Long, expected As Double
    For row = 13 To UBound(data, 1)
        If Len(CStr(data(row, 1))) = 0 Then GoTo NextSummaryLoad
        For inputRow = 2 To UBound(loads, 1)
            If CStr(loads(inputRow, 1)) = CStr(data(row, 1)) Then Exit For
        Next inputRow
        If inputRow > UBound(loads, 1) Then Err.Raise 5, , "Ввод тестового сочетания не найден."
        expected = units.InternalForceToOutput(units.InputForceToInternal(CDbl(loads(inputRow, 2))))
        TestSnapshotCheck "loads.summary.originalN." & CStr(row), TestSnapshotDouble(data(row, 28)) = TestSnapshotDouble(expected)
        expected = units.InternalMomentMxToOutput(units.InputMomentMxToInternal(CDbl(loads(inputRow, 3))))
        TestSnapshotCheck "loads.summary.originalMx." & CStr(row), TestSnapshotDouble(data(row, 29)) = TestSnapshotDouble(expected)
        expected = units.InternalMomentMyToOutput(units.InputMomentMyToInternal(CDbl(loads(inputRow, 4))))
        TestSnapshotCheck "loads.summary.originalMy." & CStr(row), TestSnapshotDouble(data(row, 30)) = TestSnapshotDouble(expected)
        For stableRow = 1 To UBound(stable, 1)
            If CStr(stable(stableRow, 1)) = CStr(data(row, 1)) And CStr(stable(stableRow, 3)) = "OK" Then
                For column = 0 To 2
                    TestSnapshotCheck "loads.summary.afterStability." & CStr(row) & "." & CStr(column), TestSnapshotDouble(data(row, 31 + column)) = TestSnapshotDouble(stable(stableRow, 12 + column))
                Next column
            End If
        Next stableRow
        If CStr(data(row, 1)) = CStr(data(1, 1)) Then
            For column = 28 To 33
                TestSnapshotCheck "loads.summary.worst." & CStr(column), TestSnapshotDouble(data(12, column)) = TestSnapshotDouble(data(row, column))
            Next column
        End If
NextSummaryLoad:
    Next row
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает прежнюю 26-колоночную сводку без подстановки текущих нагрузок Config.
Private Sub TestSnapshotLegacySummary(ByVal original As CCalculationSnapshot)
    Dim sample As CCalculationSnapshot: Set sample = New CCalculationSnapshot
    sample.RunID = original.RunID: sample.ElapsedSeconds = original.ElapsedSeconds
    Dim key As Variant, saved As Variant, data() As Variant, row As Long, column As Long
    For Each key In original.Keys()
        saved = original.Table(CStr(key))
        If CStr(key) = "Summary" Then
            ReDim data(1 To UBound(saved, 1), 1 To 26)
            For row = 1 To UBound(saved, 1)
                For column = 1 To 26: data(row, column) = saved(row, column): Next column
            Next row
            sample.AddTable CStr(key), data
        Else
            sample.AddTable CStr(key), saved
        End If
    Next key
    sample.Validate
    Dim writer As CBatchResultWriter: Set writer = New CBatchResultWriter
    writer.RestoreSummary ThisWorkbook, sample
    Dim target As Object: Set target = ThisWorkbook.Names.Item("rngBatchSummary").RefersToRange.Cells(1, 1)
    TestSnapshotCheck "loads.legacy.explicitMissing", CStr(target.Cells(11, 28).Value2) = "Не сохранено в этом снимке"
    TestSnapshotCheck "loads.legacy.noCurrentInput", Application.CountA(target.Offset(11, 27).Resize(31, 6)) = 0
    saved = original.Table("Summary")
    TestSnapshotCheck "loads.legacy.keepsWorst", CStr(target.Cells(1, 1).Value2) = CStr(saved(1, 1))
    writer.RestoreSummary ThisWorkbook, original
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет заданный горизонтальный порядок, не фиксируя адреса верхних таблиц.
Private Function TestSnapshotLowerOrder() As Boolean
    Dim name As Variant, last As Long, column As Long
    For Each name In Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMElementResults", "rngNDMSectionContours", "rngNDMSectionAnnotations")
        column = ThisWorkbook.Names.Item(CStr(name)).RefersToRange.Column
        If column <= last Then Exit Function
        last = column
    Next name
    TestSnapshotLowerOrder = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет компактный итог; успешные числа отдельных ячеек не раздувают журнал.
Private Sub TestSnapshotCheck(ByVal label As String, ByVal value As Boolean)
    If value Then mTestPassed = mTestPassed + 1 Else mTestFailed = mTestFailed + 1
    mTestReport = mTestReport & IIf(value, "PASS ", "FAIL ") & label & vbCrLf
End Sub
