Attribute VB_Name = "modAutoCADStressExport"
Option Explicit

' ==========================================================================
' Экспорт последнего расчета в AutoCAD
' ==========================================================================
' Модуль читает сохраненные таблицы Results и строит в AutoCAD области,
' подписи, главные оси, нейтральную линию и точку приложения нагрузки. Он не
' хранит модель сечения между макросами и не запускает AutoCAD import заново
' при экспорте.

Private Type TAutoCADExportSettings
    ConcreteLayer As String
    RebarLayer As String
    ConcreteTensionLayer As String
    ConcreteCompressionLayer As String
    RebarTensionLayer As String
    RebarCompressionLayer As String
    ConcreteTensionColor As Long
    ConcreteCompressionColor As Long
    RebarTensionColor As Long
    RebarCompressionColor As Long
    NeutralColor As Long
    IncludeElementNames As Boolean
    NeutralLineEnabled As Boolean
    PrincipalAxesMode As String
    LoadPointEnabled As Boolean
    ContourEnabled As Boolean
    ContourLayer As String
    OpeningContourLayer As String
    ExportCrackInteractionContour As Boolean
    CrackInteractionLayer As String
End Type

Private Const EXTENSION_WARNING_TEXT As String = "ВНЕ ФИЗИЧЕСКОЙ ДИАГРАММЫ МАТЕРИАЛА"
Private Const NUMERICAL_STATE_WARNING_TEXT As String = "ПРЯМОЕ НДС НЕ СОШЛОСЬ"
Private Const EXTENSION_WARNING_LAYER As String = "RC_NDM_Warnings"
Private Const SECTION_CONTOUR_COLOR_INDEX As Long = 30 ' Наружная граница бетона: оранжевый ACI 30.
Private Const OPENING_CONTOUR_COLOR_INDEX As Long = 4 ' Контуры отверстий: голубой ACI 4.
Private Const CRACK_REGION_COLOR_INDEX As Long = 31 ' Расчетная область взаимодействия: ACI 31.
Private Const CONTOUR_POINT_TOLERANCE As Double = 0.000001
Private mSnapshotUnits As CUnitSystem ' Пересчет по явным единицам Results; не загружается из текущего Config.
Private mSnapshotReadTables As Object ' Только подготовка одного export-state; освобождается до рисования/выхода с ошибкой.
Private mSnapshotReadWorkbook As Object ' Identity книги защищает локальное чтение от чужого контекста.
Private mSnapshotTableReadCount As Long ' ДЛЯ ТЕСТОВ: фактические Value2-чтения последней операции.

' Экспортирует выбранное сохраненное НДС и его оформление в активный чертеж.
' Читает Results, профиль отображения и единицы; решатель не вызывается,
' ошибки доступа или отсутствующего snapshot показываются пользователю.
Public Sub ExportSectionStressToAutoCAD()
    On Error GoTo Failed

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    ' Неверный флаг сообщения должен быть найден до изменения чертежа.
    Dim informationEnabled As Boolean
    informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled")
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings

    Dim exportSettings As TAutoCADExportSettings
    exportSettings = ReadAutoCADExportSettings(settings)

    Dim section As CSectionModel
    Dim resultByID As Object
    Dim physicalStateByID As Object
    Dim combinationID As String
    Dim profileId As String
    Dim stateType As String
    Dim quantity As String
    Dim resultPrecision As Long
    Dim epsilon0 As Double
    Dim kappaX As Double
    Dim kappaY As Double
    Dim loadReferenceX As Double
    Dim loadReferenceY As Double
    Dim centroidX As Double
    Dim centroidY As Double
    Dim principalAngle As Double
    Dim extensionUsed As Boolean
    Dim stateWarningText As String
    ReadResultsExportState ThisWorkbook, settings, units, exportSettings.PrincipalAxesMode, section, resultByID, physicalStateByID, combinationID, _
        profileId, stateType, quantity, resultPrecision, _
        epsilon0, kappaX, kappaY, loadReferenceX, loadReferenceY, _
        centroidX, centroidY, principalAngle, extensionUsed, stateWarningText

    Dim crackRows As Object, crackSweeps As Object
    If exportSettings.ExportCrackInteractionContour And IsCrackExportState(stateType) Then
        Set crackRows = CurrentCrackRegionRows(ThisWorkbook, combinationID)
        Set crackSweeps = ReadSavedContourArcSweeps(ThisWorkbook, crackRows, "CRACK_REGION_")
    End If

    Dim contourExportCount As Long
    DrawResultsStressExport section, resultByID, physicalStateByID, epsilon0, kappaX, kappaY, _
        loadReferenceX, loadReferenceY, centroidX, centroidY, principalAngle, resultPrecision, stateWarningText, exportSettings, _
        contourExportCount, , crackRows, crackSweeps
    If informationEnabled Then
        ShowWorkbookMessage "Экспорт в AutoCAD завершен. Волокон бетона: " & CStr(section.ConcreteCount) & _
            "; стержней арматуры: " & CStr(section.RebarCount) & _
            "; контурных полилиний: " & ContourExportStatusText(exportSettings.ContourEnabled, contourExportCount) & _
            "; сочетание: " & combinationID & _
            "; профиль: " & profileId & _
            "; состояние: " & stateType & _
            "; величина: " & quantity, vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage "Экспорт в AutoCAD не выполнен: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

' Очищает активный чертеж AutoCAD от объектов оформления, созданных NDM-export.
' Геометрические Region бетона и арматуры остаются на своих слоях. Удаляются
' только подписи на annotation-слоях, нейтральная линия, главные оси и точка
' приложения нагрузки, чтобы можно было заново выгрузить расчет поверх той же
' геометрии.
Public Sub ClearAutoCADDrawing()
    On Error GoTo Failed

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    ' Используем тот же снимок Config и проверяем флаг до удаления объектов.
    Dim informationEnabled As Boolean
    informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled")

    Dim exportSettings As TAutoCADExportSettings
    exportSettings = ReadAutoCADExportSettings(settings, True)

    Dim acad As Object
    Set acad = ConnectToRunningAutoCAD()

    Dim doc As Object
    Set doc = ActiveAutoCADDocument(acad)

    Dim deletedCount As Long
    deletedCount = DeleteAutoCADEntitiesOnLayers(doc.ModelSpace, AutoCADCleanupLayerSet(exportSettings))
    doc.Regen 1

    If informationEnabled Then
        ShowWorkbookMessage "Чертеж AutoCAD очищен от объектов оформления RC Section NDM." & vbCrLf & _
            "Удалено объектов: " & CStr(deletedCount) & "." & vbCrLf & _
            "Геометрия бетона и арматуры оставлена без изменений.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage "Очистка чертежа AutoCAD не выполнена: " & Err.Description, vbExclamation, "RC Section NDM"
End Sub

' Собирает перечень слоев, которые относятся к оформлению, а не к геометрии.
' Слои подписей берутся из Config, а технические слои осей/точки/нейтральной
' линии совпадают с теми, которые создает экспорт.
Private Function AutoCADCleanupLayerSet(ByRef exportSettings As TAutoCADExportSettings) As Object
    Dim layers As Object
    Set layers = CreateObject("Scripting.Dictionary")
    layers.CompareMode = vbTextCompare

    AddCleanupLayer layers, exportSettings.ConcreteTensionLayer
    AddCleanupLayer layers, exportSettings.ConcreteCompressionLayer
    AddCleanupLayer layers, exportSettings.RebarTensionLayer
    AddCleanupLayer layers, exportSettings.RebarCompressionLayer
    AddCleanupLayer layers, "RC_NDM_Axes"
    AddCleanupLayer layers, "RC_NDM_LoadPoint"
    AddCleanupLayer layers, "RC_NDM_NeutralLine"
    AddCleanupLayer layers, exportSettings.CrackInteractionLayer
    AddCleanupLayer layers, EXTENSION_WARNING_LAYER
    ' Общие слои защищаются даже при совпадении с другим слоем оформления.
    If Len(exportSettings.ContourLayer) > 0 Then layers.Item(exportSettings.ContourLayer) = False
    If Len(exportSettings.OpeningContourLayer) > 0 Then layers.Item(exportSettings.OpeningContourLayer) = False

    Set AutoCADCleanupLayerSet = layers
End Function

' Добавляет слой в set, пропуская пустые значения.
Private Sub AddCleanupLayer(ByVal layers As Object, ByVal layerName As String)
    layerName = Trim$(layerName)
    If Len(layerName) = 0 Then Exit Sub
    If Not layers.Exists(layerName) Then layers.Add layerName, True
End Sub

' Проходит ModelSpace с конца, чтобы безопасно удалять найденные объекты.
' Region всегда сохраняются. Остальные объекты удаляются только со слоев
' оформления, поэтому пользовательские объекты не следует размещать на них.
Private Function DeleteAutoCADEntitiesOnLayers(ByVal ms As Object, ByVal layers As Object) As Long
    Dim i As Long
    For i = ms.Count - 1 To 0 Step -1
        Dim entity As Object
        Set entity = ms.Item(i)
        If layers.Exists(CStr(entity.Layer)) And StrComp(CStr(entity.ObjectName), "AcDbRegion", vbTextCompare) <> 0 Then
            If CBool(layers.Item(CStr(entity.Layer))) Or IsNDMContourOutput(entity) Then
                entity.Delete
                DeleteAutoCADEntitiesOnLayers = DeleteAutoCADEntitiesOnLayers + 1
            End If
        End If
    Next i
End Function

' Собирает контекст экспорта из последнего snapshot: геометрию, значения
' элементов, плоскость деформаций и предупреждение о физической допустимости.
' Профиль выбирает state/величину; текущий ввод геометрии не заменяет Results.
Private Sub ReadResultsExportState(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, ByVal principalAxesMode As String, _
        ByRef section As CSectionModel, ByRef resultByID As Object, ByRef physicalStateByID As Object, ByRef combinationID As String, _
        ByRef profileId As String, ByRef stateType As String, ByRef quantity As String, ByRef resultPrecision As Long, _
        ByRef epsilon0 As Double, ByRef kappaX As Double, ByRef kappaY As Double, _
        ByRef loadReferenceX As Double, ByRef loadReferenceY As Double, _
        ByRef centroidX As Double, ByRef centroidY As Double, ByRef principalAngle As Double, _
        ByRef extensionUsed As Boolean, ByRef stateWarningText As String)
    On Error GoTo Failed
    Set mSnapshotReadTables = CreateObject("Scripting.Dictionary")
    Set mSnapshotReadWorkbook = workbook
    mSnapshotTableReadCount = 0
    Set section = ReadSectionGeometryFromResults(workbook, "Results")

    combinationID = ResolveExportCombinationID(workbook, settings.GetRequiredString("AutoCAD.Export.CombinationID"))
    profileId = ReadProfileIdForLoadCase(workbook, combinationID)
    If Len(profileId) = 0 Then Err.Raise vbObjectError + 4366, "ReadResultsExportState", _
        settings.InputErrorMessage("AutoCAD.Export.CombinationID", _
            "В сохраненных Results не найдено сочетание " & combinationID & " с расчетным профилем.", _
            "Выберите Worst или ID сохраненного сочетания; для новых нагрузок сначала выполните расчет.")

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook

    Dim profile As CCalculationProfile
    Set profile = profiles.ProfileById(profileId)
    stateType = profile.VisualizationStateText
    quantity = VisualizationQuantityToText(profile.VisualizationQuantity)
    resultPrecision = profile.VisualizationPrecisionForQuantity(profile.VisualizationQuantity)

    Set resultByID = CreateObject("Scripting.Dictionary")
    resultByID.CompareMode = vbTextCompare
    Set physicalStateByID = CreateObject("Scripting.Dictionary")
    physicalStateByID.CompareMode = vbTextCompare
    ReadElementResultsForCombination workbook, quantity, combinationID, stateType, resultByID, physicalStateByID
    ReadSectionPropertiesForCombination workbook, units, combinationID, stateType, principalAxesMode, epsilon0, kappaX, kappaY, _
        loadReferenceX, loadReferenceY, centroidX, centroidY, principalAngle, extensionUsed, stateWarningText
    Set mSnapshotReadTables = Nothing: Set mSnapshotReadWorkbook = Nothing
    Exit Sub
Failed:
    Dim errorNumber As Long, errorSource As String, errorText As String
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    Set mSnapshotReadTables = Nothing: Set mSnapshotReadWorkbook = Nothing
    Err.Raise errorNumber, errorSource, errorText
End Sub

' Восстанавливает CSectionModel из таблицы Results в внутренних единицах.
' Сохраняет реальные A/I импортированных элементов, отдельно от оболочки
' отображения; пустой snapshot отклоняется до чтения строк.
Public Function ReadSectionGeometryFromResults(ByVal workbook As Object, Optional ByVal sourceType As String = "Results") As CSectionModel
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4350, "ReadSectionGeometryFromResults", _
        "На листе Results нет таблицы расчетной геометрии. Сначала выполните расчет."

    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colType As Long: colType = ResultColumn(data, "MaterialType")
    Dim colX As Long: colX = ResultColumn(data, "X")
    Dim colY As Long: colY = ResultColumn(data, "Y")
    Dim colArea As Long: colArea = ResultColumn(data, "Area")
    Dim colShape As Long: colShape = GeometryStatusColumn(data)
    Dim colWidth As Long: colWidth = ResultColumn(data, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(data, "Height")
    Dim colDiameter As Long: colDiameter = ResultColumn(data, "Diameter")
    Dim colRotation As Long: colRotation = ResultColumn(data, "Rotation")
    Dim colLocalIx As Long: colLocalIx = ResultColumn(data, "LocalIx")
    Dim colLocalIy As Long: colLocalIy = ResultColumn(data, "LocalIy")
    Dim colLocalIxy As Long: colLocalIxy = ResultColumn(data, "LocalIxy")
    Dim colComment As Long: colComment = ResultColumn(data, "Comment")
    Dim lengthUnit As String: lengthUnit = ResultHeaderUnit(data, colX, "mm")
    Dim areaUnit As String: areaUnit = ResultHeaderUnit(data, colArea, "mm2")
    Dim fourthUnit As String: fourthUnit = ResultHeaderUnit(data, colLocalIx, lengthUnit & "4")

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = sourceType

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Len(Trim$(CStr(data(rowIndex, colID)))) > 0 Then
            If StrComp(CStr(data(rowIndex, colType)), "Concrete", vbTextCompare) = 0 Then
                ' Results хранит окончательный угол оболочки: повторный export
                ' не заменяет горизонтальную грань средним направлением соседей.
                model.AddConcreteElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                    OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), 1, vbNullString, vbNullString, _
                    CStr(data(rowIndex, colShape)), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colWidth, True), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colHeight, True), lengthUnit), ReadGeometryNumber(data, anchor, rowIndex, colRotation, True), _
                    CStr(data(rowIndex, colComment)), OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIx, True), fourthUnit), _
                    OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIy, True), fourthUnit), OutputFourthPowerLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colLocalIxy, True), fourthUnit), True
            ElseIf StrComp(CStr(data(rowIndex, colType)), "Rebar", vbTextCompare) = 0 Then
                model.AddRebarElement OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colX), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colY), lengthUnit), _
                    OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colDiameter), lengthUnit), OutputAreaToInternalByUnit(ReadGeometryNumber(data, anchor, rowIndex, colArea), areaUnit), _
                    vbNullString, 1, vbNullString, _
                    vbNullString, CStr(data(rowIndex, colComment))
            End If
        End If
    Next rowIndex

    model.ApplyAverageRotationToEquivalentAreaFallbacks
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4351, "ReadSectionGeometryFromResults", _
        "В таблице Results нет бетонных элементов."
    ReadSavedSectionContours workbook, model.Contours
    Set ReadSectionGeometryFromResults = model
End Function

' Читает только компактный v1-снимок. Отсутствующая версия или повреждение
' явно диагностируются; устаревшие строки Geometry никогда не подмешиваются.
' Единицы и SectionXY принадлежат снимку, а не текущему Config.
Public Function ReadSavedSectionContours(ByVal workbook As Object, _
        Optional ByVal contours As CSectionContours = Nothing) As CSectionContours
    Dim anchor As Object, data As Variant, headers As Variant, i As Long, row As Long
    On Error GoTo MissingAnchor
    Set anchor = workbook.Names.Item("rngNDMSectionContours").RefersToRange
    On Error GoTo 0
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionContours")
    headers = Array("RunID v1", "LoopID", "SegmentID", "Sequence", "LoopRole", "SegmentType", _
        "StartX", "StartY", "EndX", "EndY", "CenterX", "CenterY", "Radius", "SweepAngle, rad", "SourceID", "Comment", "LengthUnit (SectionXY)")
    If Not IsArray(data) Then GoTo InvalidHeader
    If UBound(data, 2) <> 17 Then GoTo InvalidHeader
    For i = 0 To UBound(headers)
        If IsError(data(1, i + 1)) Then GoTo InvalidHeader
        If CStr(data(1, i + 1)) <> CStr(headers(i)) Then GoTo InvalidHeader
    Next i
    Dim restored As CSectionContours: Set restored = New CSectionContours
    Dim kind As String, id As String, loopID As String, role As String, source As String, comment As String, unitText As String
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, radius As Double, sweep As Double, angle As Double
    Dim ids As Object: Set ids = CreateObject("Scripting.Dictionary")
    Dim snapshotRun As Variant, geometryAnchor As Object
    Set geometryAnchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    If CStr(geometryAnchor.Value2) = "RunID" Then snapshotRun = geometryAnchor.Offset(1, 0).Value2
    For row = 2 To UBound(data, 1)
        If IsError(data(row, 1)) Or IsNull(data(row, 1)) Then GoTo InvalidRow
        If Len(CStr(data(row, 1))) = 0 Then GoTo InvalidRow
        If Not IsEmpty(snapshotRun) Then
            If CStr(data(row, 1)) <> CStr(snapshotRun) Then GoTo InvalidRow
        Else
            snapshotRun = data(row, 1)
        End If
        id = ReadSavedAnnotationText(data, anchor, row, 3)
        loopID = ReadSavedAnnotationText(data, anchor, row, 2): role = ReadSavedAnnotationText(data, anchor, row, 5)
        kind = ReadSavedAnnotationText(data, anchor, row, 6): source = ReadSavedAnnotationText(data, anchor, row, 15)
        comment = ReadSavedAnnotationText(data, anchor, row, 16): unitText = ReadSavedAnnotationText(data, anchor, row, 17)
        If Len(Trim$(id)) = 0 Or Len(Trim$(loopID)) = 0 Or Len(Trim$(unitText)) = 0 Then GoTo InvalidRow
        If ids.Exists(id) Then GoTo InvalidRow
        ids.Add id, True
        If ReadGeometryNumber(data, anchor, row, 4) <> row - 1 Then GoTo InvalidRow
        If StrComp(role, "Outer", vbTextCompare) <> 0 And StrComp(role, "Opening", vbTextCompare) <> 0 Then GoTo InvalidRow
        x1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 7), unitText)
        y1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 8), unitText)
        Select Case kind
            Case "CONTOUR_LINE"
                x2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 9), unitText)
                y2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 10), unitText)
                restored.AddContourLine id, x1, y1, x2, y2, comment, loopID, role, source
            Case "CONTOUR_CIRCLE"
                radius = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 13), unitText)
                restored.AddContourCircle id, x1, y1, radius, comment, loopID, role, source
            Case "CONTOUR_ARC"
                cx = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 11), unitText)
                cy = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 12), unitText)
                sweep = ReadContourArcSweep(data(row, 14), anchor.Parent.Name & "!" & anchor.Offset(row - 1, 13).Address(False, False))
                radius = Sqr((x1 - cx) ^ 2 + (y1 - cy) ^ 2)
                ' Поворот вектора начала однозначно восстанавливает конец, включая большие и отрицательные дуги.
                x2 = cx + (x1 - cx) * Cos(sweep) - (y1 - cy) * Sin(sweep)
                y2 = cy + (x1 - cx) * Sin(sweep) + (y1 - cy) * Cos(sweep)
                restored.AddContourArc id, x1, y1, x2, y2, cx, cy, radius, sweep, comment, loopID, role, source
            Case Else
                GoTo InvalidRow
        End Select
    Next row
    If contours Is Nothing Then
        Set contours = restored
    Else
        contours.Clear
        CopySectionContours restored, contours
    End If
    Set ReadSavedSectionContours = contours
    Exit Function
MissingAnchor:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "В книге отсутствует rngNDMSectionContours. Обновите формат сохраненного снимка Results."
InvalidHeader:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "Неверная или пустая шапка контуров v1: " & anchor.Parent.Name & "!" & anchor.Address(False, False) & ". Восстановите снимок; старый блок не используется."
InvalidRow:
    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", "Неверная запись контура: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, 0).Resize(1, 17).Address(False, False) & ". Проверьте ID, порядок, роль, тип и единицы."
End Function

' Переносит проверенный контейнер без геометрической классификации.
Private Sub CopySectionContours(ByVal source As CSectionContours, ByVal target As CSectionContours)
    Dim i As Long
    For i = 1 To source.Count
        Select Case source.SegmentType(i)
            Case "CONTOUR_LINE"
                target.AddContourLine source.ContourID(i), source.StartX(i), source.StartY(i), source.EndX(i), source.EndY(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
            Case "CONTOUR_CIRCLE"
                target.AddContourCircle source.ContourID(i), source.StartX(i), source.StartY(i), source.Radius(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
            Case "CONTOUR_ARC"
                target.AddContourArc source.ContourID(i), source.StartX(i), source.StartY(i), source.EndX(i), source.EndY(i), source.CenterX(i), source.CenterY(i), source.Radius(i), source.SweepAngle(i), source.Comment(i), source.LoopID(i), source.LoopRole(i), source.SourceID(i)
        End Select
    Next i
End Sub

' Обновляет только нижнюю полосу сохраненного Results. До мутации полностью
' читает старые элементы/контуры и проверяет аналитические дуги. Ячейки справа
' сдвигаются Insert внутри этой полосы, не удаляются и не затрагивают верхние
' результаты. Повторный вызов проверяет v1 и ничего не перемещает.
Public Function MigrateSavedSectionContours(Optional ByVal workbook As Object = Nothing) As String
    If workbook Is Nothing Then Set workbook = ThisWorkbook
    Dim existing As Object, contourName As Object
    On Error Resume Next
    Set contourName = workbook.Names.Item("rngNDMSectionContours")
    On Error GoTo 0
    Dim contours As CSectionContours
    If Not contourName Is Nothing Then
        ' Существующее, но поврежденное имя не считается отсутствующей версией.
        Set existing = contourName.RefersToRange
        Set contours = ReadSavedSectionContours(workbook)
        MigrateSavedSectionContours = "Contours v1: already current; segments=" & CStr(contours.Count)
        Exit Function
    End If
    Set contours = ReadLegacySectionContours(workbook)
    Dim geometry As Object, properties As Object, ws As Object, data As Variant, elements() As Variant
    Set geometry = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Set properties = workbook.Names.Item("rngNDMSectionProperties").RefersToRange
    Set ws = geometry.Worksheet
    If ws.ProtectContents Or workbook.ProtectStructure Then Err.Raise vbObjectError + 4377, "MigrateSavedSectionContours", _
        "Не удалось обновить снимок Results: снимите защиту листа Results и структуры книги, затем повторите обновление. Старый снимок не изменен."
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    Dim count As Long, row As Long, col As Long, oldRows As Long, unitText As String, runID As Variant
    unitText = "mm": runID = Empty: oldRows = 1
    If IsArray(data) Then
        oldRows = UBound(data, 1)
        If UBound(data, 2) <> 15 And UBound(data, 2) <> 25 Then Err.Raise vbObjectError + 4377, "MigrateSavedSectionContours", "Неизвестный старый формат геометрии Results."
        unitText = ResultHeaderUnit(data, ResultColumn(data, "X"), "mm")
        For row = 2 To oldRows
            If SafeText(data(row, 3)) <> "Contour" Then count = count + 1
        Next row
        ReDim elements(1 To count + 1, 1 To 15)
        For col = 1 To 15: elements(1, col) = data(1, col): Next col
        count = 1
        For row = 2 To oldRows
            If SafeText(data(row, 3)) <> "Contour" Then
                count = count + 1
                For col = 1 To 15: elements(count, col) = data(row, col): Next col
            End If
            If IsEmpty(runID) Then runID = data(row, 1)
        Next row
    End If
    Dim i As Long, ex As Double, ey As Double, tolerance As Double
    For i = 1 To contours.Count
        If contours.SegmentType(i) = "CONTOUR_ARC" Then
            ex = contours.CenterX(i) + (contours.StartX(i) - contours.CenterX(i)) * Cos(contours.SweepAngle(i)) - (contours.StartY(i) - contours.CenterY(i)) * Sin(contours.SweepAngle(i))
            ey = contours.CenterY(i) + (contours.StartX(i) - contours.CenterX(i)) * Sin(contours.SweepAngle(i)) + (contours.StartY(i) - contours.CenterY(i)) * Cos(contours.SweepAngle(i))
            tolerance = 0.0000001 * MaxDouble(1#, contours.Radius(i))
            If Abs(ex - contours.EndX(i)) > tolerance Or Abs(ey - contours.EndY(i)) > tolerance Then Err.Raise vbObjectError + 4377, "MigrateSavedSectionContours", "Конец сохраненной дуги не согласован с центром и углом; старый снимок оставлен без изменений."
        End If
    Next i
    Dim name As Variant, block As Object, bottom As Long, shift As Long, contourColumn As Long
    bottom = geometry.Row + oldRows - 1
    For Each name In Array("rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations")
        Set block = workbook.Names.Item(CStr(name)).RefersToRange
        If Not block.Worksheet Is ws Or block.Row <> geometry.Row Then Err.Raise vbObjectError + 4377, "MigrateSavedSectionContours", "Нижние якоря Results должны находиться в одной строке одного листа; снимок оставлен без изменений."
        i = block.Row + AnchoredRowCount(block) - 1
        If i > bottom Then bottom = i
    Next name
    contourColumn = geometry.Column + 15 + 2
    shift = contourColumn + 17 + 2 - properties.Column
    ' Новый формат сначала проверяется на временном листе: новый и старый
    ' footprints частично пересекаются, поэтому нельзя чистить старые поля раньше.
    Dim stage As Object, writer As CNDMResultsWriter, restored As CSectionContours
    Set stage = workbook.Worksheets.Add
    Set writer = New CNDMResultsWriter
    workbook.Names.Add "rngNDMSectionContours", "='" & Replace(stage.Name, "'", "''") & "'!$A$2"
    On Error GoTo StageFailed
    writer.WriteSectionContours workbook, contours, runID, Nothing, unitText
    Set restored = ReadSavedSectionContours(workbook)
    If restored.Count <> contours.Count Then Err.Raise vbObjectError + 4377, "MigrateSavedSectionContours", "Не совпало количество перенесенных сегментов."
    workbook.Names.Item("rngNDMSectionContours").Delete
    On Error GoTo 0
    If shift > 0 Then ws.Cells(geometry.Row - 1, properties.Column).Resize(bottom - geometry.Row + 2, shift).Insert -4161
    ' Старый источник удаляется только после успешного обратного чтения нового.
    geometry.Resize(oldRows, 15).ClearContents
    geometry.Offset(0, 15).Resize(oldRows, 10).Clear
    If count < oldRows Then geometry.Offset(count, 0).Resize(oldRows - count, 15).Clear
    geometry.Offset(-1, 0).Resize(1, 25).UnMerge: geometry.Offset(-1, 0).Resize(1, 25).Clear
    If IsArray(data) Then
        geometry.Resize(UBound(elements, 1), 15).Value2 = elements
        writer.FormatResultsBlock geometry, UBound(elements, 1), 15, "Геометрия расчетных элементов сечения"
    Else
        writer.FormatResultsBlock geometry, 1, 15, "Геометрия расчетных элементов сечения"
    End If
    workbook.Names.Add "rngNDMSectionContours", "='" & Replace(ws.Name, "'", "''") & "'!" & ws.Cells(geometry.Row, contourColumn).Address
    writer.WriteSectionContours workbook, restored, runID, Nothing, unitText
    Set restored = ReadSavedSectionContours(workbook)
    Dim alerts As Boolean: alerts = workbook.Application.DisplayAlerts
    workbook.Application.DisplayAlerts = False: stage.Delete: workbook.Application.DisplayAlerts = alerts
    MigrateSavedSectionContours = "Contours v1: migrated; segments=" & CStr(restored.Count) & "; shifted columns=" & CStr(MaxDouble(0#, shift))
    Exit Function
StageFailed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    workbook.Names.Item("rngNDMSectionContours").Delete
    alerts = workbook.Application.DisplayAlerts
    workbook.Application.DisplayAlerts = False: stage.Delete: workbook.Application.DisplayAlerts = alerts
    On Error GoTo 0
    Err.Raise errorNumber, "MigrateSavedSectionContours", errorText & " Старый снимок не изменен."
End Function

' Узкий reader старых 25 колонок используется только миграцией до очистки.
Private Function ReadLegacySectionContours(ByVal workbook As Object, _
        Optional ByVal contours As CSectionContours = Nothing) As CSectionContours
    If contours Is Nothing Then Set contours = New CSectionContours
    contours.Clear
    Set ReadLegacySectionContours = contours
    Dim anchor As Object, data As Variant
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasResultTableRows(data) Then Exit Function
    Dim colMaterial As Long: colMaterial = ResultColumn(data, "MaterialType")
    Dim row As Long, hasContours As Boolean
    For row = 2 To UBound(data, 1)
        If StrComp(SafeText(data(row, colMaterial)), "Contour", vbTextCompare) = 0 Then hasContours = True
    Next row
    If Not hasContours Then Exit Function

    Dim colType As Long: colType = GeometryStatusColumn(data)
    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colStartX As Long: colStartX = ResultColumn(data, "X")
    Dim colStartY As Long: colStartY = ResultColumn(data, "Y")
    Dim colEndX As Long: colEndX = ResultColumn(data, "EndX")
    Dim colEndY As Long: colEndY = ResultColumn(data, "EndY")
    Dim colCenterX As Long: colCenterX = ResultColumn(data, "CenterX")
    Dim colCenterY As Long: colCenterY = ResultColumn(data, "CenterY")
    Dim colSweep As Long: colSweep = ResultColumn(data, "SweepAngle")
    Dim colRadius As Long: colRadius = ResultColumn(data, "Radius")
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")
    Dim colComment As Long: colComment = ResultColumn(data, "Comment")
    Dim colLoop As Long: colLoop = ResultColumn(data, "LoopID")
    Dim colRole As Long: colRole = ResultColumn(data, "LoopRole")
    Dim colSource As Long: colSource = ResultColumn(data, "SourceID")
    Dim kind As String, contourID As String, unitText As String, comment As String
    Dim loopID As String, role As String, source As String
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, radius As Double, sweep As Double
    For row = 2 To UBound(data, 1)
        If StrComp(ReadSavedAnnotationText(data, anchor, row, colMaterial), "Contour", vbTextCompare) = 0 Then
            kind = UCase$(Trim$(ReadSavedAnnotationText(data, anchor, row, colType)))
            contourID = ReadSavedAnnotationText(data, anchor, row, colID)
            comment = ReadSavedAnnotationText(data, anchor, row, colComment)
            loopID = ReadSavedAnnotationText(data, anchor, row, colLoop)
            role = ReadSavedAnnotationText(data, anchor, row, colRole)
            source = ReadSavedAnnotationText(data, anchor, row, colSource)
            unitText = Trim$(ReadSavedAnnotationText(data, anchor, row, colUnit))
            If Len(unitText) = 0 Or unitText = "-" Or Len(Trim$(loopID)) = 0 Or _
                    (StrComp(role, "Outer", vbTextCompare) <> 0 And StrComp(role, "Opening", vbTextCompare) <> 0) Then
                Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", _
                    "В сохраненной геометрии Results не заданы единицы, ID замкнутого контура или его роль: " & _
                    anchor.Parent.Name & "!" & anchor.Offset(row - 1, colLoop - 1).Resize(1, 4).Address(False, False) & _
                    ". Повторите импорт геометрии или расчет."
            End If
            x1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colStartX), unitText)
            y1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colStartY), unitText)
            x2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colEndX), unitText)
            y2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colEndY), unitText)
            Select Case kind
                Case "CONTOUR_LINE"
                    contours.AddContourLine contourID, x1, y1, x2, y2, comment, loopID, role, source
                Case "CONTOUR_CIRCLE"
                    radius = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colRadius), unitText)
                    contours.AddContourCircle contourID, x1, y1, radius, comment, loopID, role, source
                Case "CONTOUR_ARC"
                    cx = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colCenterX), unitText)
                    cy = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colCenterY), unitText)
                    radius = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, colRadius), unitText)
                    sweep = ReadContourArcSweep(data(row, colSweep), anchor.Parent.Name & "!" & _
                        anchor.Offset(row - 1, colSweep - 1).Address(False, False))
                    contours.AddContourArc contourID, x1, y1, x2, y2, cx, cy, radius, sweep, comment, loopID, role, source
                Case Else
                    Err.Raise vbObjectError + 4376, "ReadSavedSectionContours", _
                        "Неизвестный тип сохраненного контура " & kind & ": " & anchor.Parent.Name & "!" & _
                        anchor.Offset(row - 1, colType - 1).Address(False, False) & ". Повторите импорт геометрии."
            End Select
        End If
    Next row
End Function

' Читает текст material-аннотации без превращения ошибки формулы в пустоту.
' Адрес вычисляется от текущего именованного якоря, а не фиксируется в коде.
Private Function ReadSavedAnnotationText(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long) As String
    If IsError(data(row, column)) Or IsNull(data(row, column)) Then Err.Raise vbObjectError + 4376, _
        "ReadSectionGeometryFromResults", "В сохраненном контуре Results поле " & CStr(data(1, column)) & _
        " содержит ошибку: " & anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Исправьте ячейку или повторите импорт геометрии."
    ReadSavedAnnotationText = CStr(data(row, column))
End Function

' Возвращает диагностическую метку источника геометрии, сохраненную в
' rngNDMSectionProperties. Расчет с Geometry.Source = AutoCAD использует ее,
' чтобы не принять generated-геометрию за импортированную.
Public Function ResultsGeometrySource(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLoadCase)), "ALL", vbTextCompare) = 0 And _
                StrComp(CStr(data(rowIndex, colParameter)), "Geometry.Source", vbTextCompare) = 0 Then
            ResultsGeometrySource = Trim$(CStr(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Находит профиль выбранного LC в последнем snapshot Results.
' Основной источник - rngNDMElementResults; свойства LC используются только
' для понятной ошибки, если выбранное состояние профиля не записывалось.
Private Function ReadProfileIdForLoadCase(ByVal workbook As Object, ByVal combinationID As String) As String
    ReadProfileIdForLoadCase = ReadProfileIdFromElementResults(workbook, combinationID)
    If Len(ReadProfileIdForLoadCase) = 0 Then
        ReadProfileIdForLoadCase = ReadLoadCasePropertyText(workbook, combinationID, "ProfileId")
    End If
End Function

' Ищет профиль LC в сохраненной поэлементной таблице. Пустой ответ означает,
' что профиль надо искать в свойствах LC; модель и расчет не создаются.
Private Function ReadProfileIdFromElementResults(ByVal workbook As Object, ByVal combinationID As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colProfileId As Long: colProfileId = ResultColumn(data, "ProfileId")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colLoadCase))), combinationID, vbTextCompare) = 0 Then
            ReadProfileIdFromElementResults = Trim$(SafeText(data(rowIndex, colProfileId)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Находит текстовое свойство выбранного LC в длинной таблице параметров.
' Не подставляет значение из текущего ввода при отсутствии свойства snapshot.
Private Function ReadLoadCasePropertyText(ByVal workbook As Object, ByVal combinationID As String, ByVal propertyName As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colLoadCase))), combinationID, vbTextCompare) = 0 And _
                StrComp(Trim$(SafeText(data(rowIndex, colParameter))), propertyName, vbTextCompare) = 0 Then
            ReadLoadCasePropertyText = Trim$(SafeText(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Отбирает Stress/Strain и PhysicalState элементов только для выбранных LC
' и named-state. Отсутствие рассчитанного состояния дает адресную ошибку,
' а не незаметную замену результатом другого сочетания.
Private Sub ReadElementResultsForCombination(ByVal workbook As Object, ByVal resultType As String, _
        ByRef combinationID As String, ByVal stateType As String, _
        ByVal resultByID As Object, ByVal physicalStateByID As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4352, "ReadElementResultsForCombination", _
        "На листе Results нет таблицы результатов НДМ. Сначала выполните расчет."

    Dim colCombination As Long: colCombination = ResultColumn(data, "LoadCase")
    Dim colStateType As Long: colStateType = ResultColumn(data, "StateType")
    Dim colID As Long: colID = ResultColumn(data, "ElementID")
    Dim colState As Long: colState = ResultColumn(data, "PhysicalState")
    Dim colValue As Long
    Select Case LCase$(Trim$(resultType))
        Case "stress"
            colValue = ResultColumn(data, "Stress")
        Case "strain"
            colValue = ResultColumn(data, "Strain")
        Case Else
            Err.Raise vbObjectError + 4359, "ReadElementResultsForCombination", _
                "Visualization.Quantity профиля должен быть Stress или Strain."
    End Select

    If Len(combinationID) = 0 Then Err.Raise vbObjectError + 4353, "ReadElementResultsForCombination", _
        "В Results нет рассчитанных сочетаний для экспорта."

    Dim found As Boolean
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colCombination))), combinationID, vbTextCompare) = 0 And _
                StrComp(Trim$(SafeText(data(rowIndex, colStateType))), stateType, vbTextCompare) = 0 Then
            found = True
            resultByID(CStr(data(rowIndex, colID))) = CDbl(data(rowIndex, colValue))
            physicalStateByID(CStr(data(rowIndex, colID))) = CStr(data(rowIndex, colState))
        End If
    Next rowIndex

    If Not found Then Err.Raise vbObjectError + 4354, "ReadElementResultsForCombination", _
        MissingExportStateMessage(combinationID, stateType)
End Sub

' Читает сохраненную плоскость, точку нагрузки и выбранные главные оси.
' Переводит длины/кривизны из единиц snapshot и готовит предупреждение о
' несошедшемся либо вспомогательном НДС; равновесие повторно не ищется.
Private Sub ReadSectionPropertiesForCombination(ByVal workbook As Object, ByVal units As CUnitSystem, _
        ByVal combinationID As String, ByVal stateType As String, ByVal principalAxesMode As String, _
        ByRef epsilon0 As Double, ByRef kappaX As Double, _
        ByRef kappaY As Double, ByRef loadReferenceX As Double, ByRef loadReferenceY As Double, _
        ByRef centroidX As Double, ByRef centroidY As Double, ByRef principalAngle As Double, _
        ByRef extensionUsed As Boolean, ByRef stateWarningText As String)
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Err.Raise vbObjectError + 4360, "ReadSectionPropertiesForCombination", _
        "На листе Results нет rngNDMSectionProperties. Сначала выполните расчет."

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")

    Dim foundState As Boolean
    Dim stateStatus As String
    Dim policy As CResultStatusPolicy
    Dim rowIndex As Long
    Dim concreteCentroidX As Double
    Dim concreteCentroidY As Double
    Dim concretePrincipalAngle As Double
    Dim transformedCentroidX As Double
    Dim transformedCentroidY As Double
    Dim transformedPrincipalAngle As Double
    Set policy = New CResultStatusPolicy
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(CStr(data(rowIndex, colLoadCase)), "ALL", vbTextCompare) = 0 Then
            Select Case LCase$(Trim$(CStr(data(rowIndex, colParameter))))
                Case "concrete.centroidx"
                    concreteCentroidX = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "concrete.centroidy"
                    concreteCentroidY = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "concrete.principalangle"
                    concretePrincipalAngle = CDbl(data(rowIndex, colValue))
                Case "transformed.centroidx"
                    transformedCentroidX = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "transformed.centroidy"
                    transformedCentroidY = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "transformed.principalangle"
                    transformedPrincipalAngle = CDbl(data(rowIndex, colValue))
                Case "loadreferencex"
                    loadReferenceX = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "loadreferencey"
                    loadReferenceY = OutputLengthToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
            End Select
        ElseIf StrComp(CStr(data(rowIndex, colLoadCase)), combinationID, vbTextCompare) = 0 Then
            Dim key As String
            Dim prefix As String
            key = LCase$(Trim$(SafeText(data(rowIndex, colParameter))))
            prefix = "state." & LCase$(stateType) & "."
            If Left$(key, Len(prefix)) = prefix Then
                Select Case Mid$(key, Len(prefix) + 1)
                Case "epsilon0"
                    epsilon0 = CDbl(data(rowIndex, colValue))
                    foundState = True
                Case "kappax"
                    kappaX = OutputCurvatureToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "kappay"
                    kappaY = OutputCurvatureToInternalByUnit(CDbl(data(rowIndex, colValue)), CStr(data(rowIndex, colUnit)))
                Case "extensionused"
                    extensionUsed = SafeBoolean(data(rowIndex, colValue))
                Case "status"
                    stateStatus = Trim$(SafeText(data(rowIndex, colValue)))
                End Select
            End If
        End If
    Next rowIndex

    If extensionUsed Or policy.ToUserStatus(stateStatus) = policy.Fail Then
        stateWarningText = EXTENSION_WARNING_TEXT
    ElseIf policy.ToUserStatus(stateStatus) = policy.NumFail Then
        stateWarningText = NUMERICAL_STATE_WARNING_TEXT
    End If

    If Not foundState Then Err.Raise vbObjectError + 4361, "ReadSectionPropertiesForCombination", _
        MissingExportStateMessage(combinationID, stateType)

    Select Case principalAxesMode
        Case "Concrete"
            centroidX = concreteCentroidX
            centroidY = concreteCentroidY
            principalAngle = concretePrincipalAngle
        Case Else
            centroidX = transformedCentroidX
            centroidY = transformedCentroidY
            principalAngle = transformedPrincipalAngle
    End Select
End Sub

' Объясняет отсутствие выбранного state, различая выключенный расчет трещин
' и отсутствие пороговой точки. Это подсказка экспорта, не назначение статуса.
Private Function MissingExportStateMessage(ByVal combinationID As String, ByVal stateType As String) As String
    MissingExportStateMessage = "Запрашиваемое состояние """ & stateType & _
        """ не найдено в Results для AutoCAD export, сочетание " & combinationID & "."
    If IsCrackExportState(stateType) Then
        If Not ExportCombinationHasCrackWidth(combinationID) Then
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для запрашиваемого сочетания расчет трещин не выполнялся: этому сочетанию назначен профиль, где Calculation.Crack.Width = No."
        ElseIf StrComp(stateType, "CrackedState", vbTextCompare) = 0 Then
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, поэтому CrackedState должен быть в snapshot. Проверьте статус расчета трещин и сообщения о сходимости CrackedState."
        Else
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, но пороговое состояние образования трещины не записано. Это штатно только когда нет действия для поиска нормальной трещины, например чистое сжатие без момента. Если растяжение или изгиб есть, проверьте статус проверки образования трещины."
        End If
    End If
End Function

' Читает численное поле готовой геометрии. Только разрешенная пустота означает
' отсутствие оболочки/локальной инерции; обязательные координаты/A/диаметр,
' текстовый префикс, Boolean и ошибка формулы не превращаются в ноль/число.
' Диагностика указывает фактическую ячейку даже после переноса таблицы.
Private Function ReadGeometryNumber(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, Optional ByVal allowBlank As Boolean = False) As Double
    Dim value As Variant
    value = data(row, column)
    If IsError(value) Or IsNull(value) Then GoTo InvalidNumber
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then GoTo InvalidNumber
    If Len(Trim$(CStr(value))) = 0 Then
        If allowBlank Then Exit Function
        GoTo InvalidNumber
    End If
    If Not IsNumeric(value) Then GoTo InvalidNumber
    On Error GoTo InvalidNumber
    ReadGeometryNumber = CDbl(value)
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4367, "ReadSectionGeometryFromResults", _
        "В сохраненной геометрии Results поле " & ResultHeaderBase(CStr(data(1, column))) & _
        " содержит нечисловое значение: " & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Повторите импорт геометрии или расчет, чтобы восстановить корректный снимок."
End Function

' Определяет принадлежность named-state к трещинам для сообщения экспорта.
' Неизвестное имя не получает ложный crack-specific совет.
Private Function IsCrackExportState(ByVal stateType As String) As Boolean
    On Error GoTo NotCrackState
    Select Case SectionStateTypeFromText(stateType)
        Case sstPreCrackState, sstPostCrackState, sstCrackedState
            IsCrackExportState = True
    End Select
NotCrackState:
End Function

' Проверяет, запрошены ли трещины профилем LC, указанным в snapshot.
' Используется только для объяснения отсутствующей визуализации.
Private Function ExportCombinationHasCrackWidth(ByVal combinationID As String) As Boolean
    On Error GoTo Failed
    Dim workbook As Object
    Set workbook = ThisWorkbook
    Dim profileId As String
    profileId = ReadProfileIdForLoadCase(workbook, combinationID)
    If Len(profileId) = 0 Then Exit Function

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook
    ExportCombinationHasCrackWidth = profiles.ProfileById(profileId).CrackWidthEnabled
Failed:
End Function

' Читает выбранное определяющее сочетание из готовой batch-сводки.
' При отсутствии метаданных вызывающий слой сам выбирает сохраненный fallback LC.
Private Function ReadGoverningCombinationID(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    ReadGoverningCombinationID = Trim$(SafeText(anchor.Value2))
Failed:
End Function

' Безопасно читает текст из ячейки Results.
Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    SafeText = CStr(value)
End Function

' Распознает сохраненный логический признак snapshot, включая текст Excel.
' Ошибка ячейки не трактуется как доказанное использование расширения.
Private Function SafeBoolean(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        SafeBoolean = CBool(value)
        Exit Function
    End If
    Select Case LCase$(Trim$(CStr(value)))
        Case "true", "yes", "да", "1"
            SafeBoolean = True
    End Select
End Function

' Разрешает Worst через сохраненную сводку/таблицы либо принимает точный LC ID.
' Пустой выбор и отсутствие рассчитанных сочетаний отклоняются явно.
Private Function ResolveExportCombinationID(ByVal workbook As Object, ByVal settingValue As String) As String
    Dim valueText As String
    valueText = Trim$(settingValue)
    If Len(valueText) = 0 Then Err.Raise vbObjectError + 4357, "ResolveExportCombinationID", _
        "AutoCAD.Export.CombinationID должен быть Worst или точным CombinationID из Results."

    If StrComp(valueText, "Worst", vbTextCompare) = 0 Then
        ResolveExportCombinationID = ReadGoverningCombinationID(workbook)
        If Len(ResolveExportCombinationID) = 0 Then ResolveExportCombinationID = FirstCalculatedLoadCaseFromResults(workbook)
        If Len(ResolveExportCombinationID) = 0 Then ResolveExportCombinationID = FirstLoadCaseFromSectionProperties(workbook)
        If Len(ResolveExportCombinationID) = 0 Then Err.Raise vbObjectError + 4358, "ResolveExportCombinationID", _
            "В Results не найдено рассчитанное сочетание для AutoCAD.Export.CombinationID = Worst."
    Else
        ResolveExportCombinationID = valueText
    End If
End Function

' Возвращает первый LC из свойств Results, если batch summary не содержит
' определяющего сочетания и поэлементная таблица пуста. Дальше экспорт все
' равно проверит, что выбранное Visualization.State реально есть в snapshot.
Private Function FirstLoadCaseFromSectionProperties(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long
    colLoadCase = ResultColumn(data, "LoadCase")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        FirstLoadCaseFromSectionProperties = Trim$(SafeText(data(rowIndex, colLoadCase)))
        If Len(FirstLoadCaseFromSectionProperties) > 0 And _
                StrComp(FirstLoadCaseFromSectionProperties, "ALL", vbTextCompare) <> 0 Then Exit Function
    Next rowIndex
Failed:
End Function

' Возвращает первое сочетание, для которого в Results есть LC-зависимые
' результаты элементов. Это тот же fallback, которым пользуется Excel-схема,
' когда определяющее сочетание в batch summary не заполнено.
Private Function FirstCalculatedLoadCaseFromResults(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long
    colLoadCase = ResultColumn(data, "LoadCase")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        FirstCalculatedLoadCaseFromResults = Trim$(SafeText(data(rowIndex, colLoadCase)))
        If Len(FirstCalculatedLoadCaseFromResults) > 0 Then Exit Function
    Next rowIndex
Failed:
End Function

' Проверяет, есть ли в Variant-таблице заголовок и строка данных. Локальный
' перехват относится только к отсутствующему/непрямоугольному массиву.
Private Function HasResultTableRows(ByVal data As Variant) As Boolean
    On Error GoTo Failed
    HasResultTableRows = (UBound(data, 1) >= 2 And UBound(data, 2) >= 1)
    Exit Function
Failed:
    HasResultTableRows = False
End Function

' Находит обязательную колонку по смысловому имени без суффикса единиц.
' Отсутствующая колонка дает ошибку структуры Results до индексации данных.
Private Function ResultColumn(ByRef data As Variant, ByVal headerName As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 4355, "ResultColumn", "В Results не найден столбец: " & headerName
End Function

' Геометрический snapshot явно хранит статус интерпретации оболочки.
' Допускает ShapeType в сохраненной таблице без GeometryInterpretationStatus.
Private Function GeometryStatusColumn(ByRef data As Variant) As Long
    Dim column As Long, header As String
    For column = 1 To UBound(data, 2)
        header = ResultHeaderBase(CStr(data(1, column)))
        If StrComp(header, "GeometryInterpretationStatus", vbTextCompare) = 0 Or _
                StrComp(header, "ShapeType", vbTextCompare) = 0 Then
            GeometryStatusColumn = column
            Exit Function
        End If
    Next column
    Err.Raise vbObjectError + 4355, "ReadSectionGeometryFromResults", _
        "В Results отсутствует столбец GeometryInterpretationStatus или ShapeType. " & _
        "Повторите импорт геометрии или расчет, чтобы восстановить таблицу."
End Function

' Отделяет имя колонки от единиц после запятой, чтобы смена output-единиц
' не меняла поиск обязательного поля в сохраненной таблице.
Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

' Читает таблицу Results от именованного якоря. Метод не использует
' CurrentRegion, чтобы человекочитаемые заголовки над таблицами не попадали
' в массив расчетных данных.
Private Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    Dim useReadTables As Boolean
    useReadTables = Not mSnapshotReadTables Is Nothing And workbook Is mSnapshotReadWorkbook
    If useReadTables Then
        If mSnapshotReadTables.Exists(rangeName) Then
            ReadAnchoredResultTable = mSnapshotReadTables.Item(rangeName)
            Exit Function
        End If
    End If
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = AnchoredColumnCount(anchor)
    If columnCount <= 0 Then Exit Function

    Dim rowCount As Long
    rowCount = AnchoredRowCount(anchor)
    If rowCount <= 0 Then rowCount = 1

    ReadAnchoredResultTable = anchor.Resize(rowCount, columnCount).Value2
    If useReadTables Then
        mSnapshotReadTables.Add rangeName, ReadAnchoredResultTable
        mSnapshotTableReadCount = mSnapshotTableReadCount + 1
    End If
End Function

' Определяет ширину таблицы по непрерывной строке заголовков.
Private Function AnchoredColumnCount(ByVal anchor As Object) As Long
    AnchoredColumnCount = SnapshotAnchoredExtent(anchor, True, 256)
End Function

' Определяет высоту таблицы по первому столбцу, где все Results-таблицы имеют
' обязательный RunID/ElementID/AnnotationID в каждой строке данных.
Private Function AnchoredRowCount(ByVal anchor As Object) As Long
    AnchoredRowCount = SnapshotAnchoredExtent(anchor, False)
End Function

' Достает единицу измерения из заголовка вида "X, mm".
Private Function ResultHeaderUnit(ByRef data As Variant, ByVal colIndex As Long, ByVal defaultUnit As String) As String
    Dim headerText As String
    headerText = CStr(data(1, colIndex))

    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos <= 0 Then
        ResultHeaderUnit = defaultUnit
        Exit Function
    End If

    Dim unitText As String
    unitText = Trim$(Mid$(headerText, commaPos + 1))
    Dim bracketPos As Long
    bracketPos = InStr(1, unitText, "(", vbTextCompare)
    If bracketPos > 0 Then unitText = Trim$(Left$(unitText, bracketPos - 1))

    If Len(unitText) = 0 Then unitText = defaultUnit
    ResultHeaderUnit = unitText
End Function

' Переводит длину из единицы, сохраненной в Results snapshot, во внутренние мм.
Private Function OutputLengthToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputLengthToInternalByUnit = mSnapshotUnits.OutputLengthToInternal(value, unitText)
End Function

' Переводит площадь из единицы, сохраненной в Results snapshot, во внутренние мм2.
Private Function OutputAreaToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputAreaToInternalByUnit = mSnapshotUnits.OutputAreaToInternal(value, unitText)
End Function

' Переводит собственные моменты инерции элемента из snapshot в мм4.
Private Function OutputFourthPowerLengthToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    Dim baseUnit As String
    baseUnit = Trim$(unitText)
    If Right$(baseUnit, 1) = "4" Then baseUnit = Left$(baseUnit, Len(baseUnit) - 1)
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputFourthPowerLengthToInternalByUnit = mSnapshotUnits.OutputFourthPowerLengthToInternal(value, baseUnit)
End Function

' Переводит кривизну из единицы, сохраненной в Results snapshot, во внутренние 1/мм.
Private Function OutputCurvatureToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    If Len(Trim$(unitText)) = 0 Then Err.Raise vbObjectError + 4364, "OutputCurvatureToInternalByUnit", _
        "В Results не указана единица кривизны. Повторите расчет, чтобы восстановить снимок."
    If mSnapshotUnits Is Nothing Then Set mSnapshotUnits = New CUnitSystem
    OutputCurvatureToInternalByUnit = value / mSnapshotUnits.InternalCurvatureToOutput(1#, unitText)
End Function

' Строит геометрию, подписи, оси и нулевую линию по готовому Results snapshot
' в активном документе AutoCAD. Использует сохраненные размеры элементов;
' экспорт не читает заново исходное сечение и не запускает решатель НДС.
Private Sub DrawResultsStressExport(ByVal section As CSectionModel, _
        ByVal resultByID As Object, ByVal physicalStateByID As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal centroidX As Double, ByVal centroidY As Double, ByVal principalAngle As Double, _
        ByVal resultPrecision As Long, ByVal stateWarningText As String, ByRef exportSettings As TAutoCADExportSettings, _
        ByRef contourExportCount As Long, Optional ByVal targetDocument As Object = Nothing, _
        Optional ByVal crackRows As Object = Nothing, Optional ByVal crackSweeps As Object = Nothing)
    ' Проверяем все углы до подключения и добавления объектов в чертеж.
    Dim arcSweeps As Object
    If exportSettings.ContourEnabled Then Set arcSweeps = ReadSavedContourArcSweeps(ThisWorkbook)
    Dim doc As Object
    Set doc = targetDocument
    If doc Is Nothing Then
        Dim acad As Object
        Set acad = ConnectToRunningAutoCAD()
        Set doc = ActiveAutoCADDocument(acad)
    End If

    Dim ms As Object
    Set ms = doc.ModelSpace

    EnsureAcadLayer doc, exportSettings.ConcreteLayer, 8
    EnsureAcadLayer doc, exportSettings.RebarLayer, 1
    EnsureAcadLayer doc, exportSettings.ConcreteTensionLayer, exportSettings.ConcreteTensionColor
    EnsureAcadLayer doc, exportSettings.ConcreteCompressionLayer, exportSettings.ConcreteCompressionColor
    EnsureAcadLayer doc, exportSettings.RebarTensionLayer, exportSettings.RebarTensionColor
    EnsureAcadLayer doc, exportSettings.RebarCompressionLayer, exportSettings.RebarCompressionColor
    EnsureAcadLayer doc, "RC_NDM_Axes", 3
    EnsureAcadLayer doc, "RC_NDM_LoadPoint", 2
    EnsureAcadLayer doc, "RC_NDM_NeutralLine", exportSettings.NeutralColor
    If exportSettings.ContourEnabled Then EnsureAcadLayer doc, exportSettings.ContourLayer, SECTION_CONTOUR_COLOR_INDEX
    If Len(exportSettings.OpeningContourLayer) > 0 Then EnsureAcadLayer doc, exportSettings.OpeningContourLayer, OPENING_CONTOUR_COLOR_INDEX
    If Not crackRows Is Nothing Then
        If crackRows.Count > 0 Then EnsureAcadLayer doc, exportSettings.CrackInteractionLayer, CRACK_REGION_COLOR_INDEX
    End If
    EnsureAcadLayer doc, EXTENSION_WARNING_LAYER, 1

    Dim i As Long
    Dim resultValue As Double
    Dim physicalState As String
    Dim textHeight As Double

    For i = 1 To section.ConcreteCount
        Dim concreteWidth As Double
        Dim concreteHeight As Double
        Dim concreteRotation As Double
        concreteWidth = ConcreteDrawWidth(section, i)
        concreteHeight = ConcreteDrawHeight(section, i)
        concreteRotation = ConcreteDrawRotation(section, i)
        resultValue = LookupResultValue(resultByID, section.ConcreteID(i))
        physicalState = ResultPhysicalState(physicalStateByID, section.ConcreteID(i))
        textHeight = 0.22 * MinDouble(concreteWidth, concreteHeight)
        If textHeight <= 0# Then textHeight = 1#
        AddAcadRectangleRegion ms, section.ConcreteX(i), section.ConcreteY(i), concreteWidth, concreteHeight, concreteRotation, _
            exportSettings.ConcreteLayer, ResultColorByPhysicalState("Concrete", physicalState, exportSettings)

        Dim labelX As Double
        Dim labelY As Double
        ConcreteLabelPoint section.ConcreteX(i), section.ConcreteY(i), concreteWidth, concreteHeight, concreteRotation, _
            labelX, labelY
        AddAcadText ms, ResultLabelText(section.ConcreteID(i), resultValue, exportSettings.IncludeElementNames, resultPrecision), _
            labelX, labelY, textHeight, _
            ResultAnnotationLayerByPhysicalState("Concrete", physicalState, exportSettings), _
            ResultColorByPhysicalState("Concrete", physicalState, exportSettings)
    Next i

    For i = 1 To section.RebarCount
        resultValue = LookupResultValue(resultByID, section.RebarID(i))
        physicalState = ResultPhysicalState(physicalStateByID, section.RebarID(i))
        AddAcadCircleRegion ms, section.RebarX(i), section.RebarY(i), section.RebarDiameter(i) / 2#, _
            exportSettings.RebarLayer, ResultColorByPhysicalState("Rebar", physicalState, exportSettings)
        AddAcadText ms, ResultLabelText(section.RebarID(i), resultValue, exportSettings.IncludeElementNames, resultPrecision), _
            section.RebarX(i) + section.RebarDiameter(i) / 2#, section.RebarY(i) + section.RebarDiameter(i) / 2#, _
            MaxDouble(2.5, section.RebarDiameter(i) * 0.18), _
            ResultAnnotationLayerByPhysicalState("Rebar", physicalState, exportSettings), _
            ResultColorByPhysicalState("Rebar", physicalState, exportSettings)
    Next i

    ' Контур выводим после бетонных и арматурных объектов, чтобы он не
    ' оказался закрыт AutoCAD Region, созданными для волокон расчетной сетки.
    If exportSettings.ContourEnabled Then
        contourExportCount = DrawSavedContourRows(ThisWorkbook, ms, exportSettings.ContourLayer, arcSweeps, , "CONTOUR_", exportSettings.OpeningContourLayer)
    End If
    If Not crackRows Is Nothing Then
        If crackRows.Count > 0 Then
            contourExportCount = contourExportCount + DrawSavedContourRows(ThisWorkbook, ms, _
                exportSettings.CrackInteractionLayer, crackSweeps, crackRows, "CRACK_REGION_")
        End If
    End If

    DrawCentroidAxesAndLoadPoint ms, section, centroidX, centroidY, principalAngle, _
        loadReferenceX, loadReferenceY, _
        exportSettings.PrincipalAxesMode, exportSettings.LoadPointEnabled
    If exportSettings.NeutralLineEnabled Then
        DrawNeutralLineByState ms, section, epsilon0, kappaX, kappaY, exportSettings.NeutralColor
    End If
    If Len(stateWarningText) > 0 Then DrawStateWarning ms, section, stateWarningText

    doc.Regen 1
End Sub

' Выбирает целиком актуальную область нужного сочетания. Старый или смешанный
' snapshot пропускается целиком, чтобы не выгрузить только часть его замкнутых контуров.
' Проверяются metadata, а не внешние статусы: область выполненной FAIL-проверки
' ширины пригодна для диагностического вывода. Новое НДС здесь не решается.
Private Function CurrentCrackRegionRows(ByVal workbook As Object, ByVal combinationID As String) As Object
    Dim selected As Object, data As Variant, geometry As Variant, properties As Variant
    Set selected = CreateObject("Scripting.Dictionary")
    Set CurrentCrackRegionRows = selected
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    geometry = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    properties = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Or Not HasResultTableRows(geometry) Or Not HasResultTableRows(properties) Then Exit Function
    Dim colOwner As Long, colRevision As Long, colType As Long, colRun As Long
    ' В старых снимках нет result-region metadata; это отсутствие области,
    ' а не причина пытаться восстановить ее по численной площади.
    On Error Resume Next
    colOwner = ResultColumn(data, "CombinationID"): colRevision = ResultColumn(data, "SectionRevision")
    On Error GoTo 0
    If colOwner = 0 Or colRevision = 0 Then Exit Function
    colType = ResultColumn(data, "AnnotationType"): colRun = ResultColumn(data, "RunID")
    Dim colParameter As Long, colValue As Long, colCase As Long, revision As Variant, runID As Variant, row As Long
    colParameter = ResultColumn(properties, "Parameter"): colValue = ResultColumn(properties, "Value")
    colCase = ResultColumn(properties, "LoadCase")
    runID = geometry(2, ResultColumn(geometry, "RunID"))
    For row = 2 To UBound(properties, 1)
        If SafeText(properties(row, colCase)) = "ALL" And SafeText(properties(row, colParameter)) = "Geometry.Revision" Then
            revision = properties(row, colValue)
            If CStr(properties(row, ResultColumn(properties, "RunID"))) <> CStr(runID) Then Exit Function
            Exit For
        End If
    Next row
    If IsEmpty(revision) Then Exit Function
    Dim anchor As Object, colLoop As Long, colStandard As Long, colAnchor As Long, identity As String, rowIdentity As String
    Set anchor = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
    colLoop = ResultColumn(data, "LoopID"): colStandard = ResultColumn(data, "Standard"): colAnchor = ResultColumn(data, "AnchorID")
    For row = 2 To UBound(data, 1)
        If Left$(UCase$(SafeText(data(row, colType))), 13) = "CRACK_REGION_" And _
                StrComp(SafeText(data(row, colOwner)), combinationID, vbTextCompare) = 0 Then
            If CStr(data(row, colRun)) <> CStr(runID) Or CStr(data(row, colRevision)) <> CStr(revision) Then
                selected.RemoveAll: Exit Function
            End If
            rowIdentity = ReadSavedAnnotationText(data, anchor, row, colStandard) & "|" & ReadSavedAnnotationText(data, anchor, row, colAnchor)
            If Len(identity) = 0 Then identity = rowIdentity
            If rowIdentity <> identity Then selected.RemoveAll: Exit Function
            If ReadGeometryNumber(data, anchor, row, colLoop) < 1# Then Err.Raise vbObjectError + 4376, _
                "CurrentCrackRegionRows", "Номер замкнутого контура расчетной области должен быть положительным: " & _
                anchor.Parent.Name & "!" & anchor.Offset(row - 1, colLoop - 1).Address(False, False) & ". Повторите расчет."
            selected.Add CStr(row), True
        End If
    Next row
End Function

' Рисует сохраненные сегменты material-contour или result-region одним
' polyline-механизмом. Не восстанавливает область по сетке или ее площади.
' Каждое opening остается отдельной замкнутой петлей с аналитическими дугами.
Private Function DrawSavedContourRows(ByVal workbook As Object, ByVal ms As Object, ByVal contourLayer As String, _
        ByVal arcSweeps As Object, Optional ByVal selectedRows As Object = Nothing, _
        Optional ByVal typePrefix As String = "CONTOUR_", Optional ByVal openingLayer As String = vbNullString) As Long
    If typePrefix = "CONTOUR_" Then
        DrawSavedContourRows = DrawSavedMaterialContours(workbook, ms, contourLayer, openingLayer)
        Exit Function
    End If
    Dim tableName As Object
    On Error Resume Next
    Set tableName = workbook.Names.Item("rngNDMSectionAnnotations")
    On Error GoTo 0
    If tableName Is Nothing Then Exit Function
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colType As Long: colType = ResultColumn(data, "AnnotationType")
    Dim colID As Long: colID = ResultColumn(data, "AnnotationID")
    Dim colStartX As Long: colStartX = ResultColumn(data, "StartX")
    Dim colStartY As Long: colStartY = ResultColumn(data, "StartY")
    Dim colEndX As Long: colEndX = ResultColumn(data, "EndX")
    Dim colEndY As Long: colEndY = ResultColumn(data, "EndY")
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")
    Dim defaultLengthUnit As String
    defaultLengthUnit = ResultsOutputLengthUnit(workbook)
    If Len(defaultLengthUnit) = 0 Then defaultLengthUnit = ResultHeaderUnit(data, colStartX, "mm")

    Dim segStartX() As Double
    Dim segStartY() As Double
    Dim segEndX() As Double
    Dim segEndY() As Double
    Dim segBulge() As Double
    Dim segCount As Long
    Dim currentLoopKey As String
    Dim currentLayer As String, rowLayer As String
    Dim colLoop As Long
    If typePrefix = "CRACK_REGION_" Then colLoop = ResultColumn(data, "LoopID")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Not selectedRows Is Nothing Then
            If Not selectedRows.Exists(CStr(rowIndex)) Then GoTo NextContourRow
        End If
        Dim annotationType As String
        annotationType = UCase$(Trim$(SafeText(data(rowIndex, colType))))
        If Left$(annotationType, Len(typePrefix)) = typePrefix Then
            Dim loopKey As String
            loopKey = ContourLoopKey(SafeText(data(rowIndex, colID)))
            If colLoop > 0 Then loopKey = CStr(data(rowIndex, colLoop))
            rowLayer = contourLayer
            If Len(openingLayer) > 0 And InStr(1, UCase$(SafeText(data(rowIndex, colID))), "_OPENING_", vbBinaryCompare) > 0 Then rowLayer = openingLayer
            If segCount > 0 And StrComp(loopKey, currentLoopKey, vbTextCompare) <> 0 Then
                DrawSavedContourRows = DrawSavedContourRows + _
                    DrawContourSegmentPolyline(ms, currentLayer, segStartX, segStartY, segEndX, segEndY, segBulge, segCount, CRACK_REGION_COLOR_INDEX)
                Erase segStartX
                Erase segStartY
                Erase segEndX
                Erase segEndY
                Erase segBulge
                segCount = 0
            End If
            currentLoopKey = loopKey
            currentLayer = rowLayer
        End If

        Select Case annotationType
            Case typePrefix & "LINE"
                Dim lineUnit As String
                lineUnit = AnnotationLengthUnit(data, rowIndex, colUnit, defaultLengthUnit)
                AppendContourSegment segStartX, segStartY, segEndX, segEndY, segBulge, segCount, _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartX)), lineUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartY)), lineUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colEndX)), lineUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colEndY)), lineUnit), _
                    0#
            Case typePrefix & "ARC"
                Dim arcUnit As String
                arcUnit = AnnotationLengthUnit(data, rowIndex, colUnit, defaultLengthUnit)
                AppendContourSegment segStartX, segStartY, segEndX, segEndY, segBulge, segCount, _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartX)), arcUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartY)), arcUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colEndX)), arcUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colEndY)), arcUnit), _
                    Tan(CDbl(arcSweeps.Item(CStr(rowIndex))) / 4#)
            Case typePrefix & "CIRCLE"
                If segCount > 0 Then
                    DrawSavedContourRows = DrawSavedContourRows + _
                        DrawContourSegmentPolyline(ms, currentLayer, segStartX, segStartY, segEndX, segEndY, segBulge, segCount, CRACK_REGION_COLOR_INDEX)
                    Erase segStartX
                    Erase segStartY
                    Erase segEndX
                    Erase segEndY
                    Erase segBulge
                    segCount = 0
                    currentLoopKey = vbNullString
                End If
                Dim circleUnit As String
                circleUnit = AnnotationLengthUnit(data, rowIndex, colUnit, defaultLengthUnit)
                DrawSavedContourRows = DrawSavedContourRows + _
                    DrawContourCirclePolyline(ms, rowLayer, _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartX)), circleUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colStartY)), circleUnit), _
                    OutputLengthToInternalByUnit(CDbl(data(rowIndex, colEndX)), circleUnit), CRACK_REGION_COLOR_INDEX)
        End Select
NextContourRow:
    Next rowIndex

    If segCount > 0 Then
        DrawSavedContourRows = DrawSavedContourRows + _
            DrawContourSegmentPolyline(ms, currentLayer, segStartX, segStartY, segEndX, segEndY, segBulge, segCount, CRACK_REGION_COLOR_INDEX)
    End If
End Function

' Выгружает исходные замкнутые контуры из геометрии Results теми же низкоуровневыми
' polyline-примитивами, что расчетные области. Роль отверстия и группировка
' сегментов берутся из metadata геометрии; аннотации здесь не читаются.
' Экспорт всегда добавляет новый вывод: не ищет совпадения и не очищает DWG.
Private Function DrawSavedMaterialContours(ByVal workbook As Object, ByVal ms As Object, _
        ByVal outerLayer As String, ByVal openingLayer As String) As Long
    Dim contours As CSectionContours, loops As Object, indices As Collection
    Dim i As Long, key As String, loopKey As Variant, indexValue As Variant, layer As String, colorIndex As Long
    Set contours = ReadSavedSectionContours(workbook)
    Set loops = CreateObject("Scripting.Dictionary")
    For i = 1 To contours.Count
        key = UCase$(contours.LoopRole(i)) & ":" & contours.LoopID(i)
        If Not loops.Exists(key) Then
            Set indices = New Collection
            loops.Add key, indices
        End If
        loops.Item(key).Add i
    Next i
    Dim sx() As Double, sy() As Double, ex() As Double, ey() As Double, bulge() As Double, count As Long
    For Each loopKey In loops.Keys
        Set indices = loops.Item(loopKey)
        i = CLng(indices(1))
        layer = outerLayer: colorIndex = SECTION_CONTOUR_COLOR_INDEX
        If StrComp(contours.LoopRole(i), "Opening", vbTextCompare) = 0 Then
            If Len(openingLayer) > 0 Then layer = openingLayer
            colorIndex = OPENING_CONTOUR_COLOR_INDEX
        End If
        Erase sx: Erase sy: Erase ex: Erase ey: Erase bulge: count = 0
        For Each indexValue In indices
            i = CLng(indexValue)
            If contours.SegmentType(i) = "CONTOUR_CIRCLE" Then
                If indices.Count <> 1 Then Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", "Круговой контур не должен содержать дополнительные сегменты."
                DrawSavedMaterialContours = DrawSavedMaterialContours + _
                    DrawContourCirclePolyline(ms, layer, contours.StartX(i), contours.StartY(i), contours.Radius(i), colorIndex, True)
            Else
                AppendContourSegment sx, sy, ex, ey, bulge, count, contours.StartX(i), contours.StartY(i), _
                    contours.EndX(i), contours.EndY(i), Tan(contours.SweepAngle(i) / 4#)
            End If
        Next indexValue
        If count > 0 Then DrawSavedMaterialContours = DrawSavedMaterialContours + _
            DrawContourSegmentPolyline(ms, layer, sx, sy, ex, ey, bulge, count, colorIndex, True)
    Next loopKey
End Function

' Подготавливает численные углы дуг одним проходом до записи DWG.
' Отсутствие необязательных аннотаций допустимо; ошибка Text не становится
' нулевым bulge. Ключом служит строка прочитанной таблицы, а не AutoCAD handle.
Private Function ReadSavedContourArcSweeps(ByVal workbook As Object, Optional ByVal selectedRows As Object = Nothing, _
        Optional ByVal typePrefix As String = "CONTOUR_") As Object
    Dim result As Object, tableName As Object, anchor As Object, data As Variant
    Set result = CreateObject("Scripting.Dictionary")
    Set ReadSavedContourArcSweeps = result
    If typePrefix = "CONTOUR_" Then
        Dim contours As CSectionContours, i As Long
        Set contours = ReadSavedSectionContours(workbook)
        For i = 1 To contours.Count
            If contours.SegmentType(i) = "CONTOUR_ARC" Then result.Add CStr(i), contours.SweepAngle(i)
        Next i
        Exit Function
    End If
    On Error Resume Next
    Set tableName = workbook.Names.Item("rngNDMSectionAnnotations")
    On Error GoTo 0
    If tableName Is Nothing Then Exit Function
    Set anchor = tableName.RefersToRange
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Function
    Dim colType As Long, colText As Long, rowIndex As Long
    colType = ResultColumn(data, "AnnotationType"): colText = ResultColumn(data, "Text")
    For rowIndex = 2 To UBound(data, 1)
        If Not selectedRows Is Nothing Then
            If Not selectedRows.Exists(CStr(rowIndex)) Then GoTo NextSweepRow
        End If
        If StrComp(Trim$(SafeText(data(rowIndex, colType))), typePrefix & "ARC", vbTextCompare) = 0 Then
            result.Add CStr(rowIndex), ReadContourArcSweep(data(rowIndex, colText), anchor.Parent.Name & "!" & _
                anchor.Offset(rowIndex - 1, colText - 1).Address(False, False))
        End If
NextSweepRow:
    Next rowIndex
End Function

' Читает единицу длины последнего расчетного снимка. Для annotation-таблицы это
' важнее заголовков StartX/EndX: сами заголовки не содержат единицу длины, но
' значения уже переведены writer-ом в Output.LengthUnit.
Private Function ResultsOutputLengthUnit(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colParameter))), "Output.LengthUnit", vbTextCompare) = 0 Then
            ResultsOutputLengthUnit = Trim$(SafeText(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Возвращает единицу длины для одной строки semantic-аннотации.
' В таблице rngNDMSectionAnnotations координатные заголовки не содержат ", mm":
' координаты уже сохранены в пользовательских output-единицах. В новых снимках
' contour-строки явно несут Unit, а для старых снимков без Unit берем общий
' Output.LengthUnit из свойств сечения. Иначе при Output.LengthUnit = m контур
' получается в 1000 раз меньше бетонной и арматурной геометрии.
Private Function AnnotationLengthUnit(ByRef data As Variant, ByVal rowIndex As Long, _
        ByVal colUnit As Long, ByVal defaultUnit As String) As String
    Dim unitText As String
    unitText = Trim$(SafeText(data(rowIndex, colUnit)))
    If Len(unitText) = 0 Or unitText = "-" Then unitText = defaultUnit
    If Len(unitText) = 0 Or unitText = "-" Then unitText = "mm"
    AnnotationLengthUnit = unitText
End Function

' Возвращает имя петли contour-аннотаций. Старые ID вида CONTOUR_LINE_1
' попадают в одну пустую петлю; CONTOUR_OUTER_* и CONTOUR_OPENING_* сохраняют
' прежние роли генератора. В CAD_<роль>_<source>_<loop>_<segment> убирается
' только номер сегмента: несколько outer/opening не склеиваются в одну петлю.
Private Function ContourLoopKey(ByVal annotationID As String) As String
    Dim parts() As String
    parts = Split(UCase$(Trim$(annotationID)), "_")
    If UBound(parts) >= 4 And parts(0) = "CAD" Then
        ContourLoopKey = Left$(UCase$(Trim$(annotationID)), InStrRev(annotationID, "_") - 1)
        Exit Function
    End If
    If UBound(parts) >= 3 Then
        If parts(0) = "CONTOUR" And (parts(1) = "OUTER" Or parts(1) = "OPENING") Then
            ContourLoopKey = parts(1)
        End If
    End If
End Function

' Накопляет линейный или дуговой сегмент будущей AutoCAD LWPOLYLINE.
' Bulge хранится на начальной вершине сегмента: 0 для прямого участка и
' Tan(sweep/4) для дуги AutoCAD.
Private Sub AppendContourSegment(ByRef startX() As Double, ByRef startY() As Double, _
        ByRef endX() As Double, ByRef endY() As Double, ByRef bulge() As Double, _
        ByRef segmentCount As Long, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal bulgeValue As Double)
    segmentCount = segmentCount + 1
    ReDim Preserve startX(1 To segmentCount)
    ReDim Preserve startY(1 To segmentCount)
    ReDim Preserve endX(1 To segmentCount)
    ReDim Preserve endY(1 To segmentCount)
    ReDim Preserve bulge(1 To segmentCount)

    startX(segmentCount) = x1
    startY(segmentCount) = y1
    endX(segmentCount) = x2
    endY(segmentCount) = y2
    bulge(segmentCount) = bulgeValue
End Sub

' Строит одну непрерывную LWPOLYLINE по порядку contour-аннотаций.
' Если сохраненный набор сегментов содержит разрыв, экспорт
' останавливается с понятной ошибкой: лучше увидеть проблему, чем получить в
' AutoCAD контур с паразитной перемычкой.
Private Function DrawContourSegmentPolyline(ByVal ms As Object, ByVal contourLayer As String, _
        ByRef startX() As Double, ByRef startY() As Double, _
        ByRef endX() As Double, ByRef endY() As Double, ByRef bulge() As Double, ByVal segmentCount As Long, _
        ByVal colorIndex As Long, Optional ByVal materialContour As Boolean = False) As Long
    Dim segmentIndex As Long
    For segmentIndex = 1 To segmentCount - 1
        If Not PointsAreClose(endX(segmentIndex), endY(segmentIndex), startX(segmentIndex + 1), startY(segmentIndex + 1)) Then
            Err.Raise vbObjectError + 4370, "DrawParametricSectionContour", _
                "Параметрический контур в Results имеет разрыв между соседними сегментами. Экспорт контура остановлен."
        End If
    Next segmentIndex
    If Not PointsAreClose(endX(segmentCount), endY(segmentCount), startX(1), startY(1)) Then
        Err.Raise vbObjectError + 4371, "DrawParametricSectionContour", _
            "Параметрический контур в Results не замкнут. Экспорт контура остановлен."
    End If

    Dim points() As Double
    ReDim points(0 To segmentCount * 2 - 1)
    For segmentIndex = 1 To segmentCount
        points((segmentIndex - 1) * 2) = startX(segmentIndex)
        points((segmentIndex - 1) * 2 + 1) = startY(segmentIndex)
    Next segmentIndex

    Dim entity As Object
    Set entity = AddAcadLightWeightPolyline(ms, points, contourLayer, colorIndex)
    For segmentIndex = 1 To segmentCount
        If Abs(bulge(segmentIndex)) > 0.000000000001 Then Call entity.SetBulge(segmentIndex - 1, bulge(segmentIndex))
    Next segmentIndex
    entity.Closed = True
    entity.Update
    If materialContour Then MarkNDMContourOutput entity
    DrawContourSegmentPolyline = 1
End Function

' Окружность тоже выводится LWPOLYLINE: четыре четверти с одинаковым bulge дают
' непрерывную замкнутую полилинию с дугами, а не отдельный объект Circle.
Private Function DrawContourCirclePolyline(ByVal ms As Object, ByVal contourLayer As String, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double, _
        ByVal colorIndex As Long, Optional ByVal materialContour As Boolean = False) As Long
    If radius <= 0# Then Exit Function

    Dim points(0 To 7) As Double
    points(0) = centerX + radius: points(1) = centerY
    points(2) = centerX: points(3) = centerY + radius
    points(4) = centerX - radius: points(5) = centerY
    points(6) = centerX: points(7) = centerY - radius

    Dim entity As Object
    Set entity = AddAcadLightWeightPolyline(ms, points, contourLayer, colorIndex)

    Dim quarterBulge As Double
    quarterBulge = Tan((GEOM_PI / 2#) / 4#)
    Dim i As Long
    For i = 0 To 3
        Call entity.SetBulge(i, quarterBulge)
    Next i
    entity.Closed = True
    entity.Update
    If materialContour Then MarkNDMContourOutput entity
    DrawContourCirclePolyline = 1
End Function

' Формирует короткий фрагмент итогового сообщения по экспорту контура.
' Ноль при включенной настройке означает отсутствие сохраненного контура.
Private Function ContourExportStatusText(ByVal contourEnabled As Boolean, ByVal contourCount As Long) As String
    If Not contourEnabled Then
        ContourExportStatusText = "выключено"
    ElseIf contourCount > 0 Then
        ContourExportStatusText = CStr(contourCount)
    Else
        ContourExportStatusText = "0 (контур не сохранен)"
    End If
End Function

' Создает служебную полилинию AutoCAD на отдельном слое контура.
Private Function AddAcadLightWeightPolyline(ByVal ms As Object, ByRef points() As Double, _
        ByVal layerName As String, ByVal colorIndex As Long) As Object
    Set AddAcadLightWeightPolyline = ms.AddLightWeightPolyline(points)
    AddAcadLightWeightPolyline.Layer = layerName
    AddAcadLightWeightPolyline.Color = colorIndex
End Function

' Сравнивает соседние вершины контура в миллиметрах.
Private Function PointsAreClose(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    PointsAreClose = (Abs(x1 - x2) <= CONTOUR_POINT_TOLERANCE And Abs(y1 - y2) <= CONTOUR_POINT_TOLERANCE)
End Function

' Добавляет в AutoCAD заметное предупреждение под сечением.
' Текст берется из сохраненного Results snapshot: это может быть выход за
' физическую диаграмму или численная несходимость прямого НДС.
Private Sub DrawStateWarning(ByVal ms As Object, ByVal section As CSectionModel, ByVal warningText As String)
    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim sectionSize As Double
    sectionSize = MaxDouble(maxX - minX, maxY - minY)
    If sectionSize <= 0# Then sectionSize = 100#

    AddAcadText ms, warningText, minX, minY - 0.14 * sectionSize, _
        MaxDouble(8#, 0.035 * sectionSize), EXTENSION_WARNING_LAYER, 1
End Sub

' Возвращает результат конкретного расчетного элемента из выбранного state.
' Отсутствующее значение является неполным snapshot, а не нулевым напряжением.
Private Function LookupResultValue(ByVal resultByID As Object, ByVal elementID As String) As Double
    If resultByID.Exists(elementID) Then
        LookupResultValue = CDbl(resultByID.Item(elementID))
    Else
        Err.Raise vbObjectError + 4356, "LookupResultValue", "В Results нет выбранного результата для элемента: " & elementID
    End If
End Function

' Возвращает физическое состояние элемента из snapshot Results.
' Цвет и слой AutoCAD выбираются по этому полю, а не по знаку Stress/Strain,
' потому что пользовательская SignConvention относится к N/Mx/My, а не к
' физическому знаку напряжений и деформаций элемента.
Private Function ResultPhysicalState(ByVal physicalStateByID As Object, ByVal elementID As String) As String
    If physicalStateByID.Exists(elementID) Then
        ResultPhysicalState = CStr(physicalStateByID.Item(elementID))
    Else
        Err.Raise vbObjectError + 4356, "ResultPhysicalState", "В Results нет PhysicalState для элемента: " & elementID
    End If
End Function

' Формирует подпись Stress/Strain для AutoCAD. Само значение уже прочитано из
' Results в пользовательской единице, поэтому здесь меняется только формат.
Private Function ResultLabelText(ByVal elementID As String, ByVal value As Double, _
        ByVal includeElementName As Boolean, ByVal precision As Long) As String
    Dim valueText As String
    valueText = FormatResultValue(value, precision)
    If includeElementName And Len(Trim$(elementID)) > 0 Then
        ResultLabelText = elementID & " " & valueText
    Else
        ResultLabelText = valueText
    End If
End Function

' Форматирует число с точностью визуализации профиля независимо от локали Excel.
Private Function FormatResultValue(ByVal value As Double, ByVal precision As Long) As String
    If precision < 0 Then precision = 0
    If precision > 10 Then precision = 10

    Dim pattern As String
    If precision = 0 Then
        pattern = "0"
    Else
        pattern = "0." & String$(precision, "0")
    End If
    FormatResultValue = Replace$(Format$(value, pattern), ",", ".")
End Function

' Рисует выбранные главные оси через сохраненный центр и маркер точки нагрузки.
' Длина линий зависит от габарита только для оформления; координаты/угол
' берутся из Results и не пересчитывают геометрические характеристики.
Private Sub DrawCentroidAxesAndLoadPoint(ByVal ms As Object, ByVal section As CSectionModel, _
        ByVal centroidX As Double, ByVal centroidY As Double, ByVal principalAngle As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal principalAxesMode As String, ByVal loadPointEnabled As Boolean)
    Dim principalAxesEnabled As Boolean
    principalAxesEnabled = PrincipalAxesModeDraws(principalAxesMode)
    If Not principalAxesEnabled And Not loadPointEnabled Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim axisLength As Double
    axisLength = 0.65 * MaxDouble(maxX - minX, maxY - minY)
    If axisLength <= 0# Then axisLength = 100#

    Dim a As Double
    a = principalAngle

    If principalAxesEnabled Then
        AddAcadLine ms, centroidX - axisLength * Cos(a), centroidY - axisLength * Sin(a), _
            centroidX + axisLength * Cos(a), centroidY + axisLength * Sin(a), "RC_NDM_Axes", 3
        AddAcadLine ms, centroidX - axisLength * Cos(a + GEOM_PI / 2#), centroidY - axisLength * Sin(a + GEOM_PI / 2#), _
            centroidX + axisLength * Cos(a + GEOM_PI / 2#), centroidY + axisLength * Sin(a + GEOM_PI / 2#), "RC_NDM_Axes", 3

        AddAcadCircle ms, centroidX, centroidY, MaxDouble(axisLength * 0.018, 5#), "RC_NDM_Axes", 3
    End If
    If loadPointEnabled Then DrawLoadPointMarker ms, loadReferenceX, loadReferenceY, MaxDouble(axisLength * 0.035, 8#)
End Sub

' Подключается только к уже открытому AutoCAD.
' Если приложение не запущено, возвращаем свою русскую ошибку вместо COM-текста
' вроде "ActiveX component can't create object".
Public Function ConnectToRunningAutoCAD() As Object
    Dim progID As Variant, candidate As Object
    ' GetObject без имени файла только подключается к ROT, не запускает CAD.
    ' Версионные ProgID нужны, если общий ключ поврежден другой DWG-программой.
    For Each progID In Array("AutoCAD.Application", "AutoCAD.Application.25.1", "AutoCAD.Application.25", _
            "AutoCAD.Application.24.3", "AutoCAD.Application.24.2", "AutoCAD.Application.24.1", _
            "AutoCAD.Application.24", "AutoCAD.Application.23.1", "AutoCAD.Application.23", "AutoCAD.Application.22")
        Set candidate = RunningAutodeskAutoCAD(CStr(progID))
        If Not candidate Is Nothing Then Set ConnectToRunningAutoCAD = candidate: Exit Function
    Next progID
    Err.Raise vbObjectError + 4310, "ConnectToRunningAutoCAD", _
        "Не найден доступный Autodesk AutoCAD. Откройте AutoCAD с нужным чертежом, завершите активную команду или диалог и повторите действие."
End Function

' Проверяет именно запущенный Autodesk acad.exe, а не OEM-приложение,
' занявшее его COM-регистрацию. Ошибки одной отсутствующей регистрации не
' мешают проверить остальные версии; никакой CreateObject здесь нет.
Private Function RunningAutodeskAutoCAD(ByVal progID As String) As Object
    On Error GoTo NotAvailable
    Dim candidate As Object, executable As String
    Set candidate = GetObject(, progID)
    executable = LCase$(CStr(candidate.FullName))
    If Right$(executable, 9) = "\acad.exe" Then Set RunningAutodeskAutoCAD = candidate
NotAvailable:
End Function

' Возвращает активный чертеж AutoCAD.
' Ситуация, когда AutoCAD открыт без документа, обрабатывается отдельно, чтобы
' пользователь понял, что нужно создать или открыть DWG, а не искать ошибку в расчете.
Private Function ActiveAutoCADDocument(ByVal acad As Object) As Object
    On Error Resume Next
    Set ActiveAutoCADDocument = acad.ActiveDocument
    On Error GoTo 0
    If ActiveAutoCADDocument Is Nothing Then
        Err.Raise vbObjectError + 4311, "ActiveAutoCADDocument", _
            "AutoCAD открыт, но активного чертежа нет. Откройте или создайте чертеж и повторите экспорт."
    End If
End Function

' Собирает настройки слоев, цветов, подписей и главных осей для одного экспорта.
' Общий settings-reader проверяет типы, профиль Results выбирает величину/state;
' этот набор управляет только построением и не меняет расчетные материалы.
Private Function ReadAutoCADExportSettings(ByVal settings As CSystemSettingsReader, _
        Optional ByVal forCleanup As Boolean = False) As TAutoCADExportSettings
    Dim options As TAutoCADExportSettings
    With options
        ReadAutoCADCommonMaterialLayers settings, .ConcreteLayer, .RebarLayer
        .ConcreteTensionLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Layer.ConcreteTension")
        .ConcreteCompressionLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Layer.ConcreteCompression")
        .RebarTensionLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Layer.RebarTension")
        .RebarCompressionLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Layer.RebarCompression")
        If Not forCleanup Then
            .ConcreteTensionColor = RequiredAutoCADColor(settings, "AutoCAD.Color.ConcreteTension")
            .ConcreteCompressionColor = RequiredAutoCADColor(settings, "AutoCAD.Color.ConcreteCompression")
            .RebarTensionColor = RequiredAutoCADColor(settings, "AutoCAD.Color.RebarTension")
            .RebarCompressionColor = RequiredAutoCADColor(settings, "AutoCAD.Color.RebarCompression")
            .NeutralColor = RequiredAutoCADColor(settings, "AutoCAD.Color.Neutral")
            .IncludeElementNames = AutoCADLabelModeIncludesNames(settings.GetRequiredChoice("AutoCAD.Export.LabelMode", Array("ValuesOnly", "NamesAndValues")))
            .NeutralLineEnabled = settings.GetRequiredBoolean("AutoCAD.Export.NeutralLineEnabled")
            .PrincipalAxesMode = AutoCADPrincipalAxesMode(settings)
            .LoadPointEnabled = settings.GetRequiredBoolean("AutoCAD.Export.LoadPointEnabled")
            .ContourEnabled = settings.GetRequiredBoolean("AutoCAD.Export.ContourEnabled")
            .ExportCrackInteractionContour = settings.GetRequiredBoolean("AutoCAD.Export.ExportCrackInteractionContour")
        End If
        If forCleanup Or .ContourEnabled Then
            .ContourLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.SectionContourLayer")
            .OpeningContourLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.OpeningContourLayer")
        End If
        If forCleanup Or .ExportCrackInteractionContour Then .CrackInteractionLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Export.CrackInteractionLayer")
    End With
    Dim cleanupLayers As Object
    Set cleanupLayers = AutoCADCleanupLayerSet(options)
    If cleanupLayers.Exists(options.ConcreteLayer) Then RaiseGeometryLayerCollision settings, "AutoCAD.Common.ConcreteLayer"
    If cleanupLayers.Exists(options.RebarLayer) Then RaiseGeometryLayerCollision settings, "AutoCAD.Common.RebarLayer"
    ReadAutoCADExportSettings = options
End Function

' Читает единственные общие слои бетона и арматуры для обоих направлений.
' Совпадающие имена запрещены одинаково при импорте и экспорте, иначе
' последующий импорт не сможет различить материалы сохраненной геометрии.
Public Sub ReadAutoCADCommonMaterialLayers(ByVal settings As CSystemSettingsReader, _
        ByRef concreteLayer As String, ByRef rebarLayer As String)
    concreteLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.ConcreteLayer")
    rebarLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.RebarLayer")
    If StrComp(concreteLayer, rebarLayer, vbTextCompare) = 0 Then Err.Raise vbObjectError + 4405, "ReadAutoCADCommonMaterialLayers", _
        settings.InputErrorMessage("AutoCAD.Common.RebarLayer", "Слои бетона и арматуры совпадают; тип Region нельзя определить однозначно.", _
            "Введите разные имена AutoCAD.Common.ConcreteLayer и AutoCAD.Common.RebarLayer.")
End Sub

' Проверяет общий для import/export синтаксис длинных имен AutoCAD-слоев.
' Config-адрес берется у reader-а; длина/символы проверяются до изменения DWG.
Public Function RequiredAutoCADLayerName(ByVal settings As CSystemSettingsReader, ByVal key As String) As String
    Dim value As String, i As Long, character As String, invalid As Boolean
    value = settings.GetRequiredString(key)
    invalid = (Len(value) > 255)
    For i = 1 To Len(value)
        character = Mid$(value, i, 1)
        If AscW(character) >= 0 And AscW(character) < 32 Then invalid = True
        If InStr(1, "<>/\" & Chr$(34) & ":;?,*|='", character, vbBinaryCompare) > 0 Then invalid = True
    Next i
    If invalid Then Err.Raise vbObjectError + 4385, "RequiredAutoCADLayerName", _
        settings.InputErrorMessage(key, "Имя слоя AutoCAD недопустимо.", _
            "Введите имя длиной от 1 до 255 символов без управляющих символов и <>/\" & Chr$(34) & ":;?,*|='.")
    RequiredAutoCADLayerName = value
End Function

' Цвет применяется и к сущностям, и к новому Layer: допустимы ACI 1..255.
' Специальные ByBlock/ByLayer не являются цветами слоя и не заменяют ввод.
Private Function RequiredAutoCADColor(ByVal settings As CSystemSettingsReader, ByVal key As String) As Long
    Dim value As Long: value = settings.GetRequiredLong(key)
    If value < 1 Or value > 255 Then Err.Raise vbObjectError + 4386, "RequiredAutoCADColor", _
        settings.InputErrorMessage(key, "Индекс цвета AutoCAD выходит за допустимый диапазон.", _
            "Введите целый индекс ACI от 1 до 255; 0 (ByBlock) и 256 (ByLayer) здесь не допускаются.")
    RequiredAutoCADColor = value
End Function

' Защищает Region-слои от технического оформления и будущей очистки.
' Общие annotation-слои допустимы; запрещено смешивать их с геометрией.
Private Sub RaiseGeometryLayerCollision(ByVal settings As CSystemSettingsReader, ByVal key As String)
    Err.Raise vbObjectError + 4387, "ReadAutoCADExportSettings", _
        settings.InputErrorMessage(key, "Слой геометрии совпадает со слоем оформления, который очищается программой.", _
            "Введите отдельное имя слоя бетона или арматуры, отличное от слоев подписей, контура и RC_NDM_Axes/LoadPoint/NeutralLine/Warnings.")
End Sub

' Читает обязательный режим осей. Совместимый Boolean-ключ используется
' только при отсутствии PrincipalAxesMode; пустое каноническое поле ошибочно.
Private Function AutoCADPrincipalAxesMode(ByVal settings As CSystemSettingsReader) As String
    Dim rawValue As String
    If Not settings.HasKey("AutoCAD.Export.PrincipalAxesMode") And settings.HasKey("AutoCAD.Export.PrincipalAxesEnabled") Then
        If settings.GetRequiredBoolean("AutoCAD.Export.PrincipalAxesEnabled") Then
            rawValue = "Transformed"
        Else
            rawValue = "None"
        End If
    Else
        rawValue = settings.GetRequiredString("AutoCAD.Export.PrincipalAxesMode")
    End If
    On Error GoTo InvalidMode
    AutoCADPrincipalAxesMode = NormalizePrincipalAxesMode(rawValue, "AutoCAD.Export.PrincipalAxesMode")
    Exit Function
InvalidMode:
    Err.Raise vbObjectError + 4313, "AutoCADPrincipalAxesMode", _
        settings.InputErrorMessage("AutoCAD.Export.PrincipalAxesMode", "Вариант главных осей не распознан.", _
            "Выберите Transformed, Concrete или None.")
End Function

' Приводит пользовательский выбор осей к единому внутреннему тексту.
Private Function NormalizePrincipalAxesMode(ByVal rawValue As String, ByVal settingKey As String) As String
    Select Case LCase$(Trim$(rawValue))
        Case "transformed", "приведенное", "приведенное сечение"
            NormalizePrincipalAxesMode = "Transformed"
        Case "concrete", "бетон", "бетонное", "бетонное сечение"
            NormalizePrincipalAxesMode = "Concrete"
        Case "none", "no", "off", "нет", "не выводить"
            NormalizePrincipalAxesMode = "None"
        Case Else
            Err.Raise vbObjectError + 4313, "ReadAutoCADExportSettings", _
                settingKey & " должен быть Transformed, Concrete или None."
    End Select
End Function

' True, если оси нужно выгружать.
Private Function PrincipalAxesModeDraws(ByVal axesMode As String) As Boolean
    PrincipalAxesModeDraws = (StrComp(axesMode, "None", vbTextCompare) <> 0)
End Function

' Определяет, нужны ли расчетные имена элементов в выбранном режиме подписей.
' Режим только значений не добавляет имена, выключенные подписи не рисуются.
Private Function AutoCADLabelModeIncludesNames(ByVal labelMode As String) As Boolean
    Select Case LCase$(Trim$(labelMode))
        Case "namesandvalues"
            AutoCADLabelModeIncludesNames = True
        Case "valuesonly"
            AutoCADLabelModeIncludesNames = False
        Case Else
            Err.Raise vbObjectError + 4312, "ReadAutoCADExportSettings", _
                "AutoCAD.Export.LabelMode должен быть ValuesOnly или NamesAndValues."
    End Select
End Function

' Определяет рамку чертежа по бетонным элементам, учитывая их поворот и размер.
' Пустая модель отклоняется явно; арматура не заменяет отсутствующий бетон.
Private Sub GetSectionBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section Is Nothing Then Err.Raise vbObjectError + 4340, "GetSectionBounds", "Модель сечения не передана."
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4341, "GetSectionBounds", "В модели сечения нет бетонных элементов."

    Dim i As Long
    Dim hasBounds As Boolean
    For i = 1 To section.ConcreteCount
        ExpandSectionBoundsByConcreteElement section, i, minX, maxX, minY, maxY, hasBounds
    Next i
End Sub

' Размеры неизвестной формы нужны только для условного изображения элемента.
' Если реальные ширина/высота сохранены, используются они; иначе рисуется
' квадрат равной площади без замены расчетных A/I импортированного элемента.
Private Function ConcreteDrawWidth(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawWidth = section.ConcreteWidth(index)
    If ConcreteDrawWidth <= 0# Then ConcreteDrawWidth = Sqr(section.ConcreteArea(index))
End Function

' Возвращает высоту визуальной оболочки бетонного элемента. Если импорт
' не сохранил габарит, для рисунка используется сторона равноплощадного
' квадрата; расчетные Area/I и напряжения элемента не заменяются.
Private Function ConcreteDrawHeight(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawHeight = section.ConcreteHeight(index)
    If ConcreteDrawHeight <= 0# Then ConcreteDrawHeight = Sqr(section.ConcreteArea(index))
End Function

' Возвращает визуальный угол бетонного элемента для AutoCAD export.
' Снимок Results уже содержит окончательные направления оболочек, включая
' нулевой угол известной грани и назначенное среднее для круговых областей.
Private Function ConcreteDrawRotation(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawRotation = section.ConcreteRotation(index)
End Function

' Возвращает точку подписи бетонного элемента в его локальной системе осей.
' Для повернутых AutoCAD Region подпись остается рядом с тем же локальным
' углом элемента, а не уезжает в осевой прямоугольник глобальных X/Y.
Private Sub ConcreteLabelPoint(ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal rotationRad As Double, _
        ByRef labelX As Double, ByRef labelY As Double)
    Dim c As Double
    Dim s As Double
    c = Cos(rotationRad)
    s = Sin(rotationRad)

    labelX = x - 0.45# * width * c + 0.1# * height * s
    labelY = y - 0.45# * width * s - 0.1# * height * c
End Sub

' Расширяет габарит AutoCAD export по фактическим углам бетонного элемента.
' Это важно для импортированных повернутых прямоугольных Region: нейтральная
' линия, оси и предупреждения получают рамку по той же геометрии, которая
' реально будет выгружена в AutoCAD.
Private Sub ExpandSectionBoundsByConcreteElement(ByVal section As CSectionModel, ByVal index As Long, _
        ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double, _
        ByRef hasBounds As Boolean)
    Dim width As Double
    Dim height As Double
    Dim rotationRad As Double
    width = ConcreteDrawWidth(section, index)
    height = ConcreteDrawHeight(section, index)
    rotationRad = ConcreteDrawRotation(section, index)

    IncludeRotatedRectangleCorner section.ConcreteX(index), section.ConcreteY(index), width, height, rotationRad, _
        -1#, -1#, minX, maxX, minY, maxY, hasBounds
    IncludeRotatedRectangleCorner section.ConcreteX(index), section.ConcreteY(index), width, height, rotationRad, _
        1#, -1#, minX, maxX, minY, maxY, hasBounds
    IncludeRotatedRectangleCorner section.ConcreteX(index), section.ConcreteY(index), width, height, rotationRad, _
        1#, 1#, minX, maxX, minY, maxY, hasBounds
    IncludeRotatedRectangleCorner section.ConcreteX(index), section.ConcreteY(index), width, height, rotationRad, _
        -1#, 1#, minX, maxX, minY, maxY, hasBounds
End Sub

' Добавляет в общий габарит один угол прямоугольного элемента с учетом
' локального поворота. sx/sy равны -1 или 1 и выбирают нужный угол.
Private Sub IncludeRotatedRectangleCorner(ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal rotationRad As Double, _
        ByVal sx As Double, ByVal sy As Double, _
        ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double, _
        ByRef hasBounds As Boolean)
    Dim hw As Double
    Dim hh As Double
    Dim c As Double
    Dim s As Double
    Dim px As Double
    Dim py As Double
    hw = width / 2#
    hh = height / 2#
    c = Cos(rotationRad)
    s = Sin(rotationRad)

    px = x + sx * hw * c - sy * hh * s
    py = y + sx * hw * s + sy * hh * c

    If Not hasBounds Then
        minX = px: maxX = px
        minY = py: maxY = py
        hasBounds = True
    Else
        minX = MinDouble(minX, px)
        maxX = MaxDouble(maxX, px)
        minY = MinDouble(minY, py)
        maxY = MaxDouble(maxY, py)
    End If
End Sub

' Ставит крест и окружность в сохраненной точке приложения усилий.
' Размер маркера служит читаемости чертежа и не задает эксцентриситет нагрузки.
Private Sub DrawLoadPointMarker(ByVal ms As Object, ByVal x As Double, ByVal y As Double, ByVal size As Double)
    AddAcadLine ms, x - size, y, x + size, y, "RC_NDM_LoadPoint", 2
    AddAcadLine ms, x, y - size, x, y + size, "RC_NDM_LoadPoint", 2
    AddAcadCircle ms, x, y, size * 0.65, "RC_NDM_LoadPoint", 2
End Sub

' Строит прямую нулевой деформации сохраненного state по пересечениям с рамкой.
' Для постоянной деформации без кривизн линии нет; при далекой нулевой линии
' рамка расширяется ограниченно, чтобы сохранить понятный масштаб чертежа.
Private Sub DrawNeutralLineByState(ByVal ms As Object, ByVal section As CSectionModel, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, ByVal colorIndex As Long)
    If Abs(kappaX) + Abs(kappaY) <= 0.000000000000001 Then Exit Sub

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY

    Dim margin As Double
    margin = 0.1 * MaxDouble(maxX - minX, maxY - minY)
    minX = minX - margin
    maxX = maxX + margin
    minY = minY - margin
    maxY = maxY + margin
    ExpandNeutralBoundsByState epsilon0, kappaX, kappaY, minX, maxX, minY, maxY

    Dim pointX() As Double
    Dim pointY() As Double
    ReDim pointX(1 To 4)
    ReDim pointY(1 To 4)
    Dim pointCount As Long

    AppendNeutralIntersection pointX, pointY, pointCount, minX, _
        NeutralYAtXState(epsilon0, kappaX, kappaY, minX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, maxX, _
        NeutralYAtXState(epsilon0, kappaX, kappaY, maxX), minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtYState(epsilon0, kappaX, kappaY, minY), minY, minX, maxX, minY, maxY
    AppendNeutralIntersection pointX, pointY, pointCount, _
        NeutralXAtYState(epsilon0, kappaX, kappaY, maxY), maxY, minX, maxX, minY, maxY

    If pointCount < 2 Then Exit Sub

    Dim i As Long
    Dim j As Long
    Dim bestI As Long
    Dim bestJ As Long
    Dim bestDistance2 As Double
    For i = 1 To pointCount - 1
        For j = i + 1 To pointCount
            Dim distance2 As Double
            distance2 = (pointX(i) - pointX(j)) ^ 2 + (pointY(i) - pointY(j)) ^ 2
            If distance2 > bestDistance2 Then
                bestDistance2 = distance2
                bestI = i
                bestJ = j
            End If
        Next j
    Next i

    If bestDistance2 > 0# Then
        AddAcadLine ms, pointX(bestI), pointY(bestI), pointX(bestJ), pointY(bestJ), _
            "RC_NDM_NeutralLine", colorIndex
    End If
End Sub

' Увеличивает только рамку изображения, если нулевая линия лежит вне сечения.
' Расстояние берется из плоскости деформаций; ограничение двадцатью габаритами
' предотвращает несоразмерный чертеж и не меняет вычисленное НДС.
Private Sub ExpandNeutralBoundsByState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    Dim normKappa As Double
    normKappa = Sqr(kappaX * kappaX + kappaY * kappaY)
    If normKappa <= 0.000000000000001 Then Exit Sub

    Dim centerX As Double
    Dim centerY As Double
    centerX = 0.5 * (minX + maxX)
    centerY = 0.5 * (minY + maxY)

    Dim sectionSize As Double
    sectionSize = MaxDouble(maxX - minX, maxY - minY)
    If sectionSize <= 0# Then sectionSize = 100#

    Dim halfSize As Double
    halfSize = 0.5 * sectionSize

    Dim distanceToNeutral As Double
    distanceToNeutral = Abs(epsilon0 + kappaX * centerY + kappaY * centerX) / normKappa

    If distanceToNeutral <= halfSize Then Exit Sub

    Dim expandedHalfSize As Double
    expandedHalfSize = MinDouble(distanceToNeutral + 0.5 * sectionSize, 20# * sectionSize)
    If expandedHalfSize <= halfSize Then Exit Sub

    minX = centerX - expandedHalfSize
    maxX = centerX + expandedHalfSize
    minY = centerY - expandedHalfSize
    maxY = centerY + expandedHalfSize
End Sub

' Координаты нулевой линии при заданной X/Y. Почти нулевой делитель означает
' отсутствие пересечения с соответствующей стороной рамки; большой служебный
' результат затем отклоняется проверкой границ, а не используется как точка.
Private Function NeutralYAtXState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByVal x As Double) As Double
    If Abs(kappaX) <= 0.000000000000001 Then
        NeutralYAtXState = 1E+99
    Else
        NeutralYAtXState = -(epsilon0 + kappaY * x) / kappaX
    End If
End Function

' Находит X нулевой деформации на заданной горизонтали по сохраненной
' плоскости НДС. При нулевом kappaY возвращает заведомо внешнюю координату:
' эта горизонталь не дает отдельного пересечения для отсечения линии.
Private Function NeutralXAtYState(ByVal epsilon0 As Double, ByVal kappaX As Double, _
        ByVal kappaY As Double, ByVal y As Double) As Double
    If Abs(kappaY) <= 0.000000000000001 Then
        NeutralXAtYState = 1E+99
    Else
        NeutralXAtYState = -(epsilon0 + kappaX * y) / kappaY
    End If
End Function

' Добавляет допустимое пересечение нулевой линии с рамкой без повторов углов.
' Проверяет вместимость массива до записи; из этих точек выбирается отрезок.
Private Sub AppendNeutralIntersection(ByRef pointX() As Double, ByRef pointY() As Double, _
        ByRef pointCount As Long, ByVal x As Double, ByVal y As Double, _
        ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double)
    If x < minX - 0.0000001 Or x > maxX + 0.0000001 Then Exit Sub
    If y < minY - 0.0000001 Or y > maxY + 0.0000001 Then Exit Sub

    Dim i As Long
    For i = 1 To pointCount
        If Abs(pointX(i) - x) < 0.0000001 And Abs(pointY(i) - y) < 0.0000001 Then Exit Sub
    Next i

    If pointCount >= UBound(pointX) Then Exit Sub
    pointCount = pointCount + 1
    pointX(pointCount) = x
    pointY(pointCount) = y
End Sub

' Строит повернутый прямоугольник четырьмя WCS-отрезками и преобразует в Region.
' Это сохраняет расположение элемента независимо от UCS/OCS активного чертежа;
' временные исходные линии удаляются после создания области.
Private Sub AddAcadRectangleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal rotationRad As Double, _
        ByVal layerName As String, ByVal colorIndex As Long)
    Dim hw As Double
    Dim hh As Double
    Dim c As Double
    Dim s As Double
    hw = width / 2#
    hh = height / 2#
    c = Cos(rotationRad)
    s = Sin(rotationRad)

    Dim px(0 To 3) As Double
    Dim py(0 To 3) As Double
    px(0) = x - hw * c + hh * s: py(0) = y - hw * s - hh * c
    px(1) = x + hw * c + hh * s: py(1) = y + hw * s - hh * c
    px(2) = x + hw * c - hh * s: py(2) = y + hw * s + hh * c
    px(3) = x - hw * c - hh * s: py(3) = y - hw * s + hh * c

    ' Строим Region из WCS-линий. LWPOLYLINE принимает 2D OCS-координаты и
    ' может дать заметный перенос при нестандартной UCS/плоскости чертежа.
    Dim sourceObjects(0 To 3) As Object
    Set sourceObjects(0) = AddAcadSourceLine(ms, px(0), py(0), px(1), py(1))
    Set sourceObjects(1) = AddAcadSourceLine(ms, px(1), py(1), px(2), py(2))
    Set sourceObjects(2) = AddAcadSourceLine(ms, px(2), py(2), px(3), py(3))
    Set sourceObjects(3) = AddAcadSourceLine(ms, px(3), py(3), px(0), py(0))
    AddAcadRegionFromCurves ms, sourceObjects, layerName, colorIndex
End Sub

' Создает временный отрезок по WCS-точкам для последующего AddRegion.
Private Function AddAcadSourceLine(ByVal ms As Object, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double) As Object
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    p1(0) = x1: p1(1) = y1: p1(2) = 0#
    p2(0) = x2: p2(1) = y2: p2(2) = 0#
    Set AddAcadSourceLine = ms.AddLine(p1, p2)
End Function

' Добавляет постоянный WCS-отрезок в ModelSpace с назначенным слоем и цветом.
' В отличие от source-line для Region этот объект остается на чертеже.
Private Sub AddAcadLine(ByVal ms As Object, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p1(0 To 2) As Double
    Dim p2(0 To 2) As Double
    p1(0) = x1: p1(1) = y1: p1(2) = 0#
    p2(0) = x2: p2(1) = y2: p2(2) = 0#
    Dim entity As Object
    Set entity = ms.AddLine(p1, p2)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

' Создает круглую Region арматуры через временную окружность; диаметр берется
' из сохраненной модели, а не восстанавливается из площади бетонного элемента.
Private Sub AddAcadCircleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim source As Object
    Set source = ms.AddCircle(p, radius)
    AddAcadRegionFromCurve ms, source, layerName, colorIndex
End Sub

' Передает одну замкнутую исходную кривую общему созданию Region.
' Владелец преобразования назначает оформление и удаляет временную кривую.
Private Sub AddAcadRegionFromCurve(ByVal ms As Object, ByVal source As Object, _
        ByVal layerName As String, ByVal colorIndex As Long)
    Dim sourceObjects(0 To 0) As Object
    Set sourceObjects(0) = source
    AddAcadRegionFromCurves ms, sourceObjects, layerName, colorIndex
End Sub

' Превращает временные AutoCAD-кривые в Region и удаляет исходные линии/окружности.
Private Sub AddAcadRegionFromCurves(ByVal ms As Object, ByRef sourceObjects() As Object, _
        ByVal layerName As String, ByVal colorIndex As Long)
    On Error GoTo Failed
    Dim regions As Variant
    regions = ms.AddRegion(sourceObjects)

    Dim entity As Object
    Set entity = regions(LBound(regions))
    entity.Layer = layerName
    entity.Color = colorIndex
    DeleteAcadSourceObjects sourceObjects
    Exit Sub

Failed:
    DeleteAcadSourceObjects sourceObjects
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Удаляет временные кривые, из которых был построен Region.
Private Sub DeleteAcadSourceObjects(ByRef sourceObjects() As Object)
    On Error Resume Next
    Dim i As Long
    For i = LBound(sourceObjects) To UBound(sourceObjects)
        If Not sourceObjects(i) Is Nothing Then sourceObjects(i).Delete
    Next i
    On Error GoTo 0
End Sub

' Гарантирует наличие слоя AutoCAD. Если слой уже есть в чертеже, его цвет и
' другие свойства не меняются; цвет применяется только к вновь созданному слою.
Private Sub EnsureAcadLayer(ByVal doc As Object, ByVal layerName As String, ByVal colorIndex As Long)
    layerName = Trim$(layerName)
    If Len(layerName) = 0 Then Err.Raise vbObjectError + 4385, "EnsureAcadLayer", "Не задано имя слоя AutoCAD для экспорта."

    On Error Resume Next
    Dim layer As Object
    Set layer = doc.Layers.Item(layerName)
    Err.Clear
    On Error GoTo Failed
    If layer Is Nothing Then
        Set layer = doc.Layers.Add(layerName)
        layer.Color = colorIndex
    End If
    Exit Sub
Failed:
    Err.Raise vbObjectError + 4388, "EnsureAcadLayer", _
        "Не удалось создать или настроить слой AutoCAD """ & layerName & """. " & _
        "Проверьте допустимость имени для текущего чертежа, возможность записи и завершите активную команду AutoCAD."
End Sub

' Добавляет WCS-окружность маркера с заданным слоем и цветом.
' Она служит оформлению, не преобразуется в расчетную область сечения.
Private Sub AddAcadCircle(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim entity As Object
    Set entity = ms.AddCircle(p, radius)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

' Добавляет готовую подпись в WCS-точке с выбранными высотой, слоем и цветом.
' Текст уже сформирован из snapshot; здесь нет назначения расчетного статуса.
Private Sub AddAcadText(ByVal ms As Object, ByVal value As String, ByVal x As Double, ByVal y As Double, _
        ByVal height As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim entity As Object
    Set entity = ms.AddText(value, p, height)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

' Выбирает цвет AutoCAD по физическому состоянию, сохраненному в Results.
' Compression/Tension получают материал-зависимые цвета, а NearZero и
' InactiveTensionConcrete выводятся нейтральным серым.
Private Function ResultColorByPhysicalState(ByVal materialType As String, ByVal physicalState As String, _
        ByRef exportSettings As TAutoCADExportSettings) As Long
    If StrComp(physicalState, "Compression", vbTextCompare) = 0 Then
        If StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
            ResultColorByPhysicalState = exportSettings.RebarCompressionColor
        Else
            ResultColorByPhysicalState = exportSettings.ConcreteCompressionColor
        End If
    ElseIf StrComp(physicalState, "Tension", vbTextCompare) = 0 Then
        If StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
            ResultColorByPhysicalState = exportSettings.RebarTensionColor
        Else
            ResultColorByPhysicalState = exportSettings.ConcreteTensionColor
        End If
    Else
        ResultColorByPhysicalState = exportSettings.NeutralColor
    End If
End Function

' Помещает подписи в tension/compression-слои по PhysicalState.
' Нейтральные и выключенные растянутые бетонные элементы уходят в
' compression-слой материала, чтобы нулевые значения не теряли слой вывода.
Private Function ResultAnnotationLayerByPhysicalState(ByVal materialType As String, ByVal physicalState As String, _
        ByRef exportSettings As TAutoCADExportSettings) As String
    If StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        If StrComp(physicalState, "Tension", vbTextCompare) = 0 Then
            ResultAnnotationLayerByPhysicalState = exportSettings.RebarTensionLayer
        Else
            ResultAnnotationLayerByPhysicalState = exportSettings.RebarCompressionLayer
        End If
    Else
        If StrComp(physicalState, "Tension", vbTextCompare) = 0 Then
            ResultAnnotationLayerByPhysicalState = exportSettings.ConcreteTensionLayer
        Else
            ResultAnnotationLayerByPhysicalState = exportSettings.ConcreteCompressionLayer
        End If
    End If
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' ============================== ДЛЯ ТЕСТОВ ==============================

' ДЛЯ ТЕСТОВ: исполняет рабочий exporter на запретившем чтение ModelSpace
' fixture. Не создает приложение AutoCAD, слои и не меняет расчетные данные.
Public Function SP35ExportSavedContoursToModelSpaceForTests(ByVal ms As Object, _
        Optional ByVal crackRegion As Boolean = False) As Long
    Dim sweeps As Object, prefix As String
    prefix = "CONTOUR_": If crackRegion Then prefix = "CRACK_REGION_"
    Set sweeps = ReadSavedContourArcSweeps(ThisWorkbook, , prefix)
    SP35ExportSavedContoursToModelSpaceForTests = DrawSavedContourRows(ThisWorkbook, ms, _
        "TEST_OUTER", sweeps, , prefix, "TEST_OPENING")
End Function

' Выводит только сохраненные материальные контуры в собственный документ
' теста. Вызывает рабочие preflight и polyline exporter без поиска НДС.
Public Function SP35ExportSavedMaterialContoursForTests(ByVal doc As Object, _
        Optional ByVal outerLayer As String = "SP35_SAVED_CONTOURS", Optional ByVal openingLayer As String = "") As Long
    Dim sweeps As Object
    Set sweeps = ReadSavedContourArcSweeps(ThisWorkbook)
    EnsureAcadLayer doc, outerLayer, SECTION_CONTOUR_COLOR_INDEX
    If Len(openingLayer) > 0 Then EnsureAcadLayer doc, openingLayer, OPENING_CONTOUR_COLOR_INDEX
    SP35ExportSavedMaterialContoursForTests = DrawSavedContourRows(ThisWorkbook, doc.ModelSpace, outerLayer, sweeps, , "CONTOUR_", openingLayer)
End Function

' ДЛЯ ТЕСТОВ: проверяет штатную фильтрацию области по состоянию/флагу и
' записывает сохраненные замкнутые контуры в собственный CAD-документ без solver/search.
Public Function SP35ExportSavedCrackRegionForTests(ByVal doc As Object, ByVal combinationID As String, _
        ByVal stateType As String, ByVal enabled As Boolean) As Long
    If Not enabled Or Not IsCrackExportState(stateType) Then Exit Function
    Dim rows As Object, sweeps As Object
    Set rows = CurrentCrackRegionRows(ThisWorkbook, combinationID)
    If rows.Count = 0 Then Exit Function
    Set sweeps = ReadSavedContourArcSweeps(ThisWorkbook, rows, "CRACK_REGION_")
    EnsureAcadLayer doc, "SP35_SAVED_CRACK_REGION", CRACK_REGION_COLOR_INDEX
    SP35ExportSavedCrackRegionForTests = DrawSavedContourRows(ThisWorkbook, doc.ModelSpace, _
        "SP35_SAVED_CRACK_REGION", sweeps, rows, "CRACK_REGION_")
End Function

' Возвращает реальные подготовленные углы экспортного контура без AutoCAD.
' Использует тот же preflight, который выполняется до изменения DWG; тест
' проверяет парсинг/адрес ошибки, но не заявляет приемку фактической записи DWG.
Public Function Audit03ReadContourArcSweepsForTests(ByVal workbook As Object) As String
    Dim sweeps As Object, key As Variant
    Set sweeps = ReadSavedContourArcSweeps(workbook)
    For Each key In sweeps.Keys
        Audit03ReadContourArcSweepsForTests = Audit03ReadContourArcSweepsForTests & CStr(key) & "|" & CStr(sweeps.Item(key)) & vbLf
    Next key
End Function

' Читает ровно тот snapshot, который используется при экспорте в AutoCAD,
' не создавая AutoCAD application и не запуская solver. Текст нужен только
' для точного сравнения сохраненных значений после изменения Config.
Public Function Audit02ReadExportSnapshotForTests(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    units.LoadFromSettings settings
    Dim section As CSectionModel
    Dim values As Object, physicalStates As Object
    Dim combinationID As String, profileID As String, stateType As String, quantity As String
    Dim precision As Long
    Dim eps0 As Double, kx As Double, ky As Double
    Dim referenceX As Double, referenceY As Double
    Dim centroidX As Double, centroidY As Double, angle As Double
    Dim extensionUsed As Boolean, warningText As String
    ReadResultsExportState workbook, settings, units, "Concrete", section, values, physicalStates, _
        combinationID, profileID, stateType, quantity, precision, eps0, kx, ky, _
        referenceX, referenceY, centroidX, centroidY, angle, extensionUsed, warningText
    Dim result As String
    result = combinationID & "|" & profileID & "|" & stateType & "|" & quantity & "|" & _
        CStr(extensionUsed) & "|" & warningText & "|" & CStr(eps0) & "|" & _
        CStr(kx) & "|" & CStr(ky) & "|" & _
        CStr(referenceX) & "|" & CStr(referenceY)
    Dim key As Variant
    For Each key In values.Keys
        result = result & vbCrLf & CStr(key) & "|" & CStr(CDbl(values(key))) & "|" & CStr(physicalStates(key))
    Next key
    Audit02ReadExportSnapshotForTests = result
End Function

' ДЛЯ ТЕСТОВ: счетчик одной операции и отсутствие удержанных таблиц/книги.
' Обычный экспорт не обращается к этой диагностике и не сохраняет модель.
Public Function PerformanceExportReadDiagnosticsForTests() As String
    PerformanceExportReadDiagnosticsForTests = "tables=" & CStr(mSnapshotTableReadCount) & ";released=" & _
        CStr(mSnapshotReadTables Is Nothing And mSnapshotReadWorkbook Is Nothing)
End Function

' Проверяет рабочие правила оформления без создания DWG. Возвращает тот же
' текст, выбранный слой и цвет, которыми пользуется цикл экспорта; флаги сами
' по себе не подтверждают фактическое создание графики в настоящем AutoCAD.
Public Function Audit03AutoCADPresentationForTests(ByVal settings As CSystemSettingsReader, _
        ByVal materialType As String, ByVal physicalState As String, _
        ByVal elementID As String, ByVal value As Double, ByVal precision As Long) As Object
    Dim options As TAutoCADExportSettings, result As Object
    options = ReadAutoCADExportSettings(settings)
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Label", ResultLabelText(elementID, value, options.IncludeElementNames, precision)
    result.Add "Color", ResultColorByPhysicalState(materialType, physicalState, options)
    result.Add "AnnotationLayer", ResultAnnotationLayerByPhysicalState(materialType, physicalState, options)
    If StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        result.Add "GeometryLayer", options.RebarLayer
    Else
        result.Add "GeometryLayer", options.ConcreteLayer
    End If
    result.Add "AxesMode", options.PrincipalAxesMode
    result.Add "NeutralLine", options.NeutralLineEnabled
    result.Add "LoadPoint", options.LoadPointEnabled
    result.Add "Contour", options.ContourEnabled
    result.Add "ContourLayer", options.ContourLayer
    Set Audit03AutoCADPresentationForTests = result
End Function

' Запускает настоящий алгоритм очистки на переданном ModelSpace fixture.
' Тестовые сущности только отмечают Delete; эта проверка не заменяет DWG,
' но выявляет удаление Region и ошибочную передачу выбранных слоев.
Public Function Audit03CleanupAutoCADModelSpaceForTests(ByVal modelSpace As Object, _
        ByVal settings As CSystemSettingsReader) As Long
    Dim options As TAutoCADExportSettings
    options = ReadAutoCADExportSettings(settings, True)
    Audit03CleanupAutoCADModelSpaceForTests = DeleteAutoCADEntitiesOnLayers(modelSpace, AutoCADCleanupLayerSet(options))
End Function

' Выполняет весь рабочий export pipeline в переданный собственный тестовый
' документ, не выбирая случайный ActiveDocument. Возвращает число добавленных
' сущностей; вызывающий тест проверяет их настоящие CAD-свойства и геометрию.
Public Function Audit03ExportAutoCADDocumentForTests(ByVal doc As Object) As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, options As TAutoCADExportSettings
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set units = New CUnitSystem: units.LoadFromSettings settings
    options = ReadAutoCADExportSettings(settings)
    Dim section As CSectionModel, values As Object, physicalStates As Object
    Dim combinationID As String, profileID As String, stateType As String, quantity As String, precision As Long
    Dim eps0 As Double, kx As Double, ky As Double, referenceX As Double, referenceY As Double
    Dim centroidX As Double, centroidY As Double, angle As Double, extended As Boolean, warning As String
    ReadResultsExportState ThisWorkbook, settings, units, options.PrincipalAxesMode, section, values, physicalStates, _
        combinationID, profileID, stateType, quantity, precision, eps0, kx, ky, referenceX, referenceY, _
        centroidX, centroidY, angle, extended, warning
    Dim before As Long, contours As Long: before = doc.ModelSpace.Count
    DrawResultsStressExport section, values, physicalStates, eps0, kx, ky, referenceX, referenceY, _
        centroidX, centroidY, angle, precision, warning, options, contours, doc
    Audit03ExportAutoCADDocumentForTests = doc.ModelSpace.Count - before
End Function

' Проверяет настоящий Layer API тем же методом, который вызывает экспорт.
' Важна сохранность существующего слоя и отсутствие проглоченной ошибки Add.
Public Sub Audit03EnsureAutoCADLayerForTests(ByVal doc As Object, ByVal name As String, ByVal color As Long)
    EnsureAcadLayer doc, name, color
End Sub


