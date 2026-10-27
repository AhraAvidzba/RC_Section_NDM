Attribute VB_Name = "modAutoCADStressExport"
Option Explicit

' ==========================================================================
' Экспорт сохранённого сечения и рассчитанного НДС в AutoCAD
' ==========================================================================
' Читает сохранённые таблицы Results. Для готового НДС строит области, подписи,
' оси и нейтральную линию; для сформированной геометрии выводит только сечение.
' При экспорте исходная геометрия повторно не импортируется и НДС не пересчитывается.

Private Type TAutoCADExportSettings
    ConcreteLayer As String ' Слой геометрии бетона, общий для импорта и экспорта.
    RebarLayer As String ' Слой геометрии арматуры, общий для импорта и экспорта.
    ConcreteTensionLayer As String ' Слой текстов результатов растянутого бетона; геометрия остаётся на ConcreteLayer.
    ConcreteCompressionLayer As String ' Слой текстов результатов сжатого бетона.
    RebarTensionLayer As String ' Слой текстов результатов растянутой арматуры.
    RebarCompressionLayer As String ' Слой текстов результатов сжатой арматуры.
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
    ContourFormat As String ' Формат только исходных контуров сечения, не расчетной области трещин.
    ContourLayer As String
    OpeningContourLayer As String
    ExportCrackInteractionContour As Boolean ' Вывод готовой области только для выбранного НДС трещин, без нового расчёта.
    CrackInteractionLayer As String
End Type

Private Const EXTENSION_WARNING_TEXT As String = "ВНЕ ФИЗИЧЕСКОЙ ДИАГРАММЫ МАТЕРИАЛА"
Private Const NUMERICAL_STATE_WARNING_TEXT As String = "ПРЯМОЕ НДС НЕ СОШЛОСЬ"
Private Const GEOMETRY_ONLY_EXPORT_WARNING As String = "Сохранённые результаты расчёта сечения отсутствуют." & vbCrLf & vbCrLf & "В AutoCAD передана только геометрия сечения."
Private Const EXTENSION_WARNING_LAYER As String = "RC_NDM_Warnings"
Private Const SECTION_CONTOUR_COLOR_INDEX As Long = 30 ' Наружная граница бетона: оранжевый ACI 30.
Private Const OPENING_CONTOUR_COLOR_INDEX As Long = 4 ' Контуры отверстий: голубой ACI 4.
Private Const CRACK_REGION_COLOR_INDEX As Long = 31 ' Расчетная область взаимодействия: ACI 31.
Private Const CONTOUR_POINT_TOLERANCE As Double = 0.000001 ' Допуск совпадения соседних вершин при экспорте, мм.
Private Const TEMPORARY_SOURCE_BATCH_SIZE As Long = 1024 ' Ограничение числа живых исходных кривых одной партии.

' Выгружает сохранённое НДС либо только сформированную геометрию в активный DWG.
' Geometry-only всегда сопровождается предупреждением, независимо от необязательных
' сообщений. Ошибки сохранённого расчёта не переключают экспорт на геометрию.
Public Sub ExportSectionStressToAutoCAD()
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim units As CUnitSystem, exportSettings As TAutoCADExportSettings
    Dim section As CSectionModel, resultByID As Object, physicalStateByID As Object
    Dim combinationID As String, profileId As String, stateType As String, quantity As String
    Dim resultPrecision As Long, epsilon0 As Double, kappaX As Double, kappaY As Double
    Dim loadReferenceX As Double, loadReferenceY As Double, centroidX As Double, centroidY As Double
    Dim principalAngle As Double, extensionUsed As Boolean, stateWarningText As String, geometryOnly As Boolean
    ReadResultsExportState ThisWorkbook, settings, units, vbNullString, section, resultByID, physicalStateByID, combinationID, _
        profileId, stateType, quantity, resultPrecision, epsilon0, kappaX, kappaY, loadReferenceX, loadReferenceY, _
        centroidX, centroidY, principalAngle, extensionUsed, stateWarningText, geometryOnly
    Dim contourExportCount As Long, acad As Object, exportDocument As Object, destinationMessage As String
    If geometryOnly Then
        exportSettings = ReadGeometryExportSettings(settings, section.Contours)
        Set acad = ConnectToRunningAutoCAD(): Set exportDocument = ActiveAutoCADDocument(acad)
        destinationMessage = AutoCADExportDestinationMessage(exportDocument)
        DrawFormedGeometryExport ThisWorkbook, section, exportSettings, contourExportCount, exportDocument
        ShowWorkbookMessage WorkbookMessageAfterOutcome(GEOMETRY_ONLY_EXPORT_WARNING, _
            SnapshotActionDataOrigin(ThisWorkbook, "Export", True)) & vbCrLf & vbCrLf & destinationMessage, vbExclamation, "RC Section NDM"
        Exit Sub
    End If
    Dim informationEnabled As Boolean: informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled")
    exportSettings = ReadAutoCADExportSettings(settings)
    Dim crackRows As Object, crackSweeps As Object
    If exportSettings.ExportCrackInteractionContour And IsCrackExportState(stateType) Then
        Set crackRows = CurrentCrackRegionRows(ThisWorkbook, combinationID)
        Set crackSweeps = ReadSavedContourArcSweeps(ThisWorkbook, crackRows, "CRACK_REGION_")
    End If
    Set acad = ConnectToRunningAutoCAD(): Set exportDocument = ActiveAutoCADDocument(acad)
    destinationMessage = AutoCADExportDestinationMessage(exportDocument)
    DrawResultsStressExport section, resultByID, physicalStateByID, epsilon0, kappaX, kappaY, _
        loadReferenceX, loadReferenceY, centroidX, centroidY, principalAngle, resultPrecision, stateWarningText, exportSettings, _
        contourExportCount, exportDocument, crackRows, crackSweeps
    If informationEnabled Then
        Dim message As String
        message = "Экспорт в AutoCAD завершен." & vbCrLf & vbCrLf & SnapshotActionDataOrigin(ThisWorkbook, "Export") & vbCrLf & vbCrLf & _
            "Волокон бетона: " & CStr(section.ConcreteCount) & _
            "; стержней арматуры: " & CStr(section.RebarCount) & _
            "; объектов контура: " & ContourExportStatusText(exportSettings.ContourEnabled, contourExportCount) & _
            "; сочетание: " & combinationID & "; профиль: " & profileId & "; состояние: " & stateType & "; величина: " & quantity & vbCrLf & vbCrLf & destinationMessage
        ShowWorkbookMessage message, vbInformation, "RC Section NDM"
    End If
    Exit Sub
Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Экспорт в AutoCAD не выполнен", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Называет фактический DWG без пути и место вывода. Оба рабочих экспортёра
' добавляют объекты в ModelSpace, независимо от открытого пользователем листа.
Private Function AutoCADExportDestinationMessage(ByVal document As Object) As String
    AutoCADExportDestinationMessage = "Сечение выгружено в чертёж «" & CStr(document.Name) & "» — в модель."
End Function

' Очищает активный чертеж AutoCAD от объектов оформления, созданных NDM-export.
' Геометрия бетона, арматуры, наружных границ и отверстий остается. Удаляются
' только подписи на annotation-слоях, нейтральная линия, главные оси и точка
' приложения нагрузки, чтобы можно было заново выгрузить расчет поверх той же
' геометрии.
Public Sub ClearAutoCADDrawing()
    On Error GoTo Failed

    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook ThisWorkbook
    Dim informationEnabled As Boolean
    informationEnabled = settings.GetRequiredBoolean("General.NonCriticalMessagesEnabled") ' Используем тот же снимок Config и проверяем флаг до удаления объектов.

    Dim exportSettings As TAutoCADExportSettings
    exportSettings = ReadAutoCADExportSettings(settings, True)

    Dim acad As Object
    Set acad = ConnectToRunningAutoCAD()

    Dim doc As Object
    Set doc = ActiveAutoCADDocument(acad)

    Dim layers As Object: Set layers = AutoCADCleanupLayerSet(exportSettings)
    Dim hadLayers As Boolean, deletedCount As Long
    deletedCount = DeleteAutoCADDocumentAnnotations(doc, layers, hadLayers)
    If Not hadLayers Then
        If informationEnabled Then ShowWorkbookMessage _
            "В активном чертеже AutoCAD «" & CStr(doc.Name) & "» нет слоёв с подписями и оформлением RC Section NDM." & vbCrLf & vbCrLf & _
            "Очистка не требуется.", vbInformation, "RC Section NDM"
        Exit Sub
    End If

    If informationEnabled Then
        ShowWorkbookMessage "Чертеж AutoCAD очищен от объектов оформления RC Section NDM." & vbCrLf & _
            "Удалено объектов: " & CStr(deletedCount) & "." & vbCrLf & _
            "Бетон, арматура, наружные контуры и отверстия оставлены без изменений.", vbInformation, "RC Section NDM"
    End If
    Exit Sub

Failed:
    ShowWorkbookMessage WorkbookErrorMessage("Очистка чертежа AutoCAD не выполнена", Err.Number, Err.Source, Err.Description), vbExclamation, "RC Section NDM"
End Sub

' Проверяет только таблицу слоёв активного чертежа, не перебирая его объекты.
' Защищённые контурные слои не считаются основанием для очистки.
Private Function HasAutoCADCleanupLayers(ByVal doc As Object, ByVal layers As Object) As Boolean
    Dim name As Variant, layer As Object, available As Object
    Set available = doc.Layers
    For Each name In layers.Keys
        If CBool(layers.Item(name)) Then
            Set layer = Nothing
            On Error Resume Next
            Set layer = available.Item(CStr(name))
            On Error GoTo 0
            If Not layer Is Nothing Then HasAutoCADCleanupLayers = True: Exit Function
        End If
    Next name
End Function

' Проверяет наличие слоёв до первого доступа к ModelSpace. Перестроение
' чертежа требуется только после фактического удаления объектов оформления.
Private Function DeleteAutoCADDocumentAnnotations(ByVal doc As Object, ByVal layers As Object, ByRef hadLayers As Boolean) As Long
    hadLayers = HasAutoCADCleanupLayers(doc, layers)
    If Not hadLayers Then Exit Function
    DeleteAutoCADDocumentAnnotations = DeleteAutoCADEntitiesOnLayers(doc.ModelSpace, layers)
    If DeleteAutoCADDocumentAnnotations > 0 Then doc.Regen 1
End Function

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
    If Len(exportSettings.ContourLayer) > 0 Then layers.Item(exportSettings.ContourLayer) = False ' Общие контурные слои защищаются целиком, включая созданные NDM объекты, даже при совпадении с другим слоем оформления.
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
' Region всегда сохраняются. На защищенных общих слоях сохраняются все
' объекты независимо от метки NDM; удаляются только объекты оформления.
Private Function DeleteAutoCADEntitiesOnLayers(ByVal ms As Object, ByVal layers As Object) As Long
    Dim i As Long
    For i = ms.Count - 1 To 0 Step -1
        Dim entity As Object
        Set entity = ms.Item(i)
        If layers.Exists(CStr(entity.Layer)) And StrComp(CStr(entity.ObjectName), "AcDbRegion", vbTextCompare) <> 0 Then
            If CBool(layers.Item(CStr(entity.Layer))) Then
                entity.Delete
                DeleteAutoCADEntitiesOnLayers = DeleteAutoCADEntitiesOnLayers + 1
            End If
        End If
    Next i
End Function

' Подготавливает сохранённую геометрию, значения элементов, поле деформаций
' и предупреждение о физической допустимости НДС. Профиль определяет состояние
' и отображаемую величину; текущие параметры сечения не заменяют результаты расчёта.
Private Sub ReadResultsExportState(ByVal workbook As Object, ByVal settings As CSystemSettingsReader, _
        ByVal units As CUnitSystem, ByVal principalAxesMode As String, _
        ByRef section As CSectionModel, ByRef resultByID As Object, ByRef physicalStateByID As Object, ByRef combinationID As String, _
        ByRef profileId As String, ByRef stateType As String, ByRef quantity As String, ByRef resultPrecision As Long, _
        ByRef epsilon0 As Double, ByRef kappaX As Double, ByRef kappaY As Double, _
        ByRef loadReferenceX As Double, ByRef loadReferenceY As Double, _
        ByRef centroidX As Double, ByRef centroidY As Double, ByRef principalAngle As Double, _
        ByRef extensionUsed As Boolean, ByRef stateWarningText As String, _
        Optional ByRef geometryOnly As Boolean = False)
    On Error GoTo Failed
    geometryOnly = False
    BeginSectionSnapshotRead workbook
    If IsGeometryOnlyExportSource(workbook) Then
        Dim geometrySource As String: geometrySource = ResultsGeometrySource(workbook)
        If Len(geometrySource) = 0 Then Err.Raise vbObjectError + 4380, "ReadResultsExportState", _
            "В Results нет сформированного сечения. Сначала нажмите «Сформировать сечение», затем повторите экспорт в AutoCAD."
        Set section = ReadSectionGeometryFromResults(workbook, geometrySource)
        Set resultByID = CreateObject("Scripting.Dictionary"): Set physicalStateByID = CreateObject("Scripting.Dictionary")
        geometryOnly = True: stateWarningText = GEOMETRY_ONLY_EXPORT_WARNING
        combinationID = vbNullString: profileId = vbNullString: stateType = vbNullString: quantity = vbNullString
        EndSectionSnapshotRead
        Exit Sub
    End If
    If units Is Nothing Then
        Set units = New CUnitSystem: units.LoadFromSettings settings
    End If
    If Len(principalAxesMode) = 0 Then principalAxesMode = AutoCADPrincipalAxesMode(settings)

    combinationID = ResolveExportCombinationID(workbook, settings.GetRequiredString("AutoCAD.Export.CombinationID"))
    profileId = ReadProfileIdForLoadCase(workbook, combinationID)
    If Len(profileId) = 0 Then Err.Raise vbObjectError + 4366, "ReadResultsExportState", _
        settings.InputErrorMessage("AutoCAD.Export.CombinationID", _
            "В сохраненных Results не найдено сочетание " & combinationID & " с расчетным профилем.", _
            "Выберите Worst или ID сохраненного сочетания; для новых нагрузок сначала выполните расчет.")
    Set section = ReadSectionGeometryFromResults(workbook, "Results")

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    profiles.LoadFromWorkbook workbook, True, Not snapshot Is Nothing

    Dim profile As CCalculationProfile
    If profiles.HasProfile(profileId) Then
        Set profile = profiles.ProfileById(profileId)
    Else
        If snapshot Is Nothing Then Set profile = profiles.ProfileById(profileId) Else Set profile = snapshot.SavedVisualizationProfile(profileId)
    End If
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
    EndSectionSnapshotRead
    Exit Sub
Failed:
    Dim errorNumber As Long, errorSource As String, errorText As String
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    EndSectionSnapshotRead
    Err.Raise errorNumber, errorSource, errorText
End Sub

' Разрешает geometry-only только для preview без результатов и любых записей LC/НДС.
' Явный файловый GeometryPreview также не содержит состояний. Полный File-снимок
' и повреждённый сохранённый расчёт остаются на строгом пути чтения.
Private Function IsGeometryOnlyExportSource(ByVal workbook As Object) As Boolean
    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then
        IsGeometryOnlyExportSource = snapshot.IsGeometryPreview
        Exit Function
    End If
    Dim data As Variant: data = ReadPhysicalResultTable(workbook, "rngNDMElementResults")
    If HasResultTableRows(data) Then Exit Function
    data = ReadPhysicalResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then IsGeometryOnlyExportSource = True: Exit Function
    Dim caseColumn As Long: caseColumn = ResultColumn(data, "LoadCase")
    Dim parameterColumn As Long: parameterColumn = ResultColumn(data, "Parameter")
    Dim row As Long, caseID As String, parameter As String
    For row = 2 To UBound(data, 1)
        caseID = Trim$(CStr(data(row, caseColumn)))
        If Len(caseID) > 0 And StrComp(caseID, "ALL", vbTextCompare) <> 0 Then Exit Function
        parameter = LCase$(Trim$(CStr(data(row, parameterColumn))))
        If parameter = "profileid" Or InStr(parameter, ".epsilon0") > 0 Or InStr(parameter, ".kappax") > 0 Or _
                InStr(parameter, ".kappay") > 0 Or InStr(parameter, ".stateavailable") > 0 Or InStr(parameter, ".statestatus") > 0 Then Exit Function
    Next row
    IsGeometryOnlyExportSource = True
End Function

' Находит профиль выбранного сочетания в сохранённых Results.
' Читает rngNDMElementResults; свойства сочетания нужны для пояснения ошибки,
' если выбранное расчётное состояние не было записано.
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
    data = ReadElementResultTable(workbook, combinationID, , True)
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

' Читает напряжения, деформации и физическое состояние элементов для выбранного
' сочетания и вида НДС. Если такого НДС нет, сообщает об этом с адресом таблицы;
' результат другого сочетания не подставляется.
Private Sub ReadElementResultsForCombination(ByVal workbook As Object, ByVal resultType As String, _
        ByRef combinationID As String, ByVal stateType As String, _
        ByVal resultByID As Object, ByVal physicalStateByID As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim data As Variant
    data = ReadElementResultTable(workbook, combinationID, stateType)
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
            resultByID(CStr(data(rowIndex, colID))) = ReadExportSnapshotNumber(data, anchor, rowIndex, colValue)
            physicalStateByID(CStr(data(rowIndex, colID))) = CStr(data(rowIndex, colState))
        End If
    Next rowIndex

    If Not found Then Err.Raise vbObjectError + 4354, "ReadElementResultsForCombination", _
        MissingExportStateMessage(combinationID, stateType)
End Sub


' Читает только числовую ячейку сохранённого снимка перед экспортом. Числовые Variant
' сохраняют точность, текст проходит общую полную грамматику; ошибка называет реальный Name-адрес.
Private Function ReadExportSnapshotNumber(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long) As Double
    Dim parsed As Double
    If TryReadStrictDouble(data(row, column), parsed) Then
        ReadExportSnapshotNumber = parsed: Exit Function
    End If
    Err.Raise vbObjectError + 4361, "ReadResultsExportState", _
        "В сохранённом снимке Results поле " & ResultHeaderBase(CStr(data(1, column))) & _
        " содержит неверное число: " & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Повторите расчёт или заново сформируйте сечение для восстановления снимка."
End Function

' Читает сохранённое поле деформаций, точку нагрузки и выбранные главные оси.
' Переводит длины и кривизны в нужные единицы. Подготавливает предупреждение,
' если равновесие не достигнуто или для показа выбрано вспомогательное НДС.
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

    Dim anchor As Object: Set anchor = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim foundState As Boolean
    Dim stateFields As Object: Set stateFields = CreateObject("Scripting.Dictionary")
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
                    concreteCentroidX = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "concrete.centroidy"
                    concreteCentroidY = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "concrete.principalangle"
                    concretePrincipalAngle = ReadExportSnapshotNumber(data, anchor, rowIndex, colValue)
                Case "transformed.centroidx"
                    transformedCentroidX = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "transformed.centroidy"
                    transformedCentroidY = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "transformed.principalangle"
                    transformedPrincipalAngle = ReadExportSnapshotNumber(data, anchor, rowIndex, colValue)
                Case "loadreferencex"
                    loadReferenceX = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "loadreferencey"
                    loadReferenceY = OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
            End Select
        ElseIf StrComp(CStr(data(rowIndex, colLoadCase)), combinationID, vbTextCompare) = 0 Then
            Dim key As String
            Dim prefix As String
            key = LCase$(Trim$(SafeText(data(rowIndex, colParameter))))
            prefix = "state." & LCase$(stateType) & "."
            If Left$(key, Len(prefix)) = prefix Then
                stateFields(key) = True
                Select Case Mid$(key, Len(prefix) + 1)
                Case "epsilon0"
                    epsilon0 = ReadExportSnapshotNumber(data, anchor, rowIndex, colValue)
                    foundState = True
                Case "kappax"
                    kappaX = OutputCurvatureToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "kappay"
                    kappaY = OutputCurvatureToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colValue), CStr(data(rowIndex, colUnit)))
                Case "extensionused"
                    If Not TryReadSnapshotBoolean(data(rowIndex, colValue), extensionUsed) Then _
                        RejectExportSnapshotMetadata workbook, rowIndex, colValue, SafeText(data(rowIndex, colParameter))
                Case "status"
                    If Not TryReadSnapshotStatus(data(rowIndex, colValue), stateStatus) Then _
                        RejectExportSnapshotMetadata workbook, rowIndex, colValue, SafeText(data(rowIndex, colParameter))
                End Select
            End If
        End If
    Next rowIndex

    Dim field As Variant, requiredKey As String
    For Each field In Array("Epsilon0", "KappaX", "KappaY", "ExtensionUsed", "Status")
        requiredKey = "State." & stateType & "." & CStr(field)
        If Not stateFields.Exists(LCase$(requiredKey)) Then Err.Raise vbObjectError + 4361, "ReadSectionPropertiesForCombination", _
            "В сохраненном снимке Results для сочетания " & combinationID & " отсутствует обязательное поле " & requiredKey & _
            " (rngNDMSectionProperties, " & anchor.Worksheet.Name & "!" & anchor.Address(False, False) & "). Повторите расчет/импорт для восстановления снимка."
    Next field

    If extensionUsed Or stateStatus = policy.Fail Then
        stateWarningText = EXTENSION_WARNING_TEXT
    ElseIf stateStatus = policy.NumFail Then
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

' Поясняет, почему выбранного НДС нет: расчёт трещин отключён либо пороговая точка
' не найдена. Формирует сообщение экспорта, не меняя статусы расчёта.
Private Function MissingExportStateMessage(ByVal combinationID As String, ByVal stateType As String) As String
    MissingExportStateMessage = "Запрашиваемое состояние """ & stateType & _
        """ не найдено в Results для AutoCAD export, сочетание " & combinationID & "."
    If IsCrackExportState(stateType) Then
        If Not ExportCombinationHasCrackWidth(combinationID) Then
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для запрашиваемого сочетания расчет трещин не выполнялся: этому сочетанию назначен профиль, где Calculation.Crack.Width = No."
        ElseIf StrComp(stateType, "CrackedState", vbTextCompare) = 0 Then
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для этого сочетания расчёт трещин был включён, но НДС с учётом трещин (CrackedState) не сохранено в Results. Проверьте статус и пояснения расчёта трещин; при необходимости повторите расчёт."
        Else
            MissingExportStateMessage = MissingExportStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, но пороговое состояние образования трещины не записано. Это штатно только когда нет действия для поиска нормальной трещины, например чистое сжатие без момента. Если растяжение или изгиб есть, проверьте статус проверки образования трещины."
        End If
    End If
End Function

' Проверяет, относится ли выбранный вид НДС к расчёту трещин.
' Для неизвестного имени пояснение о трещинах не выдаётся.
Private Function IsCrackExportState(ByVal stateType As String) As Boolean
    On Error GoTo NotCrackState
    Select Case SectionStateTypeFromText(stateType)
        Case sstPreCrackState, sstPostCrackState, sstCrackedState
            IsCrackExportState = True
    End Select
NotCrackState:
End Function

' Проверяет, предусмотрен ли расчёт ширины трещин сохранённым профилем сочетания.
' Нужно для пояснения, почему выбранное НДС отсутствует.
Private Function ExportCombinationHasCrackWidth(ByVal combinationID As String) As Boolean
    On Error GoTo Failed
    Dim workbook As Object
    Set workbook = ThisWorkbook
    Dim profileId As String
    profileId = ReadProfileIdForLoadCase(workbook, combinationID)
    If Len(profileId) = 0 Then Exit Function

    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then
        Dim savedProfile As CCalculationProfile: Set savedProfile = snapshot.SavedVisualizationProfile(profileId)
        ExportCombinationHasCrackWidth = savedProfile.CrackWidthEnabled
        Exit Function
    End If

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook
    ExportCombinationHasCrackWidth = profiles.ProfileById(profileId).CrackWidthEnabled
Failed:
End Function

' Читает определяющее сочетание из готовой сводки. Если оно не указано,
' вызывающий код выберет другое сочетание среди сохранённых результатов.
Private Function ReadGoverningCombinationID(ByVal workbook As Object) As String
    ReadGoverningCombinationID = ReadSavedWorstCombinationID(workbook)
End Function

' Добавляет фактический адрес именованной таблицы к повреждённому полю снимка.
' Проверка значения общая со схемой; адаптер не назначает новый инженерный статус.
Private Sub RejectExportSnapshotMetadata(ByVal workbook As Object, ByVal row As Long, ByVal column As Long, ByVal field As String)
    Dim anchor As Object: Set anchor = workbook.Names.Item("rngNDMSectionProperties").RefersToRange
    Err.Raise vbObjectError + 4361, "ReadSectionPropertiesForCombination", _
        "В сохраненном снимке Results поле " & field & " (" & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & ") содержит недопустимое значение. " & _
        "Исправьте указанную ячейку или повторите расчет/импорт для восстановления снимка."
End Sub

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

' Возвращает первое сочетание из свойств Results, если определяющее сочетание
' не указано и таблица результатов элементов пуста. Наличие выбранного
' Visualization.State у этого сочетания проверяется отдельно.
Private Function FirstLoadCaseFromSectionProperties(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colLoadCase As Long
    colLoadCase = ResultColumn(data, "LoadCase")

    Dim rowIndex As Long, candidateID As String
    For rowIndex = 2 To UBound(data, 1)
        candidateID = Trim$(SafeText(data(rowIndex, colLoadCase)))
        If Len(candidateID) > 0 And StrComp(candidateID, "ALL", vbTextCompare) <> 0 Then
            FirstLoadCaseFromSectionProperties = candidateID: Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Возвращает первое сочетание с сохранёнными результатами элементов.
' Как и для схемы Excel, используется при отсутствии определяющего сочетания в сводке.
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

' Строит в активном чертеже геометрию, подписи, оси и нейтральную линию
' по сохранённым Results. Размеры элементов берёт из результатов расчёта;
' повторный импорт и решение НДС не выполняет.
Private Sub DrawResultsStressExport(ByVal section As CSectionModel, _
        ByVal resultByID As Object, ByVal physicalStateByID As Object, _
        ByVal epsilon0 As Double, ByVal kappaX As Double, ByVal kappaY As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        ByVal centroidX As Double, ByVal centroidY As Double, ByVal principalAngle As Double, _
        ByVal resultPrecision As Long, ByVal stateWarningText As String, ByRef exportSettings As TAutoCADExportSettings, _
        ByRef contourExportCount As Long, Optional ByVal targetDocument As Object = Nothing, _
        Optional ByVal crackRows As Object = Nothing, Optional ByVal crackSweeps As Object = Nothing)
    Dim arcSweeps As Object
    If exportSettings.ContourEnabled Then Set arcSweeps = ReadSavedContourArcSweeps(ThisWorkbook) ' Проверяем все углы до подключения и добавления объектов в чертеж.
    Dim doc As Object
    Set doc = targetDocument
    If doc Is Nothing Then
        Dim acad As Object
        Set acad = ConnectToRunningAutoCAD()
        Set doc = ActiveAutoCADDocument(acad)
    End If

    Dim ms As Object
    Set ms = doc.ModelSpace
    Dim temporarySources As Collection: Set temporarySources = New Collection
    On Error GoTo DrawingFailed

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
        Dim labelX As Double
        Dim labelY As Double
        ReadConcretePresentation section, i, concreteWidth, concreteHeight, concreteRotation, labelX, labelY
        resultValue = LookupResultValue(resultByID, section.ConcreteID(i))
        physicalState = ResultPhysicalState(physicalStateByID, section.ConcreteID(i))
        textHeight = 0.22 * MinDouble(concreteWidth, concreteHeight)
        If textHeight <= 0# Then textHeight = 1#
        AddAcadConcreteElement ms, section, i, exportSettings.ConcreteLayer, _
            ResultColorByPhysicalState("Concrete", physicalState, exportSettings), temporarySources

        AddAcadText ms, ResultLabelText(section.ConcreteID(i), resultValue, exportSettings.IncludeElementNames, resultPrecision), _
            labelX, labelY, textHeight, _
            ResultAnnotationLayerByPhysicalState("Concrete", physicalState, exportSettings), _
            ResultColorByPhysicalState("Concrete", physicalState, exportSettings)
        If temporarySources.Count >= TEMPORARY_SOURCE_BATCH_SIZE Then DeleteAcadTemporarySources doc, temporarySources
    Next i

    For i = 1 To section.RebarCount
        resultValue = LookupResultValue(resultByID, section.RebarID(i))
        physicalState = ResultPhysicalState(physicalStateByID, section.RebarID(i))
        AddAcadCircleRegion ms, section.RebarX(i), section.RebarY(i), section.RebarDiameter(i) / 2#, _
            exportSettings.RebarLayer, ResultColorByPhysicalState("Rebar", physicalState, exportSettings), temporarySources
        AddAcadText ms, ResultLabelText(section.RebarID(i), resultValue, exportSettings.IncludeElementNames, resultPrecision), _
            section.RebarX(i) + section.RebarDiameter(i) / 2#, section.RebarY(i) + section.RebarDiameter(i) / 2#, _
            MaxDouble(2.5, section.RebarDiameter(i) * 0.18), _
            ResultAnnotationLayerByPhysicalState("Rebar", physicalState, exportSettings), _
            ResultColorByPhysicalState("Rebar", physicalState, exportSettings)
        If temporarySources.Count >= TEMPORARY_SOURCE_BATCH_SIZE Then DeleteAcadTemporarySources doc, temporarySources
    Next i

    DeleteAcadTemporarySources doc, temporarySources
    If exportSettings.ContourEnabled Then ' Контур выводим после бетонных и арматурных объектов, чтобы он не оказался закрыт AutoCAD Region, созданными для волокон расчетной сетки.
        contourExportCount = DrawSavedContourRows(ThisWorkbook, ms, exportSettings.ContourLayer, arcSweeps, , "CONTOUR_", exportSettings.OpeningContourLayer, exportSettings.ContourFormat)
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
    Exit Sub
DrawingFailed:
    Dim errorNumber As Long, errorSource As String, errorText As String
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    On Error Resume Next
    DeleteAcadTemporarySources doc, temporarySources
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorText
End Sub

' Передаёт только сохранённую геометрию без напряжений, подписей НДС и расчётных линий.
' Сечение не формируется повторно; временные кривые принадлежат этой операции.
Private Sub DrawFormedGeometryExport(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByRef options As TAutoCADExportSettings, ByRef contourExportCount As Long, _
        Optional ByVal targetDocument As Object = Nothing)
    Dim document As Object: Set document = targetDocument
    If document Is Nothing Then
        Dim acad As Object: Set acad = ConnectToRunningAutoCAD()
        Set document = ActiveAutoCADDocument(acad)
    End If
    Dim ms As Object: Set ms = document.ModelSpace
    Dim sources As Collection: Set sources = New Collection
    Dim index As Long, number As Long, description As String, source As String
    On Error GoTo Failed
    EnsureAcadLayer document, options.ConcreteLayer, 8
    EnsureAcadLayer document, options.RebarLayer, 1
    For index = 1 To section.ConcreteCount
        AddAcadConcreteElement ms, section, index, options.ConcreteLayer, 8, sources
        If sources.Count >= TEMPORARY_SOURCE_BATCH_SIZE Then DeleteAcadTemporarySources document, sources
    Next index
    For index = 1 To section.RebarCount
        AddAcadCircleRegion ms, section.RebarX(index), section.RebarY(index), section.RebarDiameter(index) / 2#, options.RebarLayer, 1, sources
        If sources.Count >= TEMPORARY_SOURCE_BATCH_SIZE Then DeleteAcadTemporarySources document, sources
    Next index
    DeleteAcadTemporarySources document, sources
    If options.ContourEnabled And section.Contours.Count > 0 Then
        If Len(options.ContourLayer) > 0 Then EnsureAcadLayer document, options.ContourLayer, SECTION_CONTOUR_COLOR_INDEX
        If Len(options.OpeningContourLayer) > 0 Then EnsureAcadLayer document, options.OpeningContourLayer, OPENING_CONTOUR_COLOR_INDEX
        contourExportCount = DrawSavedMaterialContours(workbook, ms, options.ContourLayer, options.OpeningContourLayer, options.ContourFormat, section.Contours)
    End If
    document.Regen 1
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description: source = Err.Source
    Resume Cleanup
Cleanup:
    On Error Resume Next
    DeleteAcadTemporarySources document, sources
    On Error GoTo 0
    Err.Raise number, source, description
End Sub

' Читает только используемые параметры геометрического экспорта.
' При отключённых/отсутствующих контурах их формат и слои не потребляются.
Private Function ReadGeometryExportSettings(ByVal settings As CSystemSettingsReader, _
        ByVal contours As CSectionContours) As TAutoCADExportSettings
    Dim options As TAutoCADExportSettings
    ReadAutoCADCommonMaterialLayers settings, options.ConcreteLayer, options.RebarLayer
    options.ContourEnabled = settings.GetRequiredBoolean("AutoCAD.Export.ContourEnabled")
    If options.ContourEnabled And contours.Count > 0 Then
        options.ContourFormat = settings.GetRequiredChoice("AutoCAD.Export.ContourFormat", Array("Region", "Polyline"))
        Dim index As Long, hasOuter As Boolean, hasOpening As Boolean
        For index = 1 To contours.Count
            If contours.LoopRole(index) = "Opening" Then hasOpening = True Else hasOuter = True
        Next index
        If hasOuter Then options.ContourLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.SectionContourLayer")
        If hasOpening Then options.OpeningContourLayer = RequiredAutoCADLayerName(settings, "AutoCAD.Common.OpeningContourLayer")
    End If
    ReadGeometryExportSettings = options
End Function

' Отбирает все контуры актуальной области трещин нужного сочетания.
' Если данные устарели или относятся к разным состояниям, вся область пропускается.
' Проверяются сведения о состоянии и запуске. Область завершённой проверки
' с превышенной шириной трещин тоже пригодна для показа; новое НДС не рассчитывается.
Private Function CurrentCrackRegionRows(ByVal workbook As Object, ByVal combinationID As String) As Object
    Dim selected As Object, data As Variant, geometry As Variant, properties As Variant
    Set selected = CreateObject("Scripting.Dictionary")
    Set CurrentCrackRegionRows = selected
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    geometry = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    properties = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasResultTableRows(data) Or Not HasResultTableRows(geometry) Or Not HasResultTableRows(properties) Then Exit Function
    Dim colOwner As Long, colRevision As Long, colType As Long, colRun As Long
    colOwner = ResultColumn(data, "CombinationID"): colRevision = ResultColumn(data, "SectionRevision")
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
            If CStr(data(row, colRun)) <> CStr(runID) Or CStr(data(row, colRevision)) <> CStr(revision) Then ' Дата запуска и ревизия исключают смешение с другим сечением.
                selected.RemoveAll: Exit Function
            End If
            rowIdentity = ReadSavedAnnotationText(data, anchor, row, colStandard) & "|" & ReadSavedAnnotationText(data, anchor, row, colAnchor) ' Одна область должна принадлежать единому нормативному варианту и опорной группе.
            If Len(identity) = 0 Then identity = rowIdentity
            If rowIdentity <> identity Then selected.RemoveAll: Exit Function
            If ReadGeometryNumber(data, anchor, row, colLoop) < 1# Then Err.Raise vbObjectError + 4376, _
                "CurrentCrackRegionRows", "Номер замкнутого контура расчетной области должен быть положительным: " & _
                anchor.Parent.Name & "!" & anchor.Offset(row - 1, colLoop - 1).Address(False, False) & ". Повторите расчет."
            selected.Add CStr(row), True
        End If
    Next row
End Function

' Выводит исходные контуры в выбранном формате Region/Polyline,
' а сохранённую область трещин — замкнутыми полилиниями с точными дугами.
' Геометрию не восстанавливает по площади или сетке; отверстия сохраняются.
Private Function DrawSavedContourRows(ByVal workbook As Object, ByVal ms As Object, ByVal contourLayer As String, _
        ByVal arcSweeps As Object, Optional ByVal selectedRows As Object = Nothing, _
        Optional ByVal typePrefix As String = "CONTOUR_", Optional ByVal openingLayer As String = vbNullString, _
        Optional ByVal contourFormat As String = "Region") As Long
    If typePrefix = "CONTOUR_" Then
        DrawSavedContourRows = DrawSavedMaterialContours(workbook, ms, contourLayer, openingLayer, contourFormat)
        Exit Function
    End If
    Dim tableName As Object
    On Error Resume Next
    Set tableName = workbook.Names.Item("rngNDMSectionAnnotations")
    On Error GoTo 0
    If tableName Is Nothing Then Exit Function
    Dim anchor As Object: Set anchor = tableName.RefersToRange
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Function

    Dim colType As Long: colType = ResultColumn(data, "AnnotationType")
    Dim colStartX As Long: colStartX = ResultColumn(data, "StartX")
    Dim colStartY As Long: colStartY = ResultColumn(data, "StartY")
    Dim colEndX As Long: colEndX = ResultColumn(data, "EndX")
    Dim colEndY As Long: colEndY = ResultColumn(data, "EndY")
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")

    Dim segStartX() As Double
    Dim segStartY() As Double
    Dim segEndX() As Double
    Dim segEndY() As Double
    Dim segBulge() As Double
    Dim segCount As Long
    Dim currentLoopKey As String
    Dim currentLayer As String, rowLayer As String
    Dim colLoop As Long: colLoop = ResultColumn(data, "LoopID")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Not selectedRows Is Nothing Then
            If Not selectedRows.Exists(CStr(rowIndex)) Then GoTo NextContourRow
        End If
        Dim annotationType As String
        annotationType = UCase$(Trim$(SafeText(data(rowIndex, colType))))
        If Left$(annotationType, Len(typePrefix)) = typePrefix Then
            Dim loopKey As String
            loopKey = CStr(data(rowIndex, colLoop))
            rowLayer = contourLayer
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
                lineUnit = AnnotationLengthUnit(data, rowIndex, colUnit, anchor)
                AppendContourSegment segStartX, segStartY, segEndX, segEndY, segBulge, segCount, _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartX), lineUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartY), lineUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colEndX), lineUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colEndY), lineUnit), _
                    0#
            Case typePrefix & "ARC"
                Dim arcUnit As String
                arcUnit = AnnotationLengthUnit(data, rowIndex, colUnit, anchor)
                AppendContourSegment segStartX, segStartY, segEndX, segEndY, segBulge, segCount, _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartX), arcUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartY), arcUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colEndX), arcUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colEndY), arcUnit), _
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
                circleUnit = AnnotationLengthUnit(data, rowIndex, colUnit, anchor)
                DrawSavedContourRows = DrawSavedContourRows + _
                    DrawContourCirclePolyline(ms, rowLayer, _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartX), circleUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colStartY), circleUnit), _
                    OutputLengthToInternalByUnit(ReadExportSnapshotNumber(data, anchor, rowIndex, colEndX), circleUnit), CRACK_REGION_COLOR_INDEX)
        End Select
NextContourRow:
    Next rowIndex

    If segCount > 0 Then
        DrawSavedContourRows = DrawSavedContourRows + _
            DrawContourSegmentPolyline(ms, currentLayer, segStartX, segStartY, segEndX, segEndY, segBulge, segCount, CRACK_REGION_COLOR_INDEX)
    End If
End Function

' Выгружает достоверные границы выбранным форматом: отдельными полилиниями
' или как Region бетона с вырезанными отверстиями.
' Отдельные бетонные части получают свои Region. Если известны только отверстия,
' выводит их отдельными Region, не придумывая наружную границу по сетке.
' Временные кривые принадлежат этому вызову; существующие объекты DWG не читаются.
Private Function DrawSavedMaterialContours(ByVal workbook As Object, ByVal ms As Object, _
        ByVal outerLayer As String, ByVal openingLayer As String, ByVal contourFormat As String, _
        Optional ByVal preparedContours As CSectionContours = Nothing) As Long
    Dim asRegion As Boolean
    asRegion = (StrComp(contourFormat, "Region", vbTextCompare) = 0)
    If Not asRegion And StrComp(contourFormat, "Polyline", vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", "Формат экспорта контуров должен быть Region или Polyline."
    Dim contours As CSectionContours, loops As Object, indices As Collection
    Dim i As Long, key As String, loopKey As Variant, indexValue As Variant, layer As String, colorIndex As Long
    If preparedContours Is Nothing Then Set contours = ReadSavedSectionContours(workbook) Else Set contours = preparedContours
    Set loops = CreateObject("Scripting.Dictionary")
    For i = 1 To contours.Count
        key = UCase$(contours.LoopRole(i)) & ":" & contours.LoopID(i)
        If Not loops.Exists(key) Then
            Set indices = New Collection
            loops.Add key, indices
        End If
        loops.Item(key).Add i
    Next i
    If loops.Count = 0 Then Exit Function
    Dim shapes As Object, parents As Object, expected As Object
    Set shapes = CreateObject("Scripting.Dictionary"): Set parents = CreateObject("Scripting.Dictionary")
    Set expected = CreateObject("Scripting.Dictionary")
    If asRegion Then PrepareMaterialContourRegions contours, loops, shapes, parents, expected
    For Each loopKey In loops.Keys
        If Left$(CStr(loopKey), 8) = "OPENING:" And Len(openingLayer) = 0 Then
            If Not asRegion Or parents.Count = 0 Then Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", _
                "Для отдельных отверстий задайте AutoCAD.Common.OpeningContourLayer на Config."
        End If
    Next loopKey
    Dim created As Collection, regions As Object, curve As Object, entity As Object, parent As Object
    Dim curves(0 To 0) As Object, added As Variant, addedIndex As Long, ignored As Long
    Dim errorNumber As Long, errorDescription As String, expectedRegion As CConcreteRegion
    Set created = New Collection: Set regions = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
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
        Set curve = Nothing
        For Each indexValue In indices
            i = CLng(indexValue)
            If contours.SegmentType(i) = "CONTOUR_CIRCLE" Then
                If indices.Count <> 1 Then Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", "Круговой контур не должен содержать дополнительные сегменты."
                ignored = DrawContourCirclePolyline(ms, layer, contours.StartX(i), contours.StartY(i), _
                    contours.Radius(i), colorIndex, curve)
            Else
                AppendContourSegment sx, sy, ex, ey, bulge, count, contours.StartX(i), contours.StartY(i), _
                    contours.EndX(i), contours.EndY(i), Tan(contours.SweepAngle(i) / 4#)
            End If
        Next indexValue
        If count > 0 Then ignored = DrawContourSegmentPolyline(ms, layer, sx, sy, ex, ey, bulge, count, colorIndex, curve)
        created.Add curve
        If Not asRegion Then
            curve.Update
            DrawSavedMaterialContours = DrawSavedMaterialContours + 1
            Set curve = Nothing
            GoTo NextMaterialLoop
        End If
        Set curves(0) = curve
        added = ms.AddRegion(curves)
        If Not IsArray(added) Then Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", "AutoCAD не создал Region по сохраненному контуру."
        For addedIndex = LBound(added) To UBound(added): created.Add added(addedIndex): Next addedIndex
        If UBound(added) <> LBound(added) Then Err.Raise vbObjectError + 4371, _
            "DrawSavedMaterialContours", "По одному замкнутому контуру AutoCAD создал несколько областей. Проверьте контур в Results."
        Set entity = added(LBound(added)): entity.Layer = layer: entity.Color = colorIndex
        regions.Add CStr(loopKey), entity
        curve.Delete: Set curve = Nothing
NextMaterialLoop:
    Next loopKey
    If Not asRegion Then Exit Function
    For Each loopKey In parents.Keys ' Вычитаем каждый проем только из его ближайшей наружной области. Бетонная часть внутри проема остается отдельным Region.
        Set parent = regions.Item(parents.Item(loopKey))
        Set entity = regions.Item(loopKey)
        parent.Boolean 2, entity
    Next loopKey
    For Each loopKey In regions.Keys
        If Not parents.Exists(loopKey) Then
            Set entity = regions.Item(loopKey): Set expectedRegion = expected.Item(loopKey)
            If Abs(CDbl(entity.Area) - expectedRegion.Area) > MaxDouble(0.0001, expectedRegion.Area * 0.00000001) Then _
                Err.Raise vbObjectError + 4371, "DrawSavedMaterialContours", "Площадь Region в AutoCAD не совпадает с сохраненной геометрией. Экспорт контура отменен."
            If Left$(CStr(loopKey), 8) = "OPENING:" Then ' Boolean может перенести оформление вычитаемой области на результат. Назначаем итоговый слой и цвет после всех вычитаний.
                entity.Layer = openingLayer: entity.Color = OPENING_CONTOUR_COLOR_INDEX
            Else
                entity.Layer = outerLayer: entity.Color = SECTION_CONTOUR_COLOR_INDEX
            End If
            entity.Update
            DrawSavedMaterialContours = DrawSavedMaterialContours + 1
        End If
    Next loopKey
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    Resume FailureCleanup ' Освобождает активный обработчик до удаления уже поглощённых Boolean объектов.
FailureCleanup:
    On Error Resume Next ' Удаляем только объекты, созданные этим незавершенным вызовом.
    If Not curve Is Nothing Then curve.Delete
    For Each entity In created: entity.Delete: Next entity
    On Error GoTo 0
    Err.Raise errorNumber, "DrawSavedMaterialContours", "Не удалось вывести контуры сечения: " & errorDescription
End Function

' До изменения чертежа проверяет, к каким бетонным областям относятся отверстия.
' Сохраняет ключи и роли отдельных записей для слоёв Outer/Opening и opening-only.
' Проверки вложенности/вычитания выполняет CSectionGeometryQuery; инерции здесь не вычисляются.
Private Sub PrepareMaterialContourRegions(ByVal contours As CSectionContours, ByVal loops As Object, _
        ByVal shapes As Object, ByVal parents As Object, ByVal expected As Object)
    Dim query As CSectionGeometryQuery, key As Variant, otherKey As Variant, indices As Collection
    Dim region As CConcreteRegion, outer As CConcreteRegion, first As Long, selected As String, bestArea As Double, outerCount As Long
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, cx As Double, cy As Double, sweep As Double, loopNumber As Long
    Set query = New CSectionGeometryQuery
    For Each key In loops.Keys
        Set indices = loops.Item(key): first = CLng(indices(1))
        Set region = SavedMaterialLoopRegion(contours, indices, query)
        shapes.Add CStr(key), region: expected.Add CStr(key), region
        If StrComp(contours.LoopRole(first), "Outer", vbTextCompare) = 0 Then outerCount = outerCount + 1
    Next key
    Dim parts As Collection: Set parts = New Collection
    If outerCount = 0 Then
        For Each key In shapes.Keys: parts.Add shapes.Item(key): Next key
        Set region = query.MergeDisjointRegions(parts)
        Exit Sub
    End If
    For Each key In loops.Keys
        Set indices = loops.Item(key): first = CLng(indices(1))
        If StrComp(contours.LoopRole(first), "Opening", vbTextCompare) = 0 Then
            Set region = shapes.Item(key)
            region.GetSegment 1, x1, y1, x2, y2, cx, cy, sweep, loopNumber
            selected = vbNullString: bestArea = 1E+100
            For Each otherKey In loops.Keys
                Set indices = loops.Item(otherKey): first = CLng(indices(1))
                If StrComp(contours.LoopRole(first), "Outer", vbTextCompare) = 0 Then
                    Set outer = shapes.Item(otherKey)
                    If outer.Area < bestArea And query.ContainsPoint(outer, x1, y1, False) Then
                        selected = CStr(otherKey): bestArea = outer.Area
                    End If
                End If
            Next otherKey
            If Len(selected) = 0 Then Err.Raise vbObjectError + 4371, "PrepareMaterialContourRegions", _
                "В Results отверстие находится вне всех наружных контуров. Повторите импорт или расчет."
            Set outer = expected.Item(selected)
            Set expected.Item(selected) = query.SubtractContainedRegion(outer, region)
            parents.Add CStr(key), selected
        End If
    Next key
    For Each key In expected.Keys
        If Not parents.Exists(key) Then parts.Add expected.Item(key)
    Next key
    Set region = query.MergeDisjointRegions(parts)
End Sub

' Собирает одну точную границу из сохраненных линий, дуг или окружности.
Private Function SavedMaterialLoopRegion(ByVal contours As CSectionContours, ByVal indices As Collection, _
        ByVal query As CSectionGeometryQuery) As CConcreteRegion
    Dim first As Long: first = CLng(indices(1))
    If contours.SegmentType(first) = "CONTOUR_CIRCLE" Then
        If indices.Count <> 1 Then Err.Raise vbObjectError + 4371, "SavedMaterialLoopRegion", "Круговой контур содержит лишние сегменты."
        Set SavedMaterialLoopRegion = query.CircleRegion(contours.StartX(first), contours.StartY(first), contours.Radius(first))
        Exit Function
    End If
    Dim edges() As TRegionEdge, i As Long, sourceIndex As Long
    ReDim edges(1 To indices.Count)
    For i = 1 To indices.Count
        sourceIndex = CLng(indices(i))
        edges(i).X1 = contours.StartX(sourceIndex): edges(i).Y1 = contours.StartY(sourceIndex)
        edges(i).X2 = contours.EndX(sourceIndex): edges(i).Y2 = contours.EndY(sourceIndex)
        If contours.SegmentType(sourceIndex) = "CONTOUR_ARC" Then
            edges(i).CenterX = contours.CenterX(sourceIndex): edges(i).CenterY = contours.CenterY(sourceIndex)
            edges(i).Sweep = contours.SweepAngle(sourceIndex)
        End If
    Next i
    Set SavedMaterialLoopRegion = query.BoundaryRegion(edges, indices.Count, "SavedMaterialContour")
End Function

' До записи DWG читает проверенные знаковые углы дуг в радианах.
' Для исходных границ ключ — индекс CSectionContours; для области трещин —
' номер строки аннотаций. Отсутствующая аннотация допустима, повреждённый угол — нет.
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
    Dim colType As Long, colText As Long, colUnit As Long, rowIndex As Long, unitText As String
    colType = ResultColumn(data, "AnnotationType"): colText = ResultColumn(data, "Text")
    colUnit = ResultColumn(data, "Unit")
    For rowIndex = 2 To UBound(data, 1)
        If Not selectedRows Is Nothing Then
            If Not selectedRows.Exists(CStr(rowIndex)) Then GoTo NextSweepRow
        End If
        If Left$(UCase$(Trim$(SafeText(data(rowIndex, colType)))), Len(typePrefix)) = typePrefix Then
            unitText = AnnotationLengthUnit(data, rowIndex, colUnit, anchor)
        End If
        If StrComp(Trim$(SafeText(data(rowIndex, colType))), typePrefix & "ARC", vbTextCompare) = 0 Then
            result.Add CStr(rowIndex), ReadSavedContourArcSweep(data, anchor, rowIndex, colText)
        End If
NextSweepRow:
    Next rowIndex
End Function

' Читает обязательную собственную единицу координат расчетного контура.
' Пустое поле не заменяется текущей настройкой или внутренними миллиметрами.
Private Function AnnotationLengthUnit(ByRef data As Variant, ByVal rowIndex As Long, _
        ByVal colUnit As Long, ByVal anchor As Object) As String
    Dim unitText As String
    unitText = Trim$(ReadSavedAnnotationText(data, anchor, rowIndex, colUnit))
    If Len(unitText) = 0 Or unitText = "-" Then Err.Raise vbObjectError + 4376, _
        "AnnotationLengthUnit", "В расчетном контуре Results не задана единица длины: " & _
        anchor.Parent.Name & "!" & anchor.Offset(rowIndex - 1, colUnit - 1).Address(False, False) & ". Повторите расчет."
    AnnotationLengthUnit = unitText
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
        ByVal colorIndex As Long, _
        Optional ByRef createdEntity As Object = Nothing) As Long
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
    Set createdEntity = entity
    For segmentIndex = 1 To segmentCount
        If Abs(bulge(segmentIndex)) > 0.000000000001 Then Call entity.SetBulge(segmentIndex - 1, bulge(segmentIndex))
    Next segmentIndex
    entity.Closed = True
    entity.Update
    DrawContourSegmentPolyline = 1
End Function

' Окружность тоже выводится LWPOLYLINE: четыре четверти с одинаковым bulge дают
' непрерывную замкнутую полилинию с дугами, а не отдельный объект Circle.
Private Function DrawContourCirclePolyline(ByVal ms As Object, ByVal contourLayer As String, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double, _
        ByVal colorIndex As Long, _
        Optional ByRef createdEntity As Object = Nothing) As Long
    If radius <= 0# Then Exit Function

    Dim points(0 To 7) As Double
    points(0) = centerX + radius: points(1) = centerY
    points(2) = centerX: points(3) = centerY + radius
    points(4) = centerX - radius: points(5) = centerY
    points(6) = centerX: points(7) = centerY - radius

    Dim entity As Object
    Set entity = AddAcadLightWeightPolyline(ms, points, contourLayer, colorIndex)
    Set createdEntity = entity

    Dim quarterBulge As Double
    quarterBulge = Tan((GEOM_PI / 2#) / 4#)
    Dim i As Long
    For i = 0 To 3
        Call entity.SetBulge(i, quarterBulge)
    Next i
    entity.Closed = True
    entity.Update
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
    Dim entity As Object, errorNumber As Long, errorDescription As String
    On Error GoTo Failed
    Set entity = ms.AddLightWeightPolyline(points)
    entity.Layer = layerName: entity.Color = colorIndex
    Set AddAcadLightWeightPolyline = entity
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next
    If Not entity Is Nothing Then entity.Delete
    On Error GoTo 0
    Err.Raise errorNumber, "AddAcadLightWeightPolyline", errorDescription
End Function

' Сравнивает соседние вершины контура в миллиметрах.
Private Function PointsAreClose(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Boolean
    PointsAreClose = (Abs(x1 - x2) <= CONTOUR_POINT_TOLERANCE And Abs(y1 - y2) <= CONTOUR_POINT_TOLERANCE)
End Function

' Добавляет заметное предупреждение под сечением. Читает текст из сохранённых Results:
' например, о выходе за физическую диаграмму или недостигнутом равновесии.
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

' Возвращает значение для конкретного элемента в выбранном НДС.
' Отсутствующее значение означает неполные результаты, а не нулевое напряжение.
Private Function LookupResultValue(ByVal resultByID As Object, ByVal elementID As String) As Double
    If resultByID.Exists(elementID) Then
        LookupResultValue = CDbl(resultByID.Item(elementID))
    Else
        Err.Raise vbObjectError + 4356, "LookupResultValue", "В Results нет выбранного результата для элемента: " & elementID
    End If
End Function

' Читает физическое состояние элемента из Results для выбора цвета и слоя AutoCAD.
' Пользовательское правило знаков N, Mx и My не меняет физический знак
' напряжений и деформаций элемента.
Private Function ResultPhysicalState(ByVal physicalStateByID As Object, ByVal elementID As String) As String
    If physicalStateByID.Exists(elementID) Then
        ResultPhysicalState = CStr(physicalStateByID.Item(elementID))
    Else
        Err.Raise vbObjectError + 4356, "ResultPhysicalState", "В Results нет PhysicalState для элемента: " & elementID
    End If
End Function

' Формирует подпись напряжения или деформации в AutoCAD.
' Значение уже прочитано из Results в пользовательских единицах; меняется только формат текста.
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

' Рисует выбранные главные оси через сохранённый центр и отмечает точку нагрузки.
' Координаты и угол берёт из Results. Габарит влияет только на длину линий осей.
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
    For Each progID In Array("AutoCAD.Application", "AutoCAD.Application.25.1", "AutoCAD.Application.25", _
            "AutoCAD.Application.24.3", "AutoCAD.Application.24.2", "AutoCAD.Application.24.1", _
            "AutoCAD.Application.24", "AutoCAD.Application.23.1", "AutoCAD.Application.23", "AutoCAD.Application.22") ' GetObject без имени файла только подключается к ROT, не запускает CAD. Версионные ProgID нужны, если общий ключ поврежден другой DWG-программой.
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

' Читает настройки слоёв, цветов, подписей и главных осей для одного экспорта.
' Типы значений проверяет CSystemSettingsReader; профиль Results задаёт величину
' и вид НДС. Эти настройки меняют только рисунок, а не материалы расчёта.
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
        If Not forCleanup And .ContourEnabled Then _
            .ContourFormat = settings.GetRequiredChoice("AutoCAD.Export.ContourFormat", Array("Region", "Polyline"))
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

' Проверяет допустимость имени слоя AutoCAD для импорта и экспорта.
' Длина и символы проверяются до изменения чертежа; адрес ошибки читает CSystemSettingsReader.
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

' Читает обязательный текущий режим осей; отсутствие строки не заменяется другим ключом.
Private Function AutoCADPrincipalAxesMode(ByVal settings As CSystemSettingsReader) As String
    Dim rawValue As String
    rawValue = settings.GetRequiredString("AutoCAD.Export.PrincipalAxesMode")
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

' Возвращает сохранённый угол изображения бетонного элемента.
' Results уже содержит окончательное направление, включая нулевой угол
' известной грани и среднее направление для круглых областей.
Private Function ConcreteDrawRotation(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteDrawRotation = section.ConcreteRotation(index)
End Function

' Готовит размеры шрифта и подпись по той же сущности, которая выводится в DWG.
' Точные произвольные элементы используют реальные проекции и центр тяжести;
' эквивалентная оболочка используется у записей Adaptive без точной границы.
Private Sub ReadConcretePresentation(ByVal section As CSectionModel, ByVal index As Long, _
        ByRef width As Double, ByRef height As Double, ByRef rotation As Double, _
        ByRef labelX As Double, ByRef labelY As Double)
    If section.ConcreteHasBoundary(index) Then
        Dim geometry As CConcreteRegion, query As CSectionGeometryQuery, centerX As Double, centerY As Double
        Set geometry = section.ConcreteGeometry(index): Set query = New CSectionGeometryQuery
        If Not query.TryRectangle(geometry, centerX, centerY, width, height, rotation) Then
            Dim minX As Double, maxX As Double, minY As Double, maxY As Double
            query.ProjectionBounds geometry, 1#, 0#, minX, maxX
            query.ProjectionBounds geometry, 0#, 1#, minY, maxY
            width = maxX - minX: height = maxY - minY: rotation = 0#
            labelX = section.ConcreteX(index): labelY = section.ConcreteY(index)
            Exit Sub
        End If
    Else
        width = ConcreteDrawWidth(section, index): height = ConcreteDrawHeight(section, index)
        rotation = ConcreteDrawRotation(section, index)
    End If
    ConcreteLabelPoint section.ConcreteX(index), section.ConcreteY(index), width, height, rotation, labelX, labelY
End Sub

' Возвращает координаты подписи бетонного элемента в WCS, мм.
' Смещение сначала задаётся относительно локального угла элемента,
' затем поворачивается вместе с ним; глобальный габарит не подменяет локальные оси.
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
    If section.ConcreteHasBoundary(index) Then
        Dim region As CConcreteRegion, query As CSectionGeometryQuery
        Dim firstX As Double, lastX As Double, firstY As Double, lastY As Double
        Set region = section.ConcreteGeometry(index): Set query = New CSectionGeometryQuery
        query.ProjectionBounds region, 1#, 0#, firstX, lastX
        query.ProjectionBounds region, 0#, 1#, firstY, lastY
        If Not hasBounds Then
            minX = firstX: maxX = lastX: minY = firstY: maxY = lastY: hasBounds = True
        Else
            minX = MinDouble(minX, firstX): maxX = MaxDouble(maxX, lastX)
            minY = MinDouble(minY, firstY): maxY = MaxDouble(maxY, lastY)
        End If
        Exit Sub
    End If
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

' Строит прямую нулевой деформации по сохранённому НДС и её пересечениям с рамкой.
' При постоянной деформации без кривизн линия не выводится. Для далёкой нейтральной
' линии рамка расширяется с ограничением, чтобы рисунок оставался читаемым.
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

' Выгружает одну универсальную бетонную ячейку по ее фактическим границам.
' Прямоугольная специализация использует компактный WCS-путь; криволинейная
' область сохраняет дуги, отверстия и все несвязные части без изменения площади.
Private Sub AddAcadConcreteElement(ByVal ms As Object, ByVal section As CSectionModel, ByVal index As Long, _
        ByVal layerName As String, ByVal colorIndex As Long, ByVal temporarySources As Collection)
    If Not section.ConcreteHasBoundary(index) Then
        AddAcadRectangleRegion ms, section.ConcreteX(index), section.ConcreteY(index), ConcreteDrawWidth(section, index), _
            ConcreteDrawHeight(section, index), ConcreteDrawRotation(section, index), layerName, colorIndex, temporarySources
        Exit Sub
    End If
    Dim region As CConcreteRegion, query As CSectionGeometryQuery
    Dim x As Double, y As Double, width As Double, height As Double, rotation As Double, ignored As Long
    Set region = section.ConcreteGeometry(index): Set query = New CSectionGeometryQuery
    If query.TryRectangle(region, x, y, width, height, rotation) Then
        AddAcadRectangleRegion ms, x, y, width, height, rotation, layerName, colorIndex, temporarySources
    Else
        ignored = AddAcadConcreteGeometry(ms, region, layerName, colorIndex, temporarySources)
    End If
End Sub

' Создаёт CAD Region для связных компонентов, подготовленных общим ядром.
' Принадлежность отверстий определяет ConnectedRegions; экспорт создаёт объекты,
' вычитает openings и проверяет площадь. При отказе удаляет только свои Region.
Private Function AddAcadConcreteGeometry(ByVal ms As Object, ByVal geometry As CConcreteRegion, _
        ByVal layerName As String, ByVal colorIndex As Long, ByVal temporarySources As Collection) As Long
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Dim components As Collection: Set components = query.ConnectedRegions(geometry)
    Dim component As CConcreteRegion, loopRegion As CConcreteRegion
    Dim edges() As TRegionEdge, count As Long
    Dim loopID As Long, outerLoop As Long, entity As Object, parent As Object
    Dim created As Collection: Set created = New Collection
    Dim actualArea As Double, number As Long, description As String
    On Error GoTo Failed
    For Each component In components
        component.CopyEdges edges, count
        outerLoop = 0
        For loopID = 1 To component.LoopCount
            If Not component.LoopIsOpening(loopID) Then outerLoop = loopID: Exit For
        Next loopID
        If outerLoop = 0 Then Err.Raise vbObjectError + 4371, "AddAcadConcreteGeometry", "У связного бетонного элемента отсутствует наружная граница."
        Set loopRegion = AcadConcreteBoundaryLoop(component, outerLoop, edges, count, query)
        Set parent = AddAcadConcreteLoop(ms, loopRegion, layerName, colorIndex, temporarySources)
        created.Add parent
        For loopID = 1 To component.LoopCount
            If component.LoopIsOpening(loopID) Then
                Set loopRegion = AcadConcreteBoundaryLoop(component, loopID, edges, count, query)
                Set entity = AddAcadConcreteLoop(ms, loopRegion, layerName, colorIndex, temporarySources)
                created.Add entity
                parent.Boolean 2, entity ' Вычитается только отверстие текущей готовой связной компоненты.
            End If
        Next loopID
        actualArea = actualArea + CDbl(parent.Area)
        parent.Layer = layerName: parent.Color = colorIndex
        AddAcadConcreteGeometry = AddAcadConcreteGeometry + 1
    Next component
    If Abs(actualArea - geometry.Area) > MaxDouble(0.0001, geometry.Area * 0.0000001) Then _
        Err.Raise vbObjectError + 4371, "AddAcadConcreteGeometry", "Площадь выгруженного бетонного элемента не совпадает с сохраненным фактическим контуром."
    Exit Function
Failed:
    number = Err.Number: description = Err.Description
    Resume FailureCleanup ' Выходит из активного обработчика до COM-очистки поглощённых Boolean объектов.
FailureCleanup:
    On Error Resume Next
    For Each entity In created: entity.Delete: Next entity
    On Error GoTo 0
    Err.Raise number, "AddAcadConcreteGeometry", description
End Function

' Выбирает уже классифицированный ядром loop для создания отдельного CAD Region.
' Не определяет его владельца: компоненту и роль отверстия передаёт ConnectedRegions.
Private Function AcadConcreteBoundaryLoop(ByVal component As CConcreteRegion, ByVal loopID As Long, _
        ByRef edges() As TRegionEdge, ByVal count As Long, ByVal query As CSectionGeometryQuery) As CConcreteRegion
    Dim one() As TRegionEdge, selectedCount As Long, i As Long
    For i = 1 To count
        If edges(i).LoopID = loopID Then
            selectedCount = selectedCount + 1: ReDim Preserve one(1 To selectedCount)
            LSet one(selectedCount) = edges(i): one(selectedCount).LoopID = 0
        End If
    Next i
    Set AcadConcreteBoundaryLoop = query.BoundaryRegion(one, selectedCount, component.Source)
End Function

' Передает один проверенный loop как WCS LWPOLYLINE с точными bulge дуг.
' Список источников принадлежит общей операции: в нем нет исходных объектов DWG.
Private Function AddAcadConcreteLoop(ByVal ms As Object, ByVal region As CConcreteRegion, _
        ByVal layerName As String, ByVal colorIndex As Long, ByVal temporarySources As Collection) As Object
    Dim edges() As TRegionEdge, count As Long, i As Long, points() As Double, source As Object
    region.CopyEdges edges, count
    ReDim points(0 To count * 2 - 1)
    For i = 1 To count
        points((i - 1) * 2) = edges(i).X1: points((i - 1) * 2 + 1) = edges(i).Y1
    Next i
    Set source = AddAcadLightWeightPolyline(ms, points, layerName, colorIndex)
    temporarySources.Add source
    Dim normal(0 To 2) As Double: normal(2) = 1#
    source.Normal = normal: source.Elevation = 0# ' WCS XY не зависит от текущей высоты плоских объектов в чертеже.
    For i = 1 To count
        If Abs(edges(i).Sweep) > 0.000000000001 Then source.SetBulge i - 1, Tan(edges(i).Sweep / 4#)
    Next i
    source.Closed = True
    Set AddAcadConcreteLoop = AddAcadRegionFromCurve(ms, source, layerName, colorIndex)
End Function

' Строит прямоугольный Region по вершинам в мировых координатах.
' Вершины задаются в WCS независимо от UCS/OCS активного чертежа;
' временные полилинии собираются для удаления одним собственным набором.
Private Function AddAcadRectangleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal width As Double, ByVal height As Double, ByVal rotationRad As Double, _
        ByVal layerName As String, ByVal colorIndex As Long, ByVal temporarySources As Collection) As Object
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

    Dim points(0 To 11) As Double, corner As Long, vertex As Long, source As Object
    For corner = 0 To 3 ' Одна замкнутая полилиния сокращает обращения к AutoCAD. Начало справа внизу задаёт воспроизводимое направление первого ребра Region при импорте.
        vertex = (corner + 1) Mod 4
        points(corner * 3) = px(vertex): points(corner * 3 + 1) = py(vertex)
    Next corner
    Set source = ms.Add3DPoly(points)
    temporarySources.Add source
    source.Closed = True
    Set AddAcadRectangleRegion = AddAcadRegionFromCurve(ms, source, layerName, colorIndex)
End Function

' Добавляет отрезок в мировых координатах с заданным слоем и цветом.
' Отрезок остаётся на чертеже, в отличие от временных линий для создания Region.
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
Private Function AddAcadCircleRegion(ByVal ms As Object, ByVal x As Double, ByVal y As Double, _
        ByVal radius As Double, ByVal layerName As String, ByVal colorIndex As Long, ByVal temporarySources As Collection) As Object
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim source As Object
    Set source = ms.AddCircle(p, radius)
    temporarySources.Add source
    Set AddAcadCircleRegion = AddAcadRegionFromCurve(ms, source, layerName, colorIndex)
End Function

' Передает одну замкнутую исходную кривую общему созданию Region.
' Владелец преобразования назначает оформление и удаляет временную кривую.
Private Function AddAcadRegionFromCurve(ByVal ms As Object, ByVal source As Object, _
        ByVal layerName As String, ByVal colorIndex As Long) As Object
    Dim sourceObjects(0 To 0) As Object
    Set sourceObjects(0) = source
    Set AddAcadRegionFromCurve = AddAcadRegionFromCurves(ms, sourceObjects, layerName, colorIndex)
End Function

' Превращает временные AutoCAD-кривые в Region в мировой плоскости XY.
' Успешно использованные исходники удаляет владелец всей выгрузки.
Private Function AddAcadRegionFromCurves(ByVal ms As Object, ByRef sourceObjects() As Object, _
        ByVal layerName As String, ByVal colorIndex As Long) As Object
    On Error GoTo Failed
    Dim regions As Variant
    regions = ms.AddRegion(sourceObjects)

    Dim entity As Object
    Set entity = regions(LBound(regions))
    entity.Layer = layerName
    entity.Color = colorIndex
    Set AddAcadRegionFromCurves = entity
    Exit Function

Failed:
    Dim errorNumber As Long, errorSource As String, errorText As String
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    DeleteAcadSourceObjects sourceObjects
    Err.Raise errorNumber, errorSource, errorText
End Function

' Удаляет временные кривые, из которых был построен Region.
Private Sub DeleteAcadSourceObjects(ByRef sourceObjects() As Object)
    On Error Resume Next
    Dim i As Long
    For i = LBound(sourceObjects) To UBound(sourceObjects)
        If Not sourceObjects(i) Is Nothing Then sourceObjects(i).Delete
    Next i
    On Error GoTo 0
End Sub

' Удаляет только исходные кривые, созданные текущей выгрузкой. Собственный
' SelectionSet никогда не содержит существующие объекты чертежа; он удаляется сразу.
' Если пакетный вызов недоступен, сохраняется обычное удаление по тем же ссылкам.
Private Sub DeleteAcadTemporarySources(ByVal doc As Object, ByRef sources As Collection)
    If sources Is Nothing Then Exit Sub
    If sources.Count = 0 Then Exit Sub
    Dim items() As Object, i As Long, selection As Object
    ReDim items(0 To sources.Count - 1)
    For i = 1 To sources.Count
        Set items(i - 1) = sources(i)
    Next i
    On Error GoTo IndividualCleanup
    Set selection = doc.SelectionSets.Add("NDM_TEMP_SOURCES_" & Hex$(CLng(Timer * 1000#)))
    selection.AddItems items
    selection.Erase
    selection.Delete
    Set sources = New Collection
    Exit Sub
IndividualCleanup:
    On Error Resume Next
    If Not selection Is Nothing Then selection.Delete
    For i = 1 To sources.Count
        sources(i).Delete
    Next i
    Set sources = New Collection
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

' Добавляет готовую подпись в мировых координатах с заданными высотой, слоем и цветом.
' Текст уже подготовлен по сохранённым Results.
Private Sub AddAcadText(ByVal ms As Object, ByVal value As String, ByVal x As Double, ByVal y As Double, _
        ByVal height As Double, ByVal layerName As String, ByVal colorIndex As Long)
    Dim p(0 To 2) As Double
    p(0) = x: p(1) = y: p(2) = 0#
    Dim entity As Object
    Set entity = ms.AddText(value, p, height)
    entity.Layer = layerName
    entity.Color = colorIndex
End Sub

' Выбирает цвет по физическому состоянию элемента из Results.
' Цвета сжатия и растяжения зависят от материала. Почти нулевые напряжения
' и неработающий растянутый бетон показываются нейтральным серым.
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

' Возвращает большее значение для габаритов, размеров подписей и масштабируемых допусков экспорта.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Возвращает меньшее значение для границ элементов и размера подписи внутри бетонного волокна.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает рабочий вывод аналитической ячейки на переданном ModelSpace.
' Не ищет DWG и не запускает расчет; убирает только созданные временные кривые.
Public Function TestExportConcreteElementGeometry(ByVal ms As Object, ByVal geometry As CConcreteRegion) As Long
    Dim sources As Collection: Set sources = New Collection
    Dim source As Object, number As Long, description As String
    On Error GoTo Failed
    TestExportConcreteElementGeometry = AddAcadConcreteGeometry(ms, geometry, "TEST_CONCRETE", 8, sources)
Cleanup:
    On Error Resume Next
    For Each source In sources: source.Delete: Next source
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "TestExportConcreteElementGeometry", description
    Exit Function
Failed:
    number = Err.Number: description = Err.Description
    Resume Cleanup
End Function

' ============================== ДЛЯ ТЕСТОВ ==============================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выполняет обычный экспорт на подставном ModelSpace,
' который запрещает чтение объектов. Не создаёт AutoCAD и слои, не меняет расчётные данные.
Public Function SP35ExportSavedContoursToModelSpaceForTests(ByVal ms As Object, _
        Optional ByVal crackRegion As Boolean = False, Optional ByVal contourFormat As String = "Region") As Long
    Dim sweeps As Object, prefix As String
    prefix = "CONTOUR_": If crackRegion Then prefix = "CRACK_REGION_"
    Set sweeps = ReadSavedContourArcSweeps(ThisWorkbook, , prefix)
    SP35ExportSavedContoursToModelSpaceForTests = DrawSavedContourRows(ThisWorkbook, ms, _
        "TEST_OUTER", sweeps, , prefix, "TEST_OPENING", contourFormat)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выводит только сохраненные материальные контуры в собственный документ
' теста. Вызывает рабочие preflight и Region exporter без поиска НДС.
Public Function SP35ExportSavedMaterialContoursForTests(ByVal doc As Object, _
        Optional ByVal outerLayer As String = "SP35_SAVED_CONTOURS", Optional ByVal openingLayer As String = "", _
        Optional ByVal contourFormat As String = "Region") As Long
    Dim sweeps As Object
    Set sweeps = ReadSavedContourArcSweeps(ThisWorkbook)
    EnsureAcadLayer doc, outerLayer, SECTION_CONTOUR_COLOR_INDEX
    If Len(openingLayer) > 0 Then EnsureAcadLayer doc, openingLayer, OPENING_CONTOUR_COLOR_INDEX
    SP35ExportSavedMaterialContoursForTests = DrawSavedContourRows(ThisWorkbook, doc.ModelSpace, outerLayer, sweeps, , "CONTOUR_", openingLayer, contourFormat)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет обычный отбор области по состоянию и настройке вывода.
' Сохранённые контуры записывает в собственный тестовый чертёж без нового расчёта.
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает реальные подготовленные углы экспортного контура без AutoCAD.
' Использует тот же preflight, который выполняется до изменения DWG; тест
' проверяет парсинг/адрес ошибки, но не заявляет приемку фактической записи DWG.
Public Function Audit03ReadContourArcSweepsForTests(ByVal workbook As Object) As String
    Dim sweeps As Object, key As Variant
    Set sweeps = ReadSavedContourArcSweeps(workbook)
    For Each key In sweeps.Keys
        Audit03ReadContourArcSweepsForTests = Audit03ReadContourArcSweepsForTests & CStr(key) & "|" & CStr(sweeps.Item(key)) & vbLf
    Next key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' читает те же сохранённые данные, что и экспорт в AutoCAD.
' Не создаёт приложение AutoCAD и не запускает расчёт. Полученный текст
' позволяет сравнить сохранённые значения после изменения Config.
Public Function Audit02ReadExportSnapshotForTests(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader
    Set settings = New CSystemSettingsReader
    settings.LoadFromWorkbook workbook
    Dim units As CUnitSystem
    Dim section As CSectionModel
    Dim values As Object, physicalStates As Object
    Dim combinationID As String, profileID As String, stateType As String, quantity As String
    Dim precision As Long
    Dim eps0 As Double, kx As Double, ky As Double
    Dim referenceX As Double, referenceY As Double
    Dim centroidX As Double, centroidY As Double, angle As Double
    Dim extensionUsed As Boolean, warningText As String, geometryOnly As Boolean
    ReadResultsExportState workbook, settings, units, "Concrete", section, values, physicalStates, _
        combinationID, profileID, stateType, quantity, precision, eps0, kx, ky, _
        referenceX, referenceY, centroidX, centroidY, angle, extensionUsed, warningText, geometryOnly
    If geometryOnly Then
        Audit02ReadExportSnapshotForTests = "GeometryOnly|" & warningText & "|" & CStr(section.ConcreteCount) & "|" & CStr(section.RebarCount)
        Exit Function
    End If
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' счетчик одной операции и отсутствие удержанных таблиц/книги.
' Обычный экспорт не обращается к этой диагностике и не сохраняет модель.
Public Function PerformanceExportReadDiagnosticsForTests() As String
    PerformanceExportReadDiagnosticsForTests = TestSectionSnapshotReadDiagnostics()
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет рабочие правила оформления без создания DWG. Возвращает тот же
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
    result.Add "ContourFormat", options.ContourFormat
    result.Add "ContourLayer", options.ContourLayer
    Set Audit03AutoCADPresentationForTests = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' запускает обычную очистку на подставном ModelSpace.
' Тестовые объекты отмечают вызов Delete. Проверка выявляет удаление Region
' и ошибки выбора слоёв; работу с настоящим чертежом проверяют отдельно.
Public Function Audit03CleanupAutoCADModelSpaceForTests(ByVal modelSpace As Object, _
        ByVal settings As CSystemSettingsReader) As Long
    Dim options As TAutoCADExportSettings
    options = ReadAutoCADExportSettings(settings, True)
    Audit03CleanupAutoCADModelSpaceForTests = DeleteAutoCADEntitiesOnLayers(modelSpace, AutoCADCleanupLayerSet(options))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' вызывает тот же ранний выход, что и кнопка очистки подписей.
' Подставной документ запрещает доступ к ModelSpace при отсутствии слоёв.
Public Function Audit03CleanupAutoCADDocumentForTests(ByVal document As Object, ByVal settings As CSystemSettingsReader) As Long
    Dim options As TAutoCADExportSettings, present As Boolean
    options = ReadAutoCADExportSettings(settings, True)
    Audit03CleanupAutoCADDocumentForTests = DeleteAutoCADDocumentAnnotations(document, AutoCADCleanupLayerSet(options), present)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' выполняет обычный экспорт в переданный собственный тестовый чертёж.
' Возвращает число добавленных объектов для проверки их свойств и геометрии.
Public Function Audit03ExportAutoCADDocumentForTests(ByVal doc As Object, _
        Optional ByRef geometryOnlyWarning As String = "", Optional ByRef destinationMessage As String = "") As Long
    Dim settings As CSystemSettingsReader, units As CUnitSystem, options As TAutoCADExportSettings
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Dim section As CSectionModel, values As Object, physicalStates As Object
    Dim combinationID As String, profileID As String, stateType As String, quantity As String, precision As Long
    Dim eps0 As Double, kx As Double, ky As Double, referenceX As Double, referenceY As Double
    Dim centroidX As Double, centroidY As Double, angle As Double, extended As Boolean, warning As String, geometryOnly As Boolean
    geometryOnlyWarning = vbNullString
    ReadResultsExportState ThisWorkbook, settings, units, vbNullString, section, values, physicalStates, _
        combinationID, profileID, stateType, quantity, precision, eps0, kx, ky, referenceX, referenceY, _
        centroidX, centroidY, angle, extended, warning, geometryOnly
    destinationMessage = AutoCADExportDestinationMessage(doc)
    Dim before As Long, contours As Long: before = doc.ModelSpace.Count
    If geometryOnly Then
        options = ReadGeometryExportSettings(settings, section.Contours)
        DrawFormedGeometryExport ThisWorkbook, section, options, contours, doc
        geometryOnlyWarning = GEOMETRY_ONLY_EXPORT_WARNING
    Else
        options = ReadAutoCADExportSettings(settings)
        DrawResultsStressExport section, values, physicalStates, eps0, kx, ky, referenceX, referenceY, _
            centroidX, centroidY, angle, precision, warning, options, contours, doc
    End If
    Audit03ExportAutoCADDocumentForTests = doc.ModelSpace.Count - before
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' проверяет настоящий Layer API тем же методом, который вызывает экспорт.
' Важна сохранность существующего слоя и отсутствие проглоченной ошибки Add.
Public Sub Audit03EnsureAutoCADLayerForTests(ByVal doc As Object, ByVal name As String, ByVal color As Long)
    EnsureAcadLayer doc, name, color
End Sub



' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает метаданные тем же адаптером экспорта без AutoCAD, профилей и повторного расчёта.
Public Function TestReadStateMetadataForTests(ByVal workbook As Object, ByVal combinationID As String, ByVal stateType As String) As String
    Dim eps0 As Double, kx As Double, ky As Double, rx As Double, ry As Double, cx As Double, cy As Double, angle As Double
    Dim extension As Boolean, warning As String
    ReadSectionPropertiesForCombination workbook, Nothing, combinationID, stateType, "Concrete", _
        eps0, kx, ky, rx, ry, cx, cy, angle, extension, warning
    TestReadStateMetadataForTests = CStr(extension) & "|" & warning
End Function
