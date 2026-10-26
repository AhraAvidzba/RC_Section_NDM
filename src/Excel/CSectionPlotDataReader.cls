VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionPlotDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Чтение сохранённых данных для схемы сечения
' ==========================================================================
' Собирает из Results геометрию, напряжения, деформации, свойства сечения
' и аннотации. Связывает rngNDMSectionGeometry с rngNDMElementResults по ElementID.
' Длины и кривизны переводит во внутренние мм и 1/мм для отрисовки;
' напряжения и деформации сохраняют записанные значения.
' Расчёт не запускается, текущие параметры геометрии Config не перечитываются.

Private mVisualizationQuantity As String ' Что сейчас визуализируется по профилю: Stress или Strain.
Private mLoadCase As String ' LC, выбранный для схемы после обработки Worst/конкретного ID.
Private mLoadCaseComment As String ' Комментарий LC для заголовка схемы.
Private mProfileId As String ' Имя профиля выбранного сочетания, сохранённое вместе с результатами.
Private mStateType As String ' Вид НДС для схемы, выбранный настройкой Visualization.State профиля.
Private mProfileCrackWidthEnabled As Boolean ' Показывает, запрашивал ли профиль выбранного сочетания расчёт ширины трещин.
Private mGeometryOnly As Boolean ' True для схемы без расчетных Stress/Strain.
Private mGeometryOnlyTitle As String ' Заголовок схемы без НДС, например для import-preview.
Private mGeometryOnlyNote As String ' Пояснение под схемой, если НДС намеренно отсутствует.
Private mGeometryReferenceItemsEnabled As Boolean ' Можно ли показывать оси/точку нагрузки без НДС.
Private mResultUnit As String ' Единица выбранной величины из сохраненных Results.
Private mResultCaption As String ' Название выбранной величины для легенды.
Private mResultPrecision As Long ' Количество знаков после запятой для выбранной Stress/Strain.
Private mLengthUnit As String ' Единица координат после чтения Results; в данных схемы всегда мм.
Private mAreaUnit As String ' Единица площади после чтения Results; в данных схемы всегда мм2.
Private mSnapshotLengthUnit As String ' Исходная OUTPUT-единица длины, сохраненная в Results.
Private mSnapshotAreaUnit As String ' Исходная OUTPUT-единица площади, сохраненная в Results.
Private mStressUnit As String ' Единица напряжений, сохраненная при последнем расчете.
Private mCurvatureUnit As String ' Единица кривизны после чтения Results; в данных схемы всегда 1/мм.
Private mSnapshotCurvatureUnit As String ' Исходная OUTPUT-единица кривизн, сохраненная в Results.
Private mCount As Long ' Число прочитанных элементов; в режиме НДС — только имеющих результат выбранного LC.
Private mElementCapacity As Long ' Закрытый резерв одного Load; публичный Count не включает хвост.
Private mReadTables As Object ' Временные массивы значений таблиц для одного Load; освобождаются до передачи готовых данных построителю схемы.
Private mElementID() As String ' ID элемента для связи геометрии, результатов и подписей.
Private mMaterialType() As String ' Concrete/Rebar; Plotter использует это для выбора маркера.
Private mX() As Double ' Координата X центра элемента, переведённая из Results во внутренние мм.
Private mY() As Double ' Координата Y центра элемента, переведённая из Results во внутренние мм.
Private mArea() As Double ' Площадь элемента, переведённая из Results во внутренние мм2.
Private mDiameter() As Double ' Диаметр арматуры в мм; для бетона обычно 0.
Private mShapeType() As String ' Статус интерпретации оболочки: Rectangle, Equivalent rectangle, Equivalent square.
Private mWidth() As Double ' Ширина бетонного элемента в мм, если она сохранена в Results.
Private mHeight() As Double ' Высота бетонного элемента в мм, если она сохранена в Results.
Private mRotation() As Double ' Угол прямоугольного бетонного элемента, рад; для эквивалентного квадрата может быть принят средний угол сетки.
Private mElementGeometry() As CConcreteRegion ' Явные аналитические границы Contour; прямоугольные оболочки восстанавливаются по компактным полям только при запросе.
Private mResultValue() As Double ' Значение Stress или Strain для выбранного LC.
Private mPhysicalState() As String ' Физическое состояние: сжатие, растяжение, неработающий растянутый бетон или почти нулевое напряжение.
Private mMinX As Double ' Левая граница бетонной расчётной модели в мм.
Private mMaxX As Double ' Правая граница бетонной расчётной модели в мм.
Private mMinY As Double ' Нижняя граница бетонной расчётной модели в мм.
Private mMaxY As Double ' Верхняя граница бетонной расчётной модели в мм.
Private mCentroidX As Double ' X центра тяжести приведённого сечения в мм.
Private mCentroidY As Double ' Y центра тяжести приведённого сечения в мм.
Private mPrincipalAngle As Double ' Угол от +X к главной оси 1 приведенного сечения, рад.
Private mConcreteCentroidX As Double ' X центра тяжести бетонного сечения из Results, мм.
Private mConcreteCentroidY As Double ' Y центра тяжести бетонного сечения из Results, мм.
Private mConcretePrincipalAngle As Double ' Угол главной оси 1 бетонного сечения из Results, рад.
Private mEpsilon0 As Double ' Свободный член сохраненной плоскости деформаций LC.
Private mKappaX As Double ' Коэффициент при y в плоскости деформаций, 1/мм.
Private mKappaY As Double ' Коэффициент при x в плоскости деформаций, 1/мм.
Private mLoadReferenceX As Double ' Координата X общей точки приложения нагрузки, сохранённая в Results.
Private mLoadReferenceY As Double ' Координата Y общей точки приложения нагрузки, сохранённая в Results.
Private mExtensionUsed As Boolean ' Признак использования численного продолжения диаграмм в выбранном НДС.
Private mStateStatus As String ' Статус выбранного сохранённого расчётного состояния.
Private mAnnotationCount As Long ' Число размеров, подписей и исходных контуров для схемы; расчётные области трещин сюда не включаются.
Private mAnnotationCapacity As Long ' Закрытый резерв массивов аннотаций; внешний Count содержит только заполненную часть.
Private mAnnotationType() As String ' Тип аннотации: DIMENSION, REBAR_ANNOTATION, REBAR_GLOBAL_ANNOTATION, TEXT_ANNOTATION или CONTOUR_*.
Private mLabelContours As CSectionContours, mLabelBoundary As CConcreteRegion ' Сохранённые контуры и бетонная область для отсчёта отступов подписей.
Private mAnnotationStartX() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationStartY() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationEndX() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationEndY() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationNormalX() As Double ' Компонента безразмерной нормали; для CONTOUR_ARC — координата центра дуги в мм.
Private mAnnotationNormalY() As Double ' Компонента безразмерной нормали; для CONTOUR_ARC — координата центра дуги в мм.
Private mAnnotationText() As String ' Готовый текст подписи или размера.
Private mAnnotationSweepAngle() As Double ' Проверенный signed sweep CONTOUR_ARC, рад; для остальных типов 0.
Private mUnits As CUnitSystem ' Общий перевод единиц, сохранённых с результатами; текущие единицы Config на эти значения не влияют.

' Создаёт один объект перевода единиц для всего чтения. Единицы передаются
' из Results явно, а пользовательские знаки уже применены при записи.
Private Sub Class_Initialize()
    Set mUnits = New CUnitSystem
End Sub

' Читает данные схемы только из сохранённых результатов. Из Config берёт
' выбор сочетания и оформление профиля. При ошибке удаляет частично
' прочитанные данные и возвращает исходную причину.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    On Error GoTo Failed
    Dim snapshotReadEntered As Boolean
    BeginCalculationSnapshotRead workbook: snapshotReadEntered = True
    ResetReaderState
    Set mReadTables = CreateObject("Scripting.Dictionary")
    mLoadCase = ResolveLoadCase(workbook, settings.GetRequiredString("Plot.LoadCase"))
    ApplyVisualizationProfile workbook

    ReadSectionProperties workbook
    ReadGeometryAndResults workbook
    ReadAnnotations workbook
    Set mReadTables = Nothing
    EndCalculationSnapshotRead
    snapshotReadEntered = False
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, errorSource As String
    errorNumber = Err.Number: description = Err.Description: errorSource = Err.Source
    ResetReaderState
    If snapshotReadEntered Then EndCalculationSnapshotRead
    Err.Raise errorNumber, errorSource, description
End Sub

' Загружает последнее сформированное или импортированное сечение из Results
' для визуального контроля до расчёта. Читает геометрию и аннотации без НДС
' и нейтральной линии. При ошибке не оставляет частично прочитанный снимок.
Public Sub LoadGeometryPreviewFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    On Error GoTo Failed
    ResetReaderState
    Set mReadTables = CreateObject("Scripting.Dictionary")
    mGeometryOnly = True
    mGeometryOnlyTitle = "Сформированное сечение: " & ResultsGeometrySource(workbook)
    mVisualizationQuantity = "Stress"
    mLoadCase = "Сформированное сечение"

    ReadSectionProperties workbook
    mResultCaption = "Геометрия"
    mResultUnit = "-"
    ReadGeometryPreview workbook
    ReadAnnotations workbook
    Set mReadTables = Nothing
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, errorSource As String
    errorNumber = Err.Number: description = Err.Description: errorSource = Err.Source
    ResetReaderState
    Err.Raise errorNumber, errorSource, description
End Sub

' Читает геометрию последнего расчёта, если выбранное НДС не сохранено.
' Показывает схему без раскраски и поясняет, почему напряжения или деформации
' недоступны. Расчёт заново не выполняется; повреждённые данные
' не подменяются нулевыми координатами.
Public Sub LoadGeometryOnlyForMissingState(ByVal workbook As Object, _
        ByVal settings As CSystemSettingsReader, Optional ByVal noteText As String = vbNullString)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    On Error GoTo Failed
    ResetReaderState
    Set mReadTables = CreateObject("Scripting.Dictionary")
    mLoadCase = ResolveLoadCase(workbook, settings.GetRequiredString("Plot.LoadCase"))
    ApplyVisualizationProfile workbook

    mGeometryOnly = True
    mGeometryOnlyTitle = "Геометрия расчетного сечения: " & mLoadCase
    mGeometryOnlyNote = Trim$(noteText)
    If Len(mGeometryOnlyNote) = 0 Then mGeometryOnlyNote = MissingStateMessage(mLoadCase, mStateType)
    mGeometryReferenceItemsEnabled = True

    ReadSectionProperties workbook
    mResultCaption = "Геометрия"
    mResultUnit = "-"
    ReadGeometryPreview workbook
    ReadAnnotations workbook
    Set mReadTables = Nothing
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, errorSource As String
    errorNumber = Err.Number: description = Err.Description: errorSource = Err.Source
    ResetReaderState
    Err.Raise errorNumber, errorSource, description
End Sub

' Определяет, какой LC показывать на схеме.
' При выборе Worst используется сочетание из верхней строки общей сводки
' "Определяющее сочетание". Критерий выбора задает
' General.WorstCombinationCriterion, а схема только читает сохраненный результат.
Private Function ResolveLoadCase(ByVal workbook As Object, ByVal requested As String) As String
    Dim valueText As String
    valueText = Trim$(requested)

    If StrComp(valueText, "Worst", vbTextCompare) <> 0 Then
        ResolveLoadCase = valueText
        mLoadCaseComment = ReadBatchSummaryComment(workbook, ResolveLoadCase)
        Exit Function
    End If

    ResolveLoadCase = ReadWorstCombinationID(workbook)
    If Len(ResolveLoadCase) = 0 Then ResolveLoadCase = FirstCalculatedLoadCaseFromResults(workbook)
    If Len(ResolveLoadCase) = 0 Then ResolveLoadCase = FirstLoadCaseFromSectionProperties(workbook)
    mLoadCaseComment = ReadBatchSummaryComment(workbook, ResolveLoadCase)
    If Len(ResolveLoadCase) = 0 Then Err.Raise vbObjectError + 4702, "CSectionPlotDataReader", _
        "В сохранённых Results нет определяющего или другого рассчитанного сочетания для Plot.LoadCase = Worst. Сначала выполните расчёт."
End Function

' Читает поле «Определяющее сочетание» над компактной сводкой Results.
Private Function ReadWorstCombinationID(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    ReadWorstCombinationID = Trim$(SafeText(anchor.Value2))
Failed:
End Function

' Применяет оформление профиля выбранного сочетания. ProfileId берёт
' из сохранённых Results, чтобы показывать фактически рассчитанный запуск
' независимо от последующих правок таблицы нагрузок.
Private Sub ApplyVisualizationProfile(ByVal workbook As Object)
    mProfileId = ReadProfileIdForLoadCase(workbook, mLoadCase)
    If Len(mProfileId) = 0 Then Err.Raise vbObjectError + 4710, "CSectionPlotDataReader", _
        "В Results не найден ProfileId для сочетания: " & mLoadCase

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook

    Dim profile As CCalculationProfile
    If profiles.HasProfile(mProfileId) Then
        Set profile = profiles.ProfileById(mProfileId)
    Else
        Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
        If snapshot Is Nothing Then Set profile = profiles.ProfileById(mProfileId) Else Set profile = snapshot.SavedVisualizationProfile(mProfileId)
    End If
    mProfileCrackWidthEnabled = profile.CrackWidthEnabled
    mStateType = profile.VisualizationStateText
    mVisualizationQuantity = VisualizationQuantityToText(profile.VisualizationQuantity)
    mResultPrecision = profile.VisualizationPrecisionForQuantity(profile.VisualizationQuantity)
End Sub

' Находит сохранённый профиль сочетания прежде всего в результатах элементов.
' Если нужных строк нет, проверяет свойства сечения
' для более точного пояснения отсутствующего НДС.
Private Function ReadProfileIdForLoadCase(ByVal workbook As Object, ByVal loadCase As String) As String
    ReadProfileIdForLoadCase = ReadProfileIdFromElementResults(workbook, loadCase)
    If Len(ReadProfileIdForLoadCase) = 0 Then
        ReadProfileIdForLoadCase = ReadLoadCasePropertyText(workbook, loadCase, "ProfileId")
    End If
End Function

' Ищет профиль сочетания в таблице результатов элементов,
' где каждое сохранённое НДС связано со своим расчётным профилем.
Private Function ReadProfileIdFromElementResults(ByVal workbook As Object, ByVal loadCase As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadElementResultTable(workbook, loadCase, , True)
    If Not HasRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colProfileId As Long: colProfileId = ResultColumn(data, "ProfileId")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colLoadCase))), loadCase, vbTextCompare) = 0 Then
            ReadProfileIdFromElementResults = Trim$(SafeText(data(rowIndex, colProfileId)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Ищет текстовое свойство сочетания в rngNDMSectionProperties,
' если для нужного НДС не сохранились строки результатов элементов.
Private Function ReadLoadCasePropertyText(ByVal workbook As Object, ByVal loadCase As String, ByVal propertyName As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasRows(data) Then Exit Function

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, colLoadCase))), loadCase, vbTextCompare) = 0 And _
                StrComp(Trim$(SafeText(data(rowIndex, colParameter))), propertyName, vbTextCompare) = 0 Then
            ReadLoadCasePropertyText = Trim$(SafeText(data(rowIndex, colValue)))
            Exit Function
        End If
    Next rowIndex
Failed:
End Function

' Возвращает первое сочетание с результатами элементов, если Worst нельзя
' получить из запаса прочности, например при одном НДС без поиска предела.
' Предварительный просмотр сформированного сечения этим способом не выбирается.
Private Function FirstCalculatedLoadCaseFromResults(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadElementResultTable(workbook, , , True)
    If Not HasRows(data) Then Exit Function

    Dim colLoadCase As Long
    colLoadCase = ResultColumn(data, "LoadCase")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        FirstCalculatedLoadCaseFromResults = Trim$(SafeText(data(rowIndex, colLoadCase)))
        If Len(FirstCalculatedLoadCaseFromResults) > 0 Then Exit Function
    Next rowIndex
Failed:
End Function

' Возвращает первое сочетание из свойств сечения. Даже при отсутствии
' НДС для раскраски это позволяет показать, какие результаты сохранены.
Private Function FirstLoadCaseFromSectionProperties(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasRows(data) Then Exit Function

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

' Находит комментарий выбранного LC для заголовка схемы.
' Если комментария нет или общая сводка ещё пуста, возвращается пустая строка
' без ошибки: сама схема может строиться и без комментария.
Private Function ReadBatchSummaryComment(ByVal workbook As Object, ByVal loadCase As String) As String
    On Error GoTo Failed
    If Len(Trim$(loadCase)) = 0 Then Exit Function

    Dim snapshot As CCalculationSnapshot: Set snapshot = CurrentCalculationSnapshot(workbook)
    If Not snapshot Is Nothing Then
        Dim summary As Variant, summaryRow As Long: summary = snapshot.Table("Summary")
        For summaryRow = 13 To UBound(summary, 1)
            If StrComp(Trim$(SafeText(summary(summaryRow, 1))), loadCase, vbTextCompare) = 0 Then
                ReadBatchSummaryComment = Trim$(SafeText(summary(summaryRow, 2))): Exit Function
            End If
        Next summaryRow
        Exit Function
    End If

    Dim anchor As Object, nextAnchor As Object, summaryWriter As CStrengthSummaryWriter
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    Dim firstRow As Long, lastRow As Long, summaryEnd As Long, row As Long, data As Variant, combinationID As String
    firstRow = anchor.Row + 12 ' CBatchResultWriter.FIRST_DATA_ROW=13; строка worst LC перед ней не является пользовательским сочетанием.
    lastRow = anchor.Worksheet.Cells(anchor.Worksheet.Rows.Count, anchor.Column).End(xlUp).Row
    On Error Resume Next
    Set nextAnchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange(1, 1)
    On Error GoTo Failed
    If Not nextAnchor Is Nothing Then
        If nextAnchor.Worksheet Is anchor.Worksheet Then
            Set summaryWriter = New CStrengthSummaryWriter
            summaryEnd = nextAnchor.Row - summaryWriter.HeaderRowsAboveAnchor - 1 ' Не читаем соседнюю таблицу прочности вместо отсутствующего комментария сводки.
            If summaryEnd < firstRow Then Exit Function
            If summaryEnd < lastRow Then lastRow = summaryEnd
        End If
    End If
    If lastRow < firstRow Then Exit Function
    data = anchor.Worksheet.Cells(firstRow, anchor.Column).Resize(lastRow - firstRow + 1, 2).Value2
    For row = 1 To UBound(data, 1)
        combinationID = Trim$(SafeText(data(row, 1)))
        If StrComp(combinationID, loadCase, vbTextCompare) = 0 Then
            ReadBatchSummaryComment = Trim$(SafeText(data(row, 2)))
            Exit Function
        End If
    Next row
Failed:
End Function

' Читает свойства сечения и выбранного сочетания: габариты, центр тяжести,
' главные оси, единицы и параметры поля деформаций epsilon0/kappaX/kappaY.
Private Sub ReadSectionProperties(ByVal workbook As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionProperties").RefersToRange
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasRows(data) Then Err.Raise vbObjectError + 4703, "CSectionPlotDataReader", _
        "На листе Results нет rngNDMSectionProperties. Сначала выполните расчет."

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")

    Dim rowIndex As Long
    Dim parameter As String, loadCase As String
    Dim stateFields As Object: Set stateFields = CreateObject("Scripting.Dictionary")
    For rowIndex = 2 To UBound(data, 1) ' Сначала читаем единицы результата, чтобы порядок строк свойств не влиял на перевод величин без отдельно указанной единицы.
        parameter = Trim$(ReadSnapshotText(data(rowIndex, colParameter), anchor, rowIndex, colParameter, "Parameter"))
        loadCase = Trim$(ReadSnapshotText(data(rowIndex, colLoadCase), anchor, rowIndex, colLoadCase, "LoadCase"))
        If StrComp(loadCase, "ALL", vbTextCompare) = 0 Or Len(loadCase) = 0 Then
            Select Case LCase$(parameter)
                Case "output.lengthunit", "output.areaunit"
                    ReadCommonProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
                Case "output.stressunit", "output.curvatureunit"
                    If Not mGeometryOnly Or mGeometryReferenceItemsEnabled Then _
                        ReadCommonProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
            End Select
        End If
    Next rowIndex
    For rowIndex = 2 To UBound(data, 1) ' Берём общие строки ALL и строки выбранного сочетания mLoadCase. Результаты других сочетаний в эту схему не включаем.
        parameter = Trim$(ReadSnapshotText(data(rowIndex, colParameter), anchor, rowIndex, colParameter, "Parameter"))
        loadCase = Trim$(ReadSnapshotText(data(rowIndex, colLoadCase), anchor, rowIndex, colLoadCase, "LoadCase"))
        If Len(parameter) > 0 Then
            If StrComp(loadCase, "ALL", vbTextCompare) = 0 Or Len(loadCase) = 0 Then
                If Left$(LCase$(parameter), 7) <> "output." Then _
                    ReadCommonProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
            ElseIf StrComp(loadCase, mLoadCase, vbTextCompare) = 0 Then
                If Not mGeometryOnly Or mGeometryReferenceItemsEnabled Then
                    ReadLoadCaseProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
                    stateFields(LCase$(parameter)) = True
                End If
            End If
        End If
    Next rowIndex

    If Not mGeometryOnly Then
        Dim field As Variant, requiredKey As String
        For Each field In Array("Epsilon0", "KappaX", "KappaY", "ExtensionUsed", "Status")
            requiredKey = "State." & mStateType & "." & CStr(field)
            If Not stateFields.Exists(LCase$(requiredKey)) Then Err.Raise vbObjectError + 4703, "CSectionPlotDataReader", _
                "В сохраненном снимке Results для сочетания " & mLoadCase & " отсутствует обязательное поле " & requiredKey & _
                " (rngNDMSectionProperties, " & anchor.Worksheet.Name & "!" & anchor.Address(False, False) & "). Повторите расчет/импорт для восстановления снимка."
        Next field
    End If

    If Len(mSnapshotLengthUnit) = 0 Then mSnapshotLengthUnit = "mm"
    If Len(mSnapshotAreaUnit) = 0 Then mSnapshotAreaUnit = mSnapshotLengthUnit & "2"
    If Len(mStressUnit) = 0 Then mStressUnit = "MPa"
    If Len(mSnapshotCurvatureUnit) = 0 Then mSnapshotCurvatureUnit = "1/mm"
    mLengthUnit = "mm"
    mAreaUnit = "mm2"
    mCurvatureUnit = "1/mm"
    If StrComp(mVisualizationQuantity, "Strain", vbTextCompare) = 0 Then
        mResultUnit = "-"
        mResultCaption = "Strain"
    Else
        mResultUnit = mStressUnit
        mResultCaption = "Stress"
    End If
End Sub

' Читает общие свойства ALL из Results независимо от способа формирования:
' габариты и центры переводит в мм, углы оставляет в радианах.
' Здесь же читается единая точка приложения нагрузки сохранённого запуска.
Private Sub ReadCommonProperty(ByVal parameter As String, ByVal value As Variant, ByVal unitValue As Variant, _
        ByVal anchor As Object, ByVal row As Long, ByVal column As Long, ByVal unitColumn As Long)
    Select Case LCase$(parameter)
        Case "bounds.minx", "bounds.maxx", "bounds.miny", "bounds.maxy", _
                "concrete.centroidx", "concrete.centroidy", "concrete.principalangle", _
                "transformed.centroidx", "transformed.centroidy", "transformed.principalangle", _
                "loadreferencex", "loadreferencey"
            value = ReadSnapshotNumber(value, anchor, row, column, parameter)
    End Select
    Dim unitText As String
    Select Case LCase$(parameter)
        Case "bounds.minx", "bounds.maxx", "bounds.miny", "bounds.maxy", _
                "concrete.centroidx", "concrete.centroidy", "transformed.centroidx", "transformed.centroidy", _
                "loadreferencex", "loadreferencey"
            unitText = ReadSnapshotUnit(unitValue, "Length", mSnapshotLengthUnit, "mm", anchor, row, unitColumn, parameter & ".Unit", True)
    End Select
    Select Case LCase$(parameter)
        Case "bounds.minx": mMinX = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "bounds.maxx": mMaxX = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "bounds.miny": mMinY = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "bounds.maxy": mMaxY = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "concrete.centroidx": mConcreteCentroidX = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "concrete.centroidy": mConcreteCentroidY = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "concrete.principalangle": mConcretePrincipalAngle = CDbl(value)
        Case "transformed.centroidx": mCentroidX = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "transformed.centroidy": mCentroidY = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "transformed.principalangle": mPrincipalAngle = CDbl(value)
        Case "loadreferencex": mLoadReferenceX = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "loadreferencey": mLoadReferenceY = OutputLengthToInternalByUnit(CDbl(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "output.lengthunit": mSnapshotLengthUnit = ReadSnapshotUnit(value, "Length", vbNullString, vbNullString, anchor, row, column, parameter)
        Case "output.areaunit": mSnapshotAreaUnit = ReadSnapshotUnit(value, "Area", vbNullString, vbNullString, anchor, row, column, parameter)
        Case "output.stressunit": mStressUnit = ReadSnapshotUnit(value, "Stress", vbNullString, vbNullString, anchor, row, column, parameter)
        Case "output.curvatureunit": mSnapshotCurvatureUnit = ReadSnapshotUnit(value, "Curvature", vbNullString, vbNullString, anchor, row, column, parameter)
    End Select
End Sub

' Читает поле деформаций и признаки допустимости только выбранного НДС
' из свойств State.<StateType>.*; данные других состояний не подмешиваются.
Private Sub ReadLoadCaseProperty(ByVal parameter As String, ByVal value As Variant, ByVal unitValue As Variant, _
        ByVal anchor As Object, ByVal row As Long, ByVal column As Long, ByVal unitColumn As Long)
    Dim key As String
    Dim prefix As String
    key = LCase$(Trim$(parameter))
    prefix = "state." & LCase$(mStateType) & "."
    If Left$(key, Len(prefix)) <> prefix Then Exit Sub

    Dim unitText As String
    If Mid$(key, Len(prefix) + 1) = "kappax" Or Mid$(key, Len(prefix) + 1) = "kappay" Then _
        unitText = ReadSnapshotUnit(unitValue, "Curvature", mSnapshotCurvatureUnit, "1/mm", anchor, row, unitColumn, parameter & ".Unit", True)
    Select Case Mid$(key, Len(prefix) + 1)
        Case "epsilon0": mEpsilon0 = ReadSnapshotNumber(value, anchor, row, column, parameter)
        Case "kappax": mKappaX = OutputCurvatureToInternalByUnit(ReadSnapshotNumber(value, anchor, row, column, parameter), UnitOrDefault(unitText, mSnapshotCurvatureUnit, "1/mm"))
        Case "kappay": mKappaY = OutputCurvatureToInternalByUnit(ReadSnapshotNumber(value, anchor, row, column, parameter), UnitOrDefault(unitText, mSnapshotCurvatureUnit, "1/mm"))
        Case "extensionused": mExtensionUsed = ReadSnapshotBoolean(value, anchor, row, column, parameter)
        Case "status": mStateStatus = ReadSnapshotStatus(value, anchor, row, column, parameter)
    End Select
End Sub

' Соединяет постоянную геометрию с LC-зависимыми результатами.
' rngNDMSectionGeometry хранит координаты и площади один раз, а
' rngNDMElementResults хранит Stress/Strain/PhysicalState по LoadCase.
Private Sub ReadGeometryAndResults(ByVal workbook As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    Dim boundaries As Object: Set boundaries = ReadSavedConcreteBoundaries(workbook, geom)
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."

    Dim results As Variant
    Dim resultAnchor As Object
    Set resultAnchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    results = ReadElementResultTable(workbook, mLoadCase, mStateType)
    If Not HasRows(results) Then results = ReadElementResultTable(workbook, , , True) ' Индекс отличает отсутствие выбранного состояния от отсутствия всех результатов без копирования полной таблицы.
    If Not HasRows(results) Then Err.Raise vbObjectError + 4705, "CSectionPlotDataReader", _
        "На листе Results нет результатов элементов."

    Dim resultMap As Object
    Set resultMap = CreateObject("Scripting.Dictionary")
    resultMap.CompareMode = vbTextCompare

    Dim colResLC As Long: colResLC = ResultColumn(results, "LoadCase")
    Dim colResStateType As Long: colResStateType = ResultColumn(results, "StateType")
    Dim colResID As Long: colResID = ResultColumn(results, "ElementID")
    Dim colStrain As Long: colStrain = ResultColumn(results, "Strain")
    Dim colStress As Long: colStress = ResultColumn(results, "Stress")
    Dim colState As Long: colState = ResultColumn(results, "PhysicalState")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(results, 1) ' Собирает напряжения и деформации выбранного сочетания по ElementID. Они хранятся отдельно от геометрии и затем связываются с нужными элементами.
        If StrComp(Trim$(SafeText(results(rowIndex, colResLC))), mLoadCase, vbTextCompare) = 0 And _
                StrComp(Trim$(SafeText(results(rowIndex, colResStateType))), mStateType, vbTextCompare) = 0 Then
            resultMap(CStr(results(rowIndex, colResID))) = Array( _
                ReadSnapshotNumber(results(rowIndex, colStrain), resultAnchor, rowIndex, colStrain, CStr(results(1, colStrain))), _
                ReadSnapshotNumber(results(rowIndex, colStress), resultAnchor, rowIndex, colStress, CStr(results(1, colStress))), _
                CStr(results(rowIndex, colState)))
        End If
    Next rowIndex
    If resultMap.Count = 0 Then Err.Raise vbObjectError + 4706, "CSectionPlotDataReader", _
        MissingStateMessage(mLoadCase, mStateType)

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    If UBound(geom, 2) < 15 Then Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", _
        "В rngNDMSectionGeometry не хватает столбцов геометрии. Заново сформируйте сечение."
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colShape As Long: colShape = ResultColumn(geom, "GeometryInterpretationStatus")
    Dim colDiameter As Long: colDiameter = ResultColumn(geom, "Diameter")
    Dim colWidth As Long: colWidth = ResultColumn(geom, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(geom, "Height")
    Dim colRotation As Long: colRotation = ResultColumn(geom, "Rotation")
    Dim lengthUnit As String: lengthUnit = ResultHeaderUnit(geom, colX, mSnapshotLengthUnit)
    Dim areaUnit As String: areaUnit = ResultHeaderUnit(geom, colArea, mSnapshotAreaUnit)

    For rowIndex = 2 To UBound(geom, 1) ' Соединяет сохранённую геометрию с прочитанным НДС по элементам. Построитель схемы получает готовые данные без нового чтения Results.
        Dim elementID As String
        elementID = Trim$(CStr(geom(rowIndex, colID)))
        If Len(elementID) > 0 And resultMap.Exists(elementID) And _
                (StrComp(CStr(geom(rowIndex, colType)), "Concrete", vbTextCompare) = 0 Or _
                StrComp(CStr(geom(rowIndex, colType)), "Rebar", vbTextCompare) = 0) Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colX), lengthUnit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colY), lengthUnit), OutputAreaToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colArea), areaUnit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colDiameter, StrComp(CStr(geom(rowIndex, colType)), "Concrete", vbTextCompare) = 0), lengthUnit), CStr(geom(rowIndex, colShape)), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colWidth, True), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colHeight, True), lengthUnit), _
                ReadGeometryNumber(geom, anchor, rowIndex, colRotation, True), resultMap(elementID)
            If boundaries.Exists(elementID) Then Set mElementGeometry(mCount) = boundaries.Item(elementID)
        End If
    Next rowIndex
End Sub

' Читает геометрию без расчётной раскраски для предварительного просмотра
' импорта или случая, когда выбранное НДС отсутствует,
' а сами элементы уже сохранены в Results.
Private Sub ReadGeometryPreview(ByVal workbook As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    Dim boundaries As Object: Set boundaries = ReadSavedConcreteBoundaries(workbook, geom)
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."
    If UBound(geom, 2) < 15 Then Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", _
        "В rngNDMSectionGeometry не хватает столбцов геометрии. Заново сформируйте сечение."

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colShape As Long: colShape = ResultColumn(geom, "GeometryInterpretationStatus")
    Dim colDiameter As Long: colDiameter = ResultColumn(geom, "Diameter")
    Dim colWidth As Long: colWidth = ResultColumn(geom, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(geom, "Height")
    Dim colRotation As Long: colRotation = ResultColumn(geom, "Rotation")
    Dim lengthUnit As String: lengthUnit = ResultHeaderUnit(geom, colX, mSnapshotLengthUnit)
    Dim areaUnit As String: areaUnit = ResultHeaderUnit(geom, colArea, mSnapshotAreaUnit)

    Dim neutralResult As Variant
    neutralResult = Array(0#, 0#, "NearZero") ' Предварительная схема не утверждает нулевое НДС: это только нейтральная окраска.

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(geom, 1)
        Dim elementID As String
        elementID = Trim$(CStr(geom(rowIndex, colID)))
        If Len(elementID) > 0 And (StrComp(CStr(geom(rowIndex, colType)), "Concrete", vbTextCompare) = 0 Or _
                StrComp(CStr(geom(rowIndex, colType)), "Rebar", vbTextCompare) = 0) Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colX), lengthUnit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colY), lengthUnit), OutputAreaToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colArea), areaUnit), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colDiameter, StrComp(CStr(geom(rowIndex, colType)), "Concrete", vbTextCompare) = 0), lengthUnit), CStr(geom(rowIndex, colShape)), _
                OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colWidth, True), lengthUnit), OutputLengthToInternalByUnit(ReadGeometryNumber(geom, anchor, rowIndex, colHeight, True), lengthUnit), _
                ReadGeometryNumber(geom, anchor, rowIndex, colRotation, True), neutralResult
            If boundaries.Exists(elementID) Then Set mElementGeometry(mCount) = boundaries.Item(elementID)
        End If
    Next rowIndex
End Sub

' Проверяет координаты и размеры элемента до добавления на схему.
' Отсутствующая необязательная графическая граница допустима,
' но повреждённые обязательные ячейки не подменяются нулями.
' Ошибка указывает фактическую ячейку Results.
Private Function ReadGeometryNumber(ByRef data As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, Optional ByVal allowBlank As Boolean = False) As Double
    ReadGeometryNumber = ReadSnapshotNumber(data(row, column), anchor, row, column, _
        ResultHeaderBase(CStr(data(1, column))), allowBlank)
End Function

' Читает число геометрии или свойства снимка до любого CDbl/пересчета единиц.
' Формульная ошибка, Boolean и нечисловой текст не становятся нулем; сообщение
' называет фактическую ячейку и действие для восстановления сохраненных данных.
Private Function ReadSnapshotNumber(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String, _
        Optional ByVal allowBlank As Boolean = False) As Double
    If IsError(value) Or IsNull(value) Then GoTo InvalidNumber
    If VarType(value) = vbBoolean Or VarType(value) = vbDate Then GoTo InvalidNumber
    If Len(Trim$(CStr(value))) = 0 Then
        If allowBlank Then Exit Function
        GoTo InvalidNumber
    End If
    Dim parsed As Double
    If Not TryReadStrictDouble(value, parsed) Then GoTo InvalidNumber
    On Error GoTo InvalidNumber
    ReadSnapshotNumber = parsed
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4719, "CSectionPlotDataReader", _
        "В сохраненном снимке Results поле " & field & _
        " содержит нечисловое значение: " & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Повторите импорт геометрии или расчет, чтобы восстановить корректный снимок."
End Function

' Добавляет элемент в подготовленные данные схемы. Для раскраски берёт
' напряжение или деформацию согласно текущему виду отображения.
Private Sub AppendElement(ByVal elementID As String, ByVal materialType As String, ByVal xCoord As Double, _
        ByVal yCoord As Double, ByVal area As Double, ByVal diameter As Double, _
        ByVal shapeType As String, ByVal width As Double, ByVal height As Double, _
        ByVal rotation As Double, ByVal resultData As Variant)
    EnsureElementCapacity mCount + 1 ' Резерв растёт сразу для всех полей элемента; предварительная схема передаёт нейтральный resultData.
    mCount = mCount + 1

    mElementID(mCount) = elementID
    mMaterialType(mCount) = materialType
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mArea(mCount) = area
    mDiameter(mCount) = diameter
    mShapeType(mCount) = shapeType
    mWidth(mCount) = width
    mHeight(mCount) = height
    mRotation(mCount) = rotation
    If StrComp(mVisualizationQuantity, "Strain", vbTextCompare) = 0 Then
        mResultValue(mCount) = CDbl(resultData(0))
    Else
        mResultValue(mCount) = CDbl(resultData(1))
    End If
    mPhysicalState(mCount) = CStr(resultData(2))
End Sub

' Увеличивает вместимость массивов с запасом, сохраняя порядок элементов.
' Чтение допускается только в заполненной части; ResetReaderState очищает также резерв.
Private Sub EnsureElementCapacity(ByVal required As Long)
    If required <= mElementCapacity Then Exit Sub
    Dim capacity As Long
    capacity = required
    If mElementCapacity = 0 Then
        If capacity < 64 Then capacity = 64
    ElseIf mElementCapacity <= 1073741823 Then
        If capacity < mElementCapacity * 2 Then capacity = mElementCapacity * 2
    End If
    ReDim Preserve mElementID(1 To capacity)
    ReDim Preserve mMaterialType(1 To capacity)
    ReDim Preserve mX(1 To capacity)
    ReDim Preserve mY(1 To capacity)
    ReDim Preserve mArea(1 To capacity)
    ReDim Preserve mDiameter(1 To capacity)
    ReDim Preserve mShapeType(1 To capacity)
    ReDim Preserve mWidth(1 To capacity)
    ReDim Preserve mHeight(1 To capacity)
    ReDim Preserve mRotation(1 To capacity)
    ReDim Preserve mElementGeometry(1 To capacity)
    ReDim Preserve mResultValue(1 To capacity)
    ReDim Preserve mPhysicalState(1 To capacity)
    mElementCapacity = capacity
End Sub

' Читает сохранённые размеры, подписи и другие аннотации схемы.
' Если таблица отсутствует или пуста, схема всё равно строится:
' это допустимо при отсутствии необязательного оформления; исходные контуры читаются отдельно.
Private Sub ReadAnnotations(ByVal workbook As Object)
    ReadMaterialContours workbook
    Dim tableName As Object, anchor As Object
    On Error Resume Next ' Отсутствие необязательного Name отличается от повреждения его данных.
    Set tableName = workbook.Names.Item("rngNDMSectionAnnotations")
    On Error GoTo 0
    If tableName Is Nothing Then Exit Sub
    Set anchor = tableName.RefersToRange
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasRows(data) Then Exit Sub

    Dim colType As Long: colType = ResultColumn(data, "AnnotationType")
    Dim colStartX As Long: colStartX = ResultColumn(data, "StartX")
    Dim colStartY As Long: colStartY = ResultColumn(data, "StartY")
    Dim colEndX As Long: colEndX = ResultColumn(data, "EndX")
    Dim colEndY As Long: colEndY = ResultColumn(data, "EndY")
    Dim colNormalX As Long: colNormalX = ResultColumn(data, "OutsideNormalX")
    Dim colNormalY As Long: colNormalY = ResultColumn(data, "OutsideNormalY")
    Dim colText As Long: colText = ResultColumn(data, "Text")
    Dim colCoordinateUnit As Long: colCoordinateUnit = ResultColumn(data, "CoordinateUnit")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim annotationType As String
        annotationType = Trim$(ReadAnnotationText(data(rowIndex, colType), anchor, rowIndex, colType, "AnnotationType"))
        If IsPlotLabelAnnotation(annotationType) Then
            Dim lengthUnit As String
            lengthUnit = ReadSnapshotUnit(data(rowIndex, colCoordinateUnit), "Length", vbNullString, vbNullString, _
                anchor, rowIndex, colCoordinateUnit, "CoordinateUnit")
            Dim startX As Double, startY As Double, endX As Double, endY As Double
            Dim normalX As Double, normalY As Double, textValue As String
            startX = ReadAnnotationNumber(data(rowIndex, colStartX), anchor, rowIndex, colStartX, "StartX")
            startY = ReadAnnotationNumber(data(rowIndex, colStartY), anchor, rowIndex, colStartY, "StartY")
            endX = ReadAnnotationNumber(data(rowIndex, colEndX), anchor, rowIndex, colEndX, "EndX")
            endY = ReadAnnotationNumber(data(rowIndex, colEndY), anchor, rowIndex, colEndY, "EndY")
            normalX = ReadAnnotationNumber(data(rowIndex, colNormalX), anchor, rowIndex, colNormalX, "OutsideNormalX")
            normalY = ReadAnnotationNumber(data(rowIndex, colNormalY), anchor, rowIndex, colNormalY, "OutsideNormalY")
            textValue = ReadAnnotationText(data(rowIndex, colText), anchor, rowIndex, colText, "Text")
            EnsureAnnotationCapacity mAnnotationCount + 1
            mAnnotationCount = mAnnotationCount + 1
            mAnnotationType(mAnnotationCount) = annotationType
            mAnnotationStartX(mAnnotationCount) = OutputLengthToInternalByUnit(startX, lengthUnit)
            mAnnotationStartY(mAnnotationCount) = OutputLengthToInternalByUnit(startY, lengthUnit)
            mAnnotationEndX(mAnnotationCount) = OutputLengthToInternalByUnit(endX, lengthUnit)
            mAnnotationEndY(mAnnotationCount) = OutputLengthToInternalByUnit(endY, lengthUnit)
            mAnnotationNormalX(mAnnotationCount) = normalX
            mAnnotationNormalY(mAnnotationCount) = normalY
            mAnnotationText(mAnnotationCount) = textValue
            mAnnotationSweepAngle(mAnnotationCount) = 0#
        End If
    Next rowIndex
End Sub

' Выбирает только размерные подписи и подписи арматуры для Excel-схемы.
' Области трещин и их пробные точки имеют другую схему полей: отсутствие
' нормали или координат неуспешной пробы не является повреждением подписи.
Private Function IsPlotLabelAnnotation(ByVal annotationType As String) As Boolean
    Select Case UCase$(Trim$(annotationType))
        Case "DIMENSION", "REBAR_ANNOTATION", "REBAR_GLOBAL_ANNOTATION", "MEASUREMENT_LABEL", "TEXT_ANNOTATION"
            IsPlotLabelAnnotation = True
    End Select
End Function

' Подготавливает сохранённые наружные контуры и отверстия для отрисовки.
' Это представление тех же границ, а не их повторная запись в расчётные аннотации.
Private Sub ReadMaterialContours(ByVal workbook As Object)
    Dim contours As CSectionContours, i As Long
    Set contours = ReadSavedSectionContours(workbook)
    Set mLabelContours = contours
    EnsureAnnotationCapacity mAnnotationCount + contours.Count ' Число исходных сегментов уже известно; резервируем их одной операцией до прохода.
    For i = 1 To contours.Count
        mAnnotationCount = mAnnotationCount + 1
        mAnnotationType(mAnnotationCount) = contours.SegmentType(i)
        mAnnotationStartX(mAnnotationCount) = contours.StartX(i)
        mAnnotationStartY(mAnnotationCount) = contours.StartY(i)
        mAnnotationEndX(mAnnotationCount) = contours.EndX(i)
        mAnnotationEndY(mAnnotationCount) = contours.EndY(i)
        mAnnotationNormalX(mAnnotationCount) = contours.CenterX(i) ' Для дуги поля NormalX/NormalY несут центр в мм, а не направление выноски.
        mAnnotationNormalY(mAnnotationCount) = contours.CenterY(i)
        mAnnotationSweepAngle(mAnnotationCount) = contours.SweepAngle(i)
    Next i
    Dim edge As TRegionEdge
    For i = 1 To contours.Count ' Границы представления включают наружные линии и дуги без перестройки расчётной области. Нулевой sweep остаётся прямым отрезком, как в поддерживаемом формате отрисовки.
        If StrComp(contours.LoopRole(i), "Outer", vbTextCompare) = 0 Then
            edge.X1 = contours.StartX(i): edge.Y1 = contours.StartY(i)
            edge.X2 = contours.EndX(i): edge.Y2 = contours.EndY(i)
            edge.CenterX = contours.CenterX(i): edge.CenterY = contours.CenterY(i): edge.Sweep = contours.SweepAngle(i)
            If contours.SegmentType(i) = "CONTOUR_CIRCLE" Then ' Окружность разворачивается в полную дугу для точного габарита, без дискретизации.
                edge.X1 = edge.CenterX + contours.Radius(i): edge.Y1 = edge.CenterY
                edge.X2 = edge.X1: edge.Y2 = edge.Y1: edge.Sweep = 2# * GEOM_PI
            End If
            modGeometryPrimitives.AccumulateEdgeProjection edge, 1#, 0#, mMinX, mMaxX
            modGeometryPrimitives.AccumulateEdgeProjection edge, 0#, 1#, mMinY, mMaxY
        End If
    Next i
End Sub

' Увеличивает закрытый резерв аннотаций геометрически; копирование массивов
' больше не повторяется для каждого сегмента импортированного контура.
Private Sub EnsureAnnotationCapacity(ByVal required As Long)
    If required <= mAnnotationCapacity Then Exit Sub
    Dim capacity As Long
    capacity = required
    If mAnnotationCapacity = 0 Then
        If capacity < 64 Then capacity = 64
    ElseIf mAnnotationCapacity <= 1073741823 Then
        If capacity < mAnnotationCapacity * 2 Then capacity = mAnnotationCapacity * 2
    End If
    ReDim Preserve mAnnotationType(1 To capacity)
    ReDim Preserve mAnnotationStartX(1 To capacity)
    ReDim Preserve mAnnotationStartY(1 To capacity)
    ReDim Preserve mAnnotationEndX(1 To capacity)
    ReDim Preserve mAnnotationEndY(1 To capacity)
    ReDim Preserve mAnnotationNormalX(1 To capacity)
    ReDim Preserve mAnnotationNormalY(1 To capacity)
    ReDim Preserve mAnnotationText(1 To capacity)
    ReDim Preserve mAnnotationSweepAngle(1 To capacity)
    mAnnotationCapacity = capacity
End Sub

' Читает координату или направление аннотации. Ноль допустим; пустое
' или нечисловое обязательное поле вызывает ошибку с адресом Results
' вместо подстановки нулевого значения.
Private Function ReadAnnotationNumber(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As Double
    If IsError(value) Then GoTo InvalidValue
    If IsNull(value) Or IsEmpty(value) Then GoTo InvalidValue
    Dim parsed As Double
    If Not TryReadStrictDouble(value, parsed) Then GoTo InvalidValue
    On Error GoTo InvalidValue
    ReadAnnotationNumber = parsed
    Exit Function
InvalidValue:
    Err.Raise vbObjectError + 4718, "CSectionPlotDataReader", AnnotationInputMessage(anchor, row, column, field)
End Function

' Пустой текст допустим для линии или контура, но ошибка Excel/Null не является
' пустой подписью. Аннотация с такой ошибкой не выдается как успешно прочитанная.
Private Function ReadAnnotationText(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    If IsError(value) Or IsNull(value) Then Err.Raise vbObjectError + 4718, _
        "CSectionPlotDataReader", AnnotationInputMessage(anchor, row, column, field)
    ReadAnnotationText = CStr(value)
End Function

' Возвращает фактический адрес ячейки через текущее именованное начало таблицы.
' Перемещение блока Results не оставляет в ошибке прежний адрес.
Private Function AnnotationInputMessage(ByVal anchor As Object, ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    AnnotationInputMessage = "В сохраненной аннотации Results поле " & field & " (" & _
        anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ") содержит недопустимое значение. Исправьте данные или повторите расчет/импорт для восстановления снимка."
End Function

' Очищает массивы перед новым чтением, чтобы элементы и аннотации
' предыдущей загрузки не попали в новую схему.
Private Sub ResetReaderState()
    Set mReadTables = Nothing
    mElementCapacity = 0
    mGeometryOnly = False
    mGeometryOnlyTitle = vbNullString
    mGeometryOnlyNote = vbNullString
    mGeometryReferenceItemsEnabled = False
    mVisualizationQuantity = vbNullString
    mLoadCase = vbNullString
    mLoadCaseComment = vbNullString
    mProfileId = vbNullString
    mStateType = vbNullString
    mProfileCrackWidthEnabled = False
    mResultUnit = vbNullString
    mResultCaption = vbNullString
    mResultPrecision = 1
    mLengthUnit = vbNullString
    mAreaUnit = vbNullString
    mSnapshotLengthUnit = vbNullString
    mSnapshotAreaUnit = vbNullString
    mStressUnit = vbNullString
    mCurvatureUnit = vbNullString
    mSnapshotCurvatureUnit = vbNullString
    mCount = 0
    mAnnotationCount = 0
    mAnnotationCapacity = 0
    Set mLabelContours = Nothing: Set mLabelBoundary = Nothing
    mMinX = 0#: mMaxX = 0#: mMinY = 0#: mMaxY = 0#
    mCentroidX = 0#: mCentroidY = 0#: mPrincipalAngle = 0#
    mConcreteCentroidX = 0#: mConcreteCentroidY = 0#: mConcretePrincipalAngle = 0#
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mLoadReferenceX = 0#: mLoadReferenceY = 0#
    mExtensionUsed = False
    mStateStatus = vbNullString

    Erase mElementID
    Erase mMaterialType
    Erase mX
    Erase mY
    Erase mArea
    Erase mDiameter
    Erase mShapeType
    Erase mWidth
    Erase mHeight
    Erase mRotation
    Erase mElementGeometry
    Erase mResultValue
    Erase mPhysicalState
    Erase mAnnotationType
    Erase mAnnotationStartX
    Erase mAnnotationStartY
    Erase mAnnotationEndX
    Erase mAnnotationEndY
    Erase mAnnotationNormalX
    Erase mAnnotationNormalY
    Erase mAnnotationText
    Erase mAnnotationSweepAngle
End Sub

' Читает таблицу от её именованного начала, без CurrentRegion.
' Текстовый заголовок над таблицей в массив данных не включается.
Private Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    If Left$(rangeName, 6) = "rngNDM" Then
        ReadAnchoredResultTable = modSectionSnapshot.ReadAnchoredResultTable(workbook, rangeName)
        Exit Function
    End If
    If Not mReadTables Is Nothing Then
        If mReadTables.Exists(rangeName) Then
            ReadAnchoredResultTable = mReadTables.Item(rangeName)
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
    If Not mReadTables Is Nothing Then mReadTables.Add rangeName, ReadAnchoredResultTable
End Function

' Определяет число столбцов по непрерывно заполненной строке заголовков от начала таблицы.
Private Function AnchoredColumnCount(ByVal anchor As Object) As Long
    AnchoredColumnCount = SnapshotAnchoredExtent(anchor, True, 256)
End Function

' Определяет высоту таблицы по первому столбцу. Пустой первый столбец означает
' конец данных; это соответствует формату всех Results-диапазонов.
Private Function AnchoredRowCount(ByVal anchor As Object) As Long
    AnchoredRowCount = SnapshotAnchoredExtent(anchor, False)
End Function

' Объясняет отсутствие запрошенного НДС в Results
' для выбранных сочетания и профиля.
Private Function MissingStateMessage(ByVal loadCase As String, ByVal stateType As String) As String
    MissingStateMessage = "Запрашиваемое состояние """ & stateType & """ не найдено в Results для сочетания " & loadCase & "."
    If IsCrackVisualizationState(stateType) Then
        If Not mProfileCrackWidthEnabled Then
            MissingStateMessage = MissingStateMessage & _
                " Для запрашиваемого сочетания расчет трещин не выполнялся: этому сочетанию назначен профиль, где Calculation.Crack.Width = No."
        ElseIf StrComp(stateType, "CrackedState", vbTextCompare) = 0 Then
            MissingStateMessage = MissingStateMessage & _
                " Для этого сочетания расчёт трещин был включён, но НДС с учётом трещин (CrackedState) не сохранено в Results. Проверьте статус и пояснения расчёта трещин; при необходимости повторите расчёт."
        Else
            MissingStateMessage = MissingStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, но пороговое состояние образования трещины не записано. Это штатно только когда нет действия для поиска нормальной трещины, например чистое сжатие без момента. Если растяжение или изгиб есть, проверьте статус проверки образования трещины."
        End If
    End If
End Function

' Проверяет, относится ли выбранный вид НДС к трещинам,
' чтобы точнее пояснить отсутствие нужного расчётного состояния.
Private Function IsCrackVisualizationState(ByVal stateType As String) As Boolean
    On Error GoTo NotCrackState
    Select Case SectionStateTypeFromText(stateType)
        Case sstPreCrackState, sstPostCrackState, sstCrackedState
            IsCrackVisualizationState = True
    End Select
NotCrackState:
End Function

' Читает сохраненный логический признак без подмены повреждения значением
' False. Допускает Boolean и явные текстовые/числовые пары Yes/No, True/False.
Private Function ReadSnapshotBoolean(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As Boolean
    Dim parsed As Boolean
    If Not TryReadSnapshotBoolean(value, parsed) Then Err.Raise vbObjectError + 4720, _
        "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
    ReadSnapshotBoolean = parsed
End Function

' Проверяет сохранённое обозначение статуса по общему словарю.
' Неизвестная строка означает повреждение данных Results,
' а не новый инженерный результат или самостоятельно назначенный N/A.
Private Function ReadSnapshotStatus(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    Dim parsed As String
    If Not TryReadSnapshotStatus(value, parsed) Then Err.Raise vbObjectError + 4720, _
        "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
    ReadSnapshotStatus = parsed
End Function

' Читает служебный текст с адресной диагностикой Excel-ошибки/Null. Пустая
' строка здесь допустима: обязательность определяет потребитель поля.
Private Function ReadSnapshotText(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    If IsError(value) Or IsNull(value) Then Err.Raise vbObjectError + 4720, _
        "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
    ReadSnapshotText = CStr(value)
End Function

' Берёт единицу из строки или общих свойств результата и проверяет её через CUnitSystem.
' Пустота и «-» допустимы только в необязательной ячейке Unit;
' ошибка Excel отсутствующей единицей не считается.
Private Function ReadSnapshotUnit(ByVal value As Variant, ByVal dimension As String, _
        ByVal fallbackUnit As String, ByVal hardDefault As String, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String, Optional ByVal allowBlank As Boolean = False) As String
    Dim unitText As String, checked As Double
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Then GoTo InvalidValue
    unitText = Trim$(CStr(value))
    If Len(unitText) = 0 Or unitText = "-" Then
        If Not allowBlank Then GoTo InvalidValue
        unitText = UnitOrDefault(unitText, fallbackUnit, hardDefault)
    End If
    Select Case dimension
        Case "Length": checked = mUnits.OutputLengthToInternal(0#, unitText)
        Case "Area": checked = mUnits.OutputAreaToInternal(0#, unitText)
        Case "Stress": checked = mUnits.OutputStressToInternal(0#, unitText)
        Case "Curvature": checked = mUnits.InternalCurvatureToOutput(0#, unitText)
    End Select
    ReadSnapshotUnit = unitText
    Exit Function
InvalidValue:
    Err.Raise vbObjectError + 4720, "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
End Function

' Сообщает о повреждённых данных с фактическим адресом Results.
' Пояснение предлагает восстановить сохранённый результат,
' а не менять текущие единицы ввода.
Private Function SnapshotMetadataInputMessage(ByVal anchor As Object, ByVal row As Long, _
        ByVal column As Long, ByVal field As String) As String
    SnapshotMetadataInputMessage = "В сохраненном снимке Results поле " & field & " (" & _
        anchor.Worksheet.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ") содержит недопустимое значение. Исправьте указанную ячейку или повторите расчет/импорт для восстановления снимка."
End Function

' Превращает значение ячейки в строку без падения на Excel-ошибках.
Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    SafeText = CStr(value)
End Function

' Выбирает единицу числового поля. Если Unit пуста или содержит «-»,
' берёт единицу из общих свойств сохранённого результата, а не текущего Config.
Private Function UnitOrDefault(ByVal unitText As String, ByVal fallbackUnit As String, _
        ByVal hardDefault As String) As String
    Dim normalized As String
    normalized = Trim$(unitText)
    If Len(normalized) > 0 And normalized <> "-" Then
        UnitOrDefault = normalized
    ElseIf Len(Trim$(fallbackUnit)) > 0 Then
        UnitOrDefault = Trim$(fallbackUnit)
    Else
        UnitOrDefault = hardDefault
    End If
End Function

' Проверяет, что табличный массив содержит хотя бы заголовок и одну строку данных.
Private Function HasRows(ByVal data As Variant) As Boolean
    On Error GoTo Failed
    HasRows = (UBound(data, 1) >= 2 And UBound(data, 2) >= 1)
    Exit Function
Failed:
End Function

' Находит индекс столбца по имени заголовка.
' Единицы в заголовке после запятой отбрасываются, поэтому "Stress, MPa"
' корректно сопоставляется с базовым именем "Stress".
Private Function ResultColumn(ByRef data As Variant, ByVal headerName As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", "В Results не найден столбец: " & headerName
End Function


' Возвращает имя столбца без единицы измерения.
Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

' Достает единицу из заголовка Results вида "X, mm". Если заголовок без
' запятой, берется единица, сохраненная в Output.*Unit.
Private Function ResultHeaderUnit(ByRef data As Variant, ByVal colIndex As Long, _
        ByVal defaultUnit As String) As String
    Dim headerText As String
    headerText = CStr(data(1, colIndex))

    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos <= 0 Then
        ResultHeaderUnit = UnitOrDefault(vbNullString, defaultUnit, "mm")
        Exit Function
    End If

    Dim unitText As String
    unitText = Trim$(Mid$(headerText, commaPos + 1))
    Dim bracketPos As Long
    bracketPos = InStr(1, unitText, "(", vbTextCompare)
    If bracketPos > 0 Then unitText = Trim$(Left$(unitText, bracketPos - 1))
    ResultHeaderUnit = UnitOrDefault(unitText, defaultUnit, "mm")
End Function

' Переводит длину из единицы, сохраненной в Results, во внутренние мм.
Private Function OutputLengthToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    OutputLengthToInternalByUnit = mUnits.OutputLengthToInternal(value, unitText)
End Function

' Переводит площадь из единицы, сохраненной в Results, во внутренние мм2.
Private Function OutputAreaToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    OutputAreaToInternalByUnit = mUnits.OutputAreaToInternal(value, unitText)
End Function

' Переводит кривизну из единицы Results во внутренние 1/мм, чтобы нейтральная
' линия строилась в той же системе, что и координаты схемы.
Private Function OutputCurvatureToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    OutputCurvatureToInternalByUnit = value / mUnits.InternalCurvatureToOutput(1#, unitText) ' Деление сохраняет прежний порядок операций для 1/m, без нового округления.
End Function

' Возвращает число прочитанных элементов. Для схемы НДС учитывает только
' элементы с результатом выбранного состояния; для предварительной схемы — всю геометрию.
Public Property Get Count() As Long: Count = mCount: End Property
' ID элемента для связи геометрии, результатов и подписей.
Public Property Get ElementID(ByVal index As Long) As String: ElementID = mElementID(index): End Property
' Concrete/Rebar; Plotter использует это для выбора маркера.
Public Property Get MaterialType(ByVal index As Long) As String: MaterialType = mMaterialType(index): End Property
' Координата X центра элемента, переведённая из Results во внутренние мм.
Public Property Get X(ByVal index As Long) As Double: X = mX(index): End Property
' Координата Y центра элемента, переведённая из Results во внутренние мм.
Public Property Get Y(ByVal index As Long) As Double: Y = mY(index): End Property
' Площадь элемента, переведённая из Results во внутренние мм2.
Public Property Get Area(ByVal index As Long) As Double: Area = mArea(index): End Property
' Диаметр арматуры в мм; для бетона обычно 0.
Public Property Get Diameter(ByVal index As Long) As Double: Diameter = mDiameter(index): End Property
' Статус интерпретации оболочки: Rectangle, Equivalent rectangle, Equivalent square.
Public Property Get ShapeType(ByVal index As Long) As String: ShapeType = mShapeType(index): End Property
' Ширина бетонного элемента в мм, если она сохранена в Results.
Public Property Get Width(ByVal index As Long) As Double: Width = mWidth(index): End Property
' Высота бетонного элемента в мм, если она сохранена в Results.
Public Property Get Height(ByVal index As Long) As Double: Height = mHeight(index): End Property
' Возвращает угол прямоугольного элемента, рад; для эквивалентного квадрата
' может быть сохранён средний угол сетки.
Public Property Get Rotation(ByVal index As Long) As Double: Rotation = mRotation(index): End Property

' Возвращает одну универсальную бетонную область во внутренних мм.
' Сохраненный Contour не заменяется прямоугольником; старые снимки получают
' прежнюю визуальную оболочку с переданной общей ориентацией эквивалентных квадратов.
Public Function ConcreteGeometry(ByVal index As Long, Optional ByVal fallbackRotation As Double = 0#) As CConcreteRegion
    If Not mElementGeometry(index) Is Nothing Then
        Set ConcreteGeometry = mElementGeometry(index)
        Exit Function
    End If
    Dim width As Double, height As Double, rotation As Double, status As String
    width = mWidth(index): height = mHeight(index): rotation = mRotation(index)
    If width <= 0# Then width = Sqr(mArea(index))
    If height <= 0# Then height = Sqr(mArea(index))
    status = Replace$(Trim$(mShapeType(index)), " ", vbNullString)
    If StrComp(status, "EquivalentSquare", vbTextCompare) = 0 Then rotation = fallbackRotation
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    Set ConcreteGeometry = query.RectangleRegion(mX(index), mY(index), width, height, rotation)
End Function

' Показывает наличие явной области, чтобы большая схема не заменяла ее
' квадратным маркером. Это свойство представления, а не режим формирования сетки.
Public Property Get ConcreteHasBoundary(ByVal index As Long) As Boolean
    ConcreteHasBoundary = Not mElementGeometry(index) Is Nothing
End Property
' Значение Stress или Strain для выбранного LC.
Public Property Get ResultValue(ByVal index As Long) As Double: ResultValue = mResultValue(index): End Property
' Сохранённое текстовое описание физической пригодности выбранного состояния НДС.
Public Property Get PhysicalState(ByVal index As Long) As String: PhysicalState = mPhysicalState(index): End Property
' LC, выбранный для схемы после обработки Worst/конкретного ID.
Public Property Get LoadCase() As String: LoadCase = mLoadCase: End Property
' Комментарий LC для заголовка схемы.
Public Property Get LoadCaseComment() As String: LoadCaseComment = mLoadCaseComment: End Property
' Имя профиля выбранного сочетания, сохранённое вместе с результатами.
Public Property Get ProfileId() As String: ProfileId = mProfileId: End Property
' Вид НДС для схемы, выбранный настройкой Visualization.State профиля.
Public Property Get StateType() As String: StateType = mStateType: End Property
' True для схемы без расчетных Stress/Strain.
Public Property Get GeometryOnly() As Boolean: GeometryOnly = mGeometryOnly: End Property
' Заголовок схемы без НДС, например для import-preview.
Public Property Get GeometryOnlyTitle() As String: GeometryOnlyTitle = mGeometryOnlyTitle: End Property
' Пояснение под схемой, если НДС намеренно отсутствует.
Public Property Get GeometryOnlyNote() As String: GeometryOnlyNote = mGeometryOnlyNote: End Property
' Можно ли показывать оси/точку нагрузки без НДС.
Public Property Get GeometryReferenceItemsEnabled() As Boolean: GeometryReferenceItemsEnabled = mGeometryReferenceItemsEnabled: End Property
' Что сейчас визуализируется по профилю: Stress или Strain.
Public Property Get VisualizationQuantity() As String: VisualizationQuantity = mVisualizationQuantity: End Property
' Единица выбранной величины из сохраненных Results.
Public Property Get ResultUnit() As String: ResultUnit = mResultUnit: End Property
' Название выбранной величины для легенды.
Public Property Get ResultCaption() As String: ResultCaption = mResultCaption: End Property
' Количество знаков после запятой для выбранной Stress/Strain.
Public Property Get ResultPrecision() As Long: ResultPrecision = mResultPrecision: End Property
' Возвращает единицу координат подготовленной схемы — мм.
Public Property Get LengthUnit() As String: LengthUnit = mLengthUnit: End Property
' Возвращает единицу площади подготовленной схемы — мм2.
Public Property Get AreaUnit() As String: AreaUnit = mAreaUnit: End Property
' Левая граница бетонной расчётной модели в мм.
Public Property Get MinX() As Double: MinX = mMinX: End Property
' Правая граница бетонной расчётной модели в мм.
Public Property Get MaxX() As Double: MaxX = mMaxX: End Property
' Нижняя граница бетонной расчётной модели в мм.
Public Property Get MinY() As Double: MinY = mMinY: End Property
' Верхняя граница бетонной расчётной модели в мм.
Public Property Get MaxY() As Double: MaxY = mMaxY: End Property
' X центра тяжести приведённого сечения в мм.
Public Property Get CentroidX() As Double: CentroidX = mCentroidX: End Property
' Y центра тяжести приведённого сечения в мм.
Public Property Get CentroidY() As Double: CentroidY = mCentroidY: End Property
' Угол от +X к главной оси 1 приведенного сечения, рад.
Public Property Get PrincipalAngle() As Double: PrincipalAngle = mPrincipalAngle: End Property

' Возвращает X центра для выбранного набора главных осей.
' Transformed - приведенное сечение бетон + арматура, Concrete - только бетон.
Public Function PrincipalAxesCentroidX(ByVal axesMode As String) As Double
    If StrComp(Trim$(axesMode), "Concrete", vbTextCompare) = 0 Then
        PrincipalAxesCentroidX = mConcreteCentroidX
    Else
        PrincipalAxesCentroidX = mCentroidX
    End If
End Function

' Возвращает Y центра для выбранного набора главных осей.
Public Function PrincipalAxesCentroidY(ByVal axesMode As String) As Double
    If StrComp(Trim$(axesMode), "Concrete", vbTextCompare) = 0 Then
        PrincipalAxesCentroidY = mConcreteCentroidY
    Else
        PrincipalAxesCentroidY = mCentroidY
    End If
End Function

' Возвращает угол главной оси 1 для выбранного набора осей.
Public Function PrincipalAxesAngle(ByVal axesMode As String) As Double
    If StrComp(Trim$(axesMode), "Concrete", vbTextCompare) = 0 Then
        PrincipalAxesAngle = mConcretePrincipalAngle
    Else
        PrincipalAxesAngle = mPrincipalAngle
    End If
End Function
' Свободный член сохраненной плоскости деформаций LC.
Public Property Get Epsilon0() As Double: Epsilon0 = mEpsilon0: End Property
' Коэффициент при y в плоскости деформаций, 1/мм.
Public Property Get KappaX() As Double: KappaX = mKappaX: End Property
' Коэффициент при x в плоскости деформаций, 1/мм.
Public Property Get KappaY() As Double: KappaY = mKappaY: End Property
' Координата X общей точки приложения нагрузки, сохранённая в Results.
Public Property Get LoadReferenceX() As Double: LoadReferenceX = mLoadReferenceX: End Property
' Координата Y общей точки приложения нагрузки, сохранённая в Results.
Public Property Get LoadReferenceY() As Double: LoadReferenceY = mLoadReferenceY: End Property
' Признак использования численного продолжения диаграмм в выбранном НДС.
Public Property Get ExtensionUsed() As Boolean: ExtensionUsed = mExtensionUsed: End Property
' Статус выбранного сохранённого расчётного состояния.
Public Property Get DirectStateStatus() As String: DirectStateStatus = mStateStatus: End Property
' Проверяет по Results, нужно ли предупреждение о продолжении диаграмм
' или физическом FAIL. Решатель не запускается; обозначение статуса
' остаётся заданным общими правилами.
Public Property Get PhysicalDiagramWarningRequired() As Boolean
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    PhysicalDiagramWarningRequired = (mExtensionUsed Or policy.ToUserStatus(mStateStatus) = policy.Fail)
End Property
' Число размеров, подписей и исходных контуров для схемы; расчётные области трещин сюда не включаются.
Public Property Get AnnotationCount() As Long: AnnotationCount = mAnnotationCount: End Property

' Отсчёт отступа подписи от сохранённого бетонного контура. На прямой грани
' это добавляет as; на дуге определяется действительное расстояние по нормали.
' Старые Results также читаются без нового расчёта или обращения к геометрии Config.
Public Function AnnotationBoundaryDistance(ByVal index As Long) As Double
    If mAnnotationType(index) = "REBAR_GLOBAL_ANNOTATION" Then Exit Function ' Общая подпись уже привязана к крайней точке бетона. Пересечение по нормали в середине повёрнутого сечения опустило бы её на контур.
    If mLabelContours Is Nothing Then Exit Function
    Dim query As CSectionGeometryQuery, i As Long, hasOuter As Boolean
    Dim nx As Double, ny As Double, norm As Double, distance As Double
    If mLabelBoundary Is Nothing Then
        For i = 1 To mLabelContours.Count
            If StrComp(mLabelContours.LoopRole(i), "Outer", vbTextCompare) = 0 Then hasOuter = True: Exit For
        Next i
        If Not hasOuter Then Exit Function ' Импорт без достоверного Outer сохраняет приблизительную опорную линию.
        Set query = New CSectionGeometryQuery: Set mLabelBoundary = query.RegionFromContours(mLabelContours)
    End If
    Set query = New CSectionGeometryQuery
    nx = mAnnotationNormalX(index): ny = mAnnotationNormalY(index)
    norm = Sqr(nx ^ 2 + ny ^ 2)
    If norm <= GEOM_TOLERANCE Then
        nx = -(mAnnotationEndY(index) - mAnnotationStartY(index))
        ny = mAnnotationEndX(index) - mAnnotationStartX(index)
        If Sqr(nx ^ 2 + ny ^ 2) <= GEOM_TOLERANCE Then nx = 0#: ny = 1#
    End If
    If query.LabelBoundaryDistance(mLabelBoundary, 0.5 * (mAnnotationStartX(index) + mAnnotationEndX(index)), _
            0.5 * (mAnnotationStartY(index) + mAnnotationEndY(index)), nx, ny, distance) Then AnnotationBoundaryDistance = distance
End Function
' Тип аннотации: DIMENSION, REBAR_ANNOTATION, REBAR_GLOBAL_ANNOTATION, TEXT_ANNOTATION или CONTOUR_*.
Public Property Get AnnotationType(ByVal index As Long) As String: AnnotationType = mAnnotationType(index): End Property
' Начало линии/размера в координатах модели.
Public Property Get AnnotationStartX(ByVal index As Long) As Double: AnnotationStartX = mAnnotationStartX(index): End Property
' Начало линии/размера в координатах модели.
Public Property Get AnnotationStartY(ByVal index As Long) As Double: AnnotationStartY = mAnnotationStartY(index): End Property
' Конец линии/размера в координатах модели.
Public Property Get AnnotationEndX(ByVal index As Long) As Double: AnnotationEndX = mAnnotationEndX(index): End Property
' Конец линии/размера в координатах модели.
Public Property Get AnnotationEndY(ByVal index As Long) As Double: AnnotationEndY = mAnnotationEndY(index): End Property
' Компонента безразмерной нормали; для CONTOUR_ARC — координата центра дуги в мм.
Public Property Get AnnotationNormalX(ByVal index As Long) As Double: AnnotationNormalX = mAnnotationNormalX(index): End Property
' Компонента безразмерной нормали; для CONTOUR_ARC — координата центра дуги в мм.
Public Property Get AnnotationNormalY(ByVal index As Long) As Double: AnnotationNormalY = mAnnotationNormalY(index): End Property
' Готовый текст подписи или размера.
Public Property Get AnnotationText(ByVal index As Long) As String: AnnotationText = mAnnotationText(index): End Property
' Возвращает прочитанный и проверенный угол без повторного разбора текстовых значений Results.
Public Property Get AnnotationSweepAngle(ByVal index As Long) As Double: AnnotationSweepAngle = mAnnotationSweepAngle(index): End Property
