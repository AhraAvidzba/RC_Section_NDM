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
' Чтение данных схемы из Results
' ==========================================================================
' Reader собирает из сохраненного Results-снимка геометрию, LC-зависимые
' Stress/Strain, свойства сечения и semantic-аннотации. Он объединяет
' rngNDMSectionGeometry с rngNDMElementResults по ElementID и отдает Plotter-у
' уже готовую модель отображения.
'
' Reader не запускает расчет и не читает текущую геометрию Config. Длины и
' кривизны из Results он переводит обратно во внутренние мм/1 мм только для
' согласованной отрисовки схемы; Stress/Strain остаются сохраненными значениями.

Private mVisualizationQuantity As String ' Что сейчас визуализируется по профилю: Stress или Strain.
Private mLoadCase As String ' LC, выбранный для схемы после обработки Worst/конкретного ID.
Private mLoadCaseComment As String ' Комментарий LC для заголовка схемы.
Private mProfileId As String ' ProfileId, сохраненный в snapshot для выбранного LC.
Private mStateType As String ' StateType, выбранный через Visualization.State текущего профиля.
Private mProfileCrackWidthEnabled As Boolean ' В профиле выбранного LC был запрошен Calculation.Crack.Width.
Private mGeometryOnly As Boolean ' True для схемы без расчетных Stress/Strain.
Private mGeometryOnlyTitle As String ' Заголовок схемы без НДС, например для import-preview.
Private mGeometryOnlyNote As String ' Пояснение под схемой, если НДС намеренно отсутствует.
Private mGeometryReferenceItemsEnabled As Boolean ' Можно ли показывать оси/точку нагрузки без НДС.
Private mResultUnit As String ' Единица выбранной величины из сохраненных Results.
Private mResultCaption As String ' Название выбранной величины для легенды.
Private mResultPrecision As Long ' Количество знаков после запятой для выбранной Stress/Strain.
Private mLengthUnit As String ' Единица модельной геометрии после чтения Results; для Plotter это всегда mm.
Private mAreaUnit As String ' Единица площади после чтения Results; для Plotter это всегда mm2.
Private mSnapshotLengthUnit As String ' Исходная OUTPUT-единица длины, сохраненная в Results.
Private mSnapshotAreaUnit As String ' Исходная OUTPUT-единица площади, сохраненная в Results.
Private mStressUnit As String ' Единица напряжений, сохраненная при последнем расчете.
Private mCurvatureUnit As String ' Единица кривизн после чтения Results; для Plotter это всегда 1/mm.
Private mSnapshotCurvatureUnit As String ' Исходная OUTPUT-единица кривизн, сохраненная в Results.
Private mCount As Long ' Количество элементов, у которых есть геометрия и результат выбранного LC.
Private mElementCapacity As Long ' Закрытый резерв одного Load; публичный Count не включает хвост.
Private mReadTables As Object ' Value2-копии только текущего Load, освобождаются до передачи reader Plotter-у.
Private mElementID() As String ' ID элемента для связи геометрии, результатов и подписей.
Private mMaterialType() As String ' Concrete/Rebar; Plotter использует это для выбора маркера.
Private mX() As Double ' X элемента во внутренних мм после чтения snapshot.
Private mY() As Double ' Y элемента во внутренних мм после чтения snapshot.
Private mArea() As Double ' Площадь элемента во внутренних мм2 после чтения snapshot.
Private mDiameter() As Double ' Диаметр арматуры; для бетона обычно 0.
Private mShapeType() As String ' Статус интерпретации оболочки: Rectangle, Equivalent rectangle, Equivalent square.
Private mWidth() As Double ' Реальная ширина бетонного элемента, если она сохранена в snapshot.
Private mHeight() As Double ' Реальная высота бетонного элемента, если она сохранена в snapshot.
Private mRotation() As Double ' Поворот прямоугольного бетонного элемента, рад; для fallback-а может хранить средний угол сетки.
Private mResultValue() As Double ' Значение Stress или Strain для выбранного LC.
Private mPhysicalState() As String ' Compression/Tension/InactiveTensionConcrete/NearZero.
Private mMinX As Double ' Левая граница бетонной расчетной модели.
Private mMaxX As Double ' Правая граница бетонной расчетной модели.
Private mMinY As Double ' Нижняя граница бетонной расчетной модели.
Private mMaxY As Double ' Верхняя граница бетонной расчетной модели.
Private mCentroidX As Double ' X центра тяжести приведенного сечения.
Private mCentroidY As Double ' Y центра тяжести приведенного сечения.
Private mPrincipalAngle As Double ' Угол от +X к главной оси 1 приведенного сечения, рад.
Private mConcreteCentroidX As Double ' X центра тяжести бетонного сечения из Results.
Private mConcreteCentroidY As Double ' Y центра тяжести бетонного сечения из Results.
Private mConcretePrincipalAngle As Double ' Угол главной оси 1 бетонного сечения из Results.
Private mEpsilon0 As Double ' Свободный член сохраненной плоскости деформаций LC.
Private mKappaX As Double ' Кривизна при y в сохраненной плоскости деформаций.
Private mKappaY As Double ' Кривизна при x в сохраненной плоскости деформаций.
Private mLoadReferenceX As Double ' X общей точки приложения нагрузки из Results snapshot.
Private mLoadReferenceY As Double ' Y общей точки приложения нагрузки из Results snapshot.
Private mExtensionUsed As Boolean ' True, если выбранный LC был найден в numerical extension.
Private mStateStatus As String ' Сохраненный статус выбранного named-state.
Private mAnnotationCount As Long ' Количество semantic-аннотаций, прочитанных из Results.
Private mAnnotationType() As String ' Dimension, RebarLabel, Outline и другие типы оформления.
Private mAnnotationStartX() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationStartY() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationEndX() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationEndY() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationNormalX() As Double ' Наружная нормаль semantic-аннотации.
Private mAnnotationNormalY() As Double ' Наружная нормаль semantic-аннотации.
Private mAnnotationText() As String ' Готовый текст подписи или размера.
Private mAnnotationSweepAngle() As Double ' Проверенный signed sweep CONTOUR_ARC, рад; для остальных типов 0.
Private mUnits As CUnitSystem ' Единый адаптер единиц самого snapshot; текущий Config не меняет сохраненные числа.

' Создает один адаптер для всего чтения, без отдельных объектов на каждый
' элемент. Единицы передаются из Results явно, пользовательские знаки уже применены.
Private Sub Class_Initialize()
    Set mUnits = New CUnitSystem
End Sub

' Загружает все данные, необходимые для Excel-схемы.
' Метод намеренно читает расчетные значения только из Results. Из Config
' используется только выбор LC и текущие настройки визуализации профиля.
' При ошибке частичный снимок удаляется; исходная причина передается caller-у.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    On Error GoTo Failed
    ResetReaderState
    Set mReadTables = CreateObject("Scripting.Dictionary")
    mLoadCase = ResolveLoadCase(workbook, settings.GetRequiredString("Plot.LoadCase"))
    ApplyVisualizationProfile workbook

    ReadSectionProperties workbook
    ReadGeometryAndResults workbook
    ReadAnnotations workbook
    Set mReadTables = Nothing
    Exit Sub
Failed:
    Dim errorNumber As Long, description As String, errorSource As String
    errorNumber = Err.Number: description = Err.Description: errorSource = Err.Source
    ResetReaderState
    Err.Raise errorNumber, errorSource, description
End Sub

' Загружает геометрию последнего AutoCAD import для визуального контроля.
' Расчетные результаты и нейтральная линия не читаются, но сохраненные
' габаритные аннотации по Bounds можно показать сразу после импорта.
' Неуспешное чтение не оставляет доступной часть предыдущего или нового снимка.
Public Sub LoadGeometryPreviewFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    On Error GoTo Failed
    ResetReaderState
    Set mReadTables = CreateObject("Scripting.Dictionary")
    mGeometryOnly = True
    mGeometryOnlyTitle = "Импортированная геометрия AutoCAD"
    mVisualizationQuantity = "Stress"
    mLoadCase = "AutoCAD import"

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

' Загружает расчетную геометрию из Results, когда выбранное состояние НДС
' отсутствует в snapshot. Это не AutoCAD-preview и не повторный расчет:
' пользователь видит актуальную геометрию последнего запуска и понятную
' подпись, почему цветовая раскраска Stress/Strain недоступна.
' Ошибка восстановления очищает reader, не подменяя отсутствующие данные нулями.
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
' Значение Worst разворачивается в LC из верхней строки batch summary
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
        "В rngBatchSummary не найдено определяющее сочетание для Plot.LoadCase = Worst."
End Function

' Читает верхнее поле "Определяющее сочетание" из новой якорной сводки.
Private Function ReadWorstCombinationID(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)
    ReadWorstCombinationID = Trim$(SafeText(anchor.Value2))
Failed:
End Function

' Применяет настройки визуализации профиля, назначенного выбранному LC.
' ProfileId читается из snapshot Results, а не из текущей строки сочетаний:
' схема должна показывать ровно тот расчетный запуск, который сохранен.
Private Sub ApplyVisualizationProfile(ByVal workbook As Object)
    mProfileId = ReadProfileIdForLoadCase(workbook, mLoadCase)
    If Len(mProfileId) = 0 Then Err.Raise vbObjectError + 4710, "CSectionPlotDataReader", _
        "В Results не найден ProfileId для сочетания: " & mLoadCase

    Dim profiles As CCalculationProfileCatalog
    Set profiles = New CCalculationProfileCatalog
    profiles.LoadFromWorkbook workbook

    Dim profile As CCalculationProfile
    Set profile = profiles.ProfileById(mProfileId)
    mProfileCrackWidthEnabled = profile.CrackWidthEnabled
    mStateType = profile.VisualizationStateText
    mVisualizationQuantity = VisualizationQuantityToText(profile.VisualizationQuantity)
    mResultPrecision = profile.VisualizationPrecisionForQuantity(profile.VisualizationQuantity)
End Sub

' Находит ProfileId выбранного LC в Results. Поэлементная таблица является
' основным snapshot-источником; если в выбранном LC нет элементов нужного
' состояния, fallback на свойства сечения нужен только для более понятной
' ошибки визуализации.
Private Function ReadProfileIdForLoadCase(ByVal workbook As Object, ByVal loadCase As String) As String
    ReadProfileIdForLoadCase = ReadProfileIdFromElementResults(workbook, loadCase)
    If Len(ReadProfileIdForLoadCase) = 0 Then
        ReadProfileIdForLoadCase = ReadLoadCasePropertyText(workbook, loadCase, "ProfileId")
    End If
End Function

' Ищет ProfileId в LC-зависимой таблице элементов Results. Это основной
' путь, потому что именно в ней каждое состояние связано с расчетным профилем.
Private Function ReadProfileIdFromElementResults(ByVal workbook As Object, ByVal loadCase As String) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
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

' Ищет текстовое свойство конкретного LC в rngNDMSectionProperties.
' Используется как fallback, если для состояния нет строк элементов.
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

' Возвращает первый LC, для которого на Results реально есть результаты элементов.
' Это запасной путь только для режимов, где Plot.LoadCase = Worst не может быть
' развернут в прочностное определяющее сочетание, например DirectState без
' расчета несущей способности. AutoCAD-preview этим путем не выбирается,
' потому что geometry-only snapshot не содержит rngNDMElementResults.
Private Function FirstCalculatedLoadCaseFromResults(ByVal workbook As Object) As String
    On Error GoTo Failed
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
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

' Возвращает первый LC из таблицы свойств сечения.
' В некоторых профилях может не быть состояния для визуализации, но свойства
' LC все равно помогают показать пользователю, что сохранено в snapshot.
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
' Если комментария нет или batch summary еще пустой, возвращается пустая строка
' без ошибки: сама схема может строиться и без комментария.
Private Function ReadBatchSummaryComment(ByVal workbook As Object, ByVal loadCase As String) As String
    On Error GoTo Failed
    If Len(Trim$(loadCase)) = 0 Then Exit Function

    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngBatchSummary").RefersToRange(1, 1)

    Dim rowOffset As Long, combinationID As String
    For rowOffset = 11 To anchor.Worksheet.Rows.Count - anchor.Row
        combinationID = Trim$(SafeText(anchor.Offset(rowOffset, 0).Value2))
        If Len(combinationID) = 0 Then Exit For
        If StrComp(combinationID, loadCase, vbTextCompare) = 0 Then
            ReadBatchSummaryComment = Trim$(SafeText(anchor.Offset(rowOffset, 1).Value2))
            Exit Function
        End If
    Next rowOffset
Failed:
End Function

' Читает свойства уровня всего сечения и выбранного LC.
' Здесь находятся Bounds, центр тяжести, главные оси, единицы Results и
' сохраненная плоскость деформаций epsilon0/kappaX/kappaY.
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
    ' Сначала читаем единицы snapshot: порядок строк не должен влиять на
    ' пересчет свойств, у которых единица в собственной строке не задана.
    For rowIndex = 2 To UBound(data, 1)
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
    ' В одной таблице лежат свойства ALL и свойства конкретных LC. Reader
    ' выбирает только общие строки и строки mLoadCase, остальные LC игнорируются.
    For rowIndex = 2 To UBound(data, 1)
        parameter = Trim$(ReadSnapshotText(data(rowIndex, colParameter), anchor, rowIndex, colParameter, "Parameter"))
        loadCase = Trim$(ReadSnapshotText(data(rowIndex, colLoadCase), anchor, rowIndex, colLoadCase, "LoadCase"))
        If Len(parameter) > 0 Then
            If StrComp(loadCase, "ALL", vbTextCompare) = 0 Or Len(loadCase) = 0 Then
                If Left$(LCase$(parameter), 7) <> "output." Then _
                    ReadCommonProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
            ElseIf StrComp(loadCase, mLoadCase, vbTextCompare) = 0 Then
                If Not mGeometryOnly Or mGeometryReferenceItemsEnabled Then _
                    ReadLoadCaseProperty parameter, data(rowIndex, colValue), data(rowIndex, colUnit), anchor, rowIndex, colValue, colUnit
            End If
        End If
    Next rowIndex

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

' Запоминает общие свойства расчетного снимка Results.
' Эти данные одинаково подходят для Generated и AutoCAD-геометрии.
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

' Запоминает свойства выбранного LC: плоскость деформаций и признаки
' допустимости найденного состояния. Точка приложения нагрузки читается из
' общего блока ALL, потому что она задается один раз для всего snapshot.
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
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."

    Dim results As Variant
    Dim resultAnchor As Object
    Set resultAnchor = workbook.Names.Item("rngNDMElementResults").RefersToRange
    results = ReadAnchoredResultTable(workbook, "rngNDMElementResults")
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
    ' LC-зависимые результаты сначала собираются в словарь по ElementID. Это
    ' сохраняет разделение Results: геометрия лежит отдельно, а Stress/Strain
    ' выбираются только для нужного LoadCase.
    For rowIndex = 2 To UBound(results, 1)
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
    If UBound(geom, 2) <> 15 Then Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", _
        "Таблица расчетной геометрии Results должна содержать 15 столбцов. Повторите расчет или импорт геометрии."
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

    ' На втором проходе геометрия соединяется с найденным состоянием. Plotter после этого
    ' получает готовые элементы и не обращается к Results напрямую.
    For rowIndex = 2 To UBound(geom, 1)
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
        End If
    Next rowIndex
End Sub

' Читает только постоянную геометрию элементов и добавляет ее в модель схемы
' с нейтральным состоянием. Этот режим используется для AutoCAD-preview и для
' схемы без выбранного StateType, если геометрия уже есть в Results.
Private Sub ReadGeometryPreview(ByVal workbook As Object)
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."
    If UBound(geom, 2) <> 15 Then Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", _
        "Таблица расчетной геометрии Results должна содержать 15 столбцов. Повторите расчет или импорт геометрии."

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
    neutralResult = Array(0#, 0#, "NearZero")

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
        End If
    Next rowIndex
End Sub

' Проверяет численную геометрию до добавления элемента схемы. Пустой диаметр
' бетона и отсутствующая оболочка допустимы; поврежденные ячейки не скрываются
' как нулевые координаты/размеры. Адрес берется от текущего Results-якоря.
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
    If Not IsNumeric(value) Then GoTo InvalidNumber
    On Error GoTo InvalidNumber
    ReadSnapshotNumber = CDbl(value)
    Exit Function
InvalidNumber:
    Err.Raise vbObjectError + 4719, "CSectionPlotDataReader", _
        "В сохраненном снимке Results поле " & field & _
        " содержит нечисловое значение: " & anchor.Worksheet.Name & "!" & _
        anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ". Повторите импорт геометрии или расчет, чтобы восстановить корректный снимок."
End Function

' Добавляет элемент схемы во внутренние массивы reader-а. Значение результата
' выбирается по текущему режиму визуализации: Stress или Strain.
Private Sub AppendElement(ByVal elementID As String, ByVal materialType As String, ByVal xCoord As Double, _
        ByVal yCoord As Double, ByVal area As Double, ByVal diameter As Double, _
        ByVal shapeType As String, ByVal width As Double, ByVal height As Double, _
        ByVal rotation As Double, ByVal resultData As Variant)
    ' Сохраняем только элементы, у которых есть и геометрия, и результат
    ' выбранного LC. Так схема не рисует “пустые” элементы после ошибочного LC.
    EnsureElementCapacity mCount + 1
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

' Сохраняет массивы элементов в прежнем порядке с геометрическим резервом.
' Незанятые позиции недоступны через getters; ResetReaderState стирает резерв.
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
    ReDim Preserve mResultValue(1 To capacity)
    ReDim Preserve mPhysicalState(1 To capacity)
    mElementCapacity = capacity
End Sub

' Читает semantic-аннотации схемы.
' Если таблица отсутствует или пуста, схема все равно строится: для AutoCAD
' import это нормальный сценарий, потому что характерные грани могут быть неизвестны.
Private Sub ReadAnnotations(ByVal workbook As Object)
    ReadMaterialContours workbook
    ' Отсутствие необязательного Name отличается от повреждения его данных.
    Dim tableName As Object, anchor As Object
    On Error Resume Next
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
    Dim colUnit As Long: colUnit = ResultColumn(data, "Unit")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim annotationType As String
        annotationType = Trim$(ReadAnnotationText(data(rowIndex, colType), anchor, rowIndex, colType, "AnnotationType"))
        If IsPlotLabelAnnotation(annotationType) Then
            Dim lengthUnit As String
            lengthUnit = ReadSnapshotUnit(data(rowIndex, colUnit), "Length", mSnapshotLengthUnit, "mm", anchor, rowIndex, colUnit, "Unit", True)
            Dim startX As Double, startY As Double, endX As Double, endY As Double
            Dim normalX As Double, normalY As Double, textValue As String
            startX = ReadAnnotationNumber(data(rowIndex, colStartX), anchor, rowIndex, colStartX, "StartX")
            startY = ReadAnnotationNumber(data(rowIndex, colStartY), anchor, rowIndex, colStartY, "StartY")
            endX = ReadAnnotationNumber(data(rowIndex, colEndX), anchor, rowIndex, colEndX, "EndX")
            endY = ReadAnnotationNumber(data(rowIndex, colEndY), anchor, rowIndex, colEndY, "EndY")
            normalX = ReadAnnotationNumber(data(rowIndex, colNormalX), anchor, rowIndex, colNormalX, "OutsideNormalX")
            normalY = ReadAnnotationNumber(data(rowIndex, colNormalY), anchor, rowIndex, colNormalY, "OutsideNormalY")
            textValue = ReadAnnotationText(data(rowIndex, colText), anchor, rowIndex, colText, "Text")
            mAnnotationCount = mAnnotationCount + 1
            ReDim Preserve mAnnotationType(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationText(1 To mAnnotationCount)
            ReDim Preserve mAnnotationSweepAngle(1 To mAnnotationCount)
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
        Case "DIMENSION", "REBAR_ANNOTATION"
            IsPlotLabelAnnotation = True
    End Select
End Function

' Добавляет исходные контуры в готовые render-массивы схемы. Это адаптация
' представления, а не повторное хранение контуров в result-аннотациях.
Private Sub ReadMaterialContours(ByVal workbook As Object)
    Dim contours As CSectionContours, i As Long
    Set contours = ReadSavedSectionContours(workbook)
    For i = 1 To contours.Count
        mAnnotationCount = mAnnotationCount + 1
        ReDim Preserve mAnnotationType(1 To mAnnotationCount)
        ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
        ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
        ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
        ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
        ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
        ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
        ReDim Preserve mAnnotationText(1 To mAnnotationCount)
        ReDim Preserve mAnnotationSweepAngle(1 To mAnnotationCount)
        mAnnotationType(mAnnotationCount) = contours.SegmentType(i)
        mAnnotationStartX(mAnnotationCount) = contours.StartX(i)
        mAnnotationStartY(mAnnotationCount) = contours.StartY(i)
        mAnnotationEndX(mAnnotationCount) = contours.EndX(i)
        mAnnotationEndY(mAnnotationCount) = contours.EndY(i)
        mAnnotationNormalX(mAnnotationCount) = contours.CenterX(i)
        mAnnotationNormalY(mAnnotationCount) = contours.CenterY(i)
        mAnnotationSweepAngle(mAnnotationCount) = contours.SweepAngle(i)
    Next i
End Sub

' Читает координату/нормаль сохраненной аннотации без подмены ошибки нулем.
' Число 0 корректно; пустое или нечисловое обязательное значение требует
' повторного получения snapshot либо исправления указанной ячейки Results.
Private Function ReadAnnotationNumber(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As Double
    If IsError(value) Then GoTo InvalidValue
    If IsNull(value) Or IsEmpty(value) Then GoTo InvalidValue
    If Not IsNumeric(value) Then GoTo InvalidValue
    On Error GoTo InvalidValue
    ReadAnnotationNumber = CDbl(value)
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

' Указывает фактическую ячейку snapshot по текущему именованному якорю;
' перенос Results-блока не оставляет старый жестко заданный адрес ошибки.
Private Function AnnotationInputMessage(ByVal anchor As Object, ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    AnnotationInputMessage = "В сохраненной аннотации Results поле " & field & " (" & _
        anchor.Parent.Name & "!" & anchor.Offset(row - 1, column - 1).Address(False, False) & _
        ") содержит недопустимое значение. Исправьте данные или повторите расчет/импорт для восстановления снимка."
End Function

' Сбрасывает внутренние массивы перед новым чтением Results.
' Reader может переиспользоваться тестами, поэтому нельзя оставлять элементы
' и аннотации от предыдущей загрузки.
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

' Читает таблицу Results от именованного якоря, не полагаясь на CurrentRegion.
' Это позволяет размещать человекочитаемый заголовок прямо над таблицей и не
' включать его в массив данных.
Private Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
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

' Определяет ширину таблицы по непрерывной строке заголовков от якоря.
Private Function AnchoredColumnCount(ByVal anchor As Object) As Long
    AnchoredColumnCount = SnapshotAnchoredExtent(anchor, True, 256)
End Function

' Определяет высоту таблицы по первому столбцу. Пустой первый столбец означает
' конец данных; это соответствует формату всех Results-диапазонов.
Private Function AnchoredRowCount(ByVal anchor As Object) As Long
    AnchoredRowCount = SnapshotAnchoredExtent(anchor, False)
End Function

' Формирует понятное сообщение, если пользователь запросил состояние, которого
' нет в snapshot Results для выбранного сочетания и профиля.
Private Function MissingStateMessage(ByVal loadCase As String, ByVal stateType As String) As String
    MissingStateMessage = "Запрашиваемое состояние """ & stateType & """ не найдено в Results для сочетания " & loadCase & "."
    If IsCrackVisualizationState(stateType) Then
        If Not mProfileCrackWidthEnabled Then
            MissingStateMessage = MissingStateMessage & _
                " Для запрашиваемого сочетания расчет трещин не выполнялся: этому сочетанию назначен профиль, где Calculation.Crack.Width = No."
        ElseIf StrComp(stateType, "CrackedState", vbTextCompare) = 0 Then
            MissingStateMessage = MissingStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, поэтому CrackedState должен быть в snapshot. Проверьте статус расчета трещин и сообщения о сходимости CrackedState."
        Else
            MissingStateMessage = MissingStateMessage & _
                " Для запрашиваемого сочетания расчет трещин был включен, но пороговое состояние образования трещины не записано. Это штатно только когда нет действия для поиска нормальной трещины, например чистое сжатие без момента. Если растяжение или изгиб есть, проверьте статус проверки образования трещины."
        End If
    End If
End Function

' Проверяет, относится ли запрошенный StateType к расчету трещин. Для таких
' состояний можно дать пользователю более точную причину отсутствия snapshot.
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
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then GoTo InvalidValue
    If VarType(value) = vbBoolean Then
        ReadSnapshotBoolean = CBool(value)
    Else
        Select Case LCase$(Trim$(CStr(value)))
            Case "true", "yes", "да", "1"
                ReadSnapshotBoolean = True
            Case "false", "no", "нет", "0"
                ReadSnapshotBoolean = False
            Case Else
                GoTo InvalidValue
        End Select
    End If
    Exit Function
InvalidValue:
    Err.Raise vbObjectError + 4720, "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
End Function

' Проверяет сохраненный display-статус через единый словарь policy. Reader
' не назначает инженерный результат: неизвестный текст означает испорченный
' snapshot, а не новый CalcErr или N/A для найденного состояния.
Private Function ReadSnapshotStatus(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    Dim text As String, policy As CResultStatusPolicy
    text = Trim$(ReadSnapshotText(value, anchor, row, column, field))
    Set policy = New CResultStatusPolicy
    ReadSnapshotStatus = policy.ToUserStatus(text)
    If Len(text) = 0 Or StrComp(text, ReadSnapshotStatus, vbTextCompare) <> 0 Then _
        Err.Raise vbObjectError + 4720, "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
End Function

' Читает служебный текст с адресной диагностикой Excel-ошибки/Null. Пустая
' строка здесь допустима: обязательность определяет потребитель поля.
Private Function ReadSnapshotText(ByVal value As Variant, ByVal anchor As Object, _
        ByVal row As Long, ByVal column As Long, ByVal field As String) As String
    If IsError(value) Or IsNull(value) Then Err.Raise vbObjectError + 4720, _
        "CSectionPlotDataReader", SnapshotMetadataInputMessage(anchor, row, column, field)
    ReadSnapshotText = CStr(value)
End Function

' Выбирает единицу из строки либо метаданных snapshot и проверяет ее в
' CUnitSystem. Пустое/"-" допустимо только для необязательной Unit-ячейки;
' Excel-ошибка не является отсутствием единицы. Новых коэффициентов здесь нет.
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

' Формирует причину повреждения по текущему якорю Results. Рекомендация
' относится к восстановлению снимка, а не к изменению текущих input units.
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

' Выбирает рабочую единицу для числового поля Results. Пустая Unit-ячейка
' или "-" означает единицу из metadata snapshot, а не из текущего Config.
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
    ' Деление сохраняет прежний порядок операций для 1/m, без нового округления.
    OutputCurvatureToInternalByUnit = value / mUnits.InternalCurvatureToOutput(1#, unitText)
End Function

Public Property Get Count() As Long: Count = mCount: End Property
Public Property Get ElementID(ByVal index As Long) As String: ElementID = mElementID(index): End Property
Public Property Get MaterialType(ByVal index As Long) As String: MaterialType = mMaterialType(index): End Property
Public Property Get X(ByVal index As Long) As Double: X = mX(index): End Property
Public Property Get Y(ByVal index As Long) As Double: Y = mY(index): End Property
Public Property Get Area(ByVal index As Long) As Double: Area = mArea(index): End Property
Public Property Get Diameter(ByVal index As Long) As Double: Diameter = mDiameter(index): End Property
Public Property Get ShapeType(ByVal index As Long) As String: ShapeType = mShapeType(index): End Property
Public Property Get Width(ByVal index As Long) As Double: Width = mWidth(index): End Property
Public Property Get Height(ByVal index As Long) As Double: Height = mHeight(index): End Property
Public Property Get Rotation(ByVal index As Long) As Double: Rotation = mRotation(index): End Property
Public Property Get ResultValue(ByVal index As Long) As Double: ResultValue = mResultValue(index): End Property
Public Property Get PhysicalState(ByVal index As Long) As String: PhysicalState = mPhysicalState(index): End Property
Public Property Get LoadCase() As String: LoadCase = mLoadCase: End Property
Public Property Get LoadCaseComment() As String: LoadCaseComment = mLoadCaseComment: End Property
Public Property Get ProfileId() As String: ProfileId = mProfileId: End Property
Public Property Get StateType() As String: StateType = mStateType: End Property
Public Property Get GeometryOnly() As Boolean: GeometryOnly = mGeometryOnly: End Property
Public Property Get GeometryOnlyTitle() As String: GeometryOnlyTitle = mGeometryOnlyTitle: End Property
Public Property Get GeometryOnlyNote() As String: GeometryOnlyNote = mGeometryOnlyNote: End Property
Public Property Get GeometryReferenceItemsEnabled() As Boolean: GeometryReferenceItemsEnabled = mGeometryReferenceItemsEnabled: End Property
Public Property Get VisualizationQuantity() As String: VisualizationQuantity = mVisualizationQuantity: End Property
Public Property Get ResultUnit() As String: ResultUnit = mResultUnit: End Property
Public Property Get ResultCaption() As String: ResultCaption = mResultCaption: End Property
Public Property Get ResultPrecision() As Long: ResultPrecision = mResultPrecision: End Property
Public Property Get LengthUnit() As String: LengthUnit = mLengthUnit: End Property
Public Property Get AreaUnit() As String: AreaUnit = mAreaUnit: End Property
Public Property Get MinX() As Double: MinX = mMinX: End Property
Public Property Get MaxX() As Double: MaxX = mMaxX: End Property
Public Property Get MinY() As Double: MinY = mMinY: End Property
Public Property Get MaxY() As Double: MaxY = mMaxY: End Property
Public Property Get CentroidX() As Double: CentroidX = mCentroidX: End Property
Public Property Get CentroidY() As Double: CentroidY = mCentroidY: End Property
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
Public Property Get Epsilon0() As Double: Epsilon0 = mEpsilon0: End Property
Public Property Get KappaX() As Double: KappaX = mKappaX: End Property
Public Property Get KappaY() As Double: KappaY = mKappaY: End Property
Public Property Get LoadReferenceX() As Double: LoadReferenceX = mLoadReferenceX: End Property
Public Property Get LoadReferenceY() As Double: LoadReferenceY = mLoadReferenceY: End Property
Public Property Get ExtensionUsed() As Boolean: ExtensionUsed = mExtensionUsed: End Property
Public Property Get DirectStateStatus() As String: DirectStateStatus = mStateStatus: End Property
' Определяет необходимость предупреждения по сохраненному снимку Results:
' использовано extension либо записан физический FAIL. Решатель не запускается,
' внешний словарь статусов остается у общей policy.
Public Property Get PhysicalDiagramWarningRequired() As Boolean
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy
    PhysicalDiagramWarningRequired = (mExtensionUsed Or policy.ToUserStatus(mStateStatus) = policy.Fail)
End Property
Public Property Get AnnotationCount() As Long: AnnotationCount = mAnnotationCount: End Property
Public Property Get AnnotationType(ByVal index As Long) As String: AnnotationType = mAnnotationType(index): End Property
Public Property Get AnnotationStartX(ByVal index As Long) As Double: AnnotationStartX = mAnnotationStartX(index): End Property
Public Property Get AnnotationStartY(ByVal index As Long) As Double: AnnotationStartY = mAnnotationStartY(index): End Property
Public Property Get AnnotationEndX(ByVal index As Long) As Double: AnnotationEndX = mAnnotationEndX(index): End Property
Public Property Get AnnotationEndY(ByVal index As Long) As Double: AnnotationEndY = mAnnotationEndY(index): End Property
Public Property Get AnnotationNormalX(ByVal index As Long) As Double: AnnotationNormalX = mAnnotationNormalX(index): End Property
Public Property Get AnnotationNormalY(ByVal index As Long) As Double: AnnotationNormalY = mAnnotationNormalY(index): End Property
Public Property Get AnnotationText(ByVal index As Long) As String: AnnotationText = mAnnotationText(index): End Property
' Отдает готовый проверенный угол; plotter не разбирает текст Results повторно.
Public Property Get AnnotationSweepAngle(ByVal index As Long) As Double: AnnotationSweepAngle = mAnnotationSweepAngle(index): End Property

