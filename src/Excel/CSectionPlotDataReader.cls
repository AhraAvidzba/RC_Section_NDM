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

' Загружает все данные, необходимые для Excel-схемы.
' Метод намеренно читает расчетные значения только из Results. Из Config
' используется только выбор LC и текущие настройки визуализации профиля.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    ResetReaderState
    mLoadCase = ResolveLoadCase(workbook, settings.GetString("Plot.LoadCase", "Worst"))
    ApplyVisualizationProfile workbook

    ReadSectionProperties workbook
    ReadGeometryAndResults workbook
    ReadAnnotations workbook
End Sub

' Загружает геометрию последнего AutoCAD import для визуального контроля.
' Расчетные результаты и нейтральная линия не читаются, но сохраненные
' габаритные аннотации по Bounds можно показать сразу после импорта.
Public Sub LoadGeometryPreviewFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    ResetReaderState
    mGeometryOnly = True
    mGeometryOnlyTitle = "Импортированная геометрия AutoCAD"
    mVisualizationQuantity = "Stress"
    mLoadCase = "AutoCAD import"

    ReadSectionProperties workbook
    mResultCaption = "Геометрия"
    mResultUnit = "-"
    ReadGeometryPreview workbook
    ReadAnnotations workbook
End Sub

' Загружает расчетную геометрию из Results, когда выбранное состояние НДС
' отсутствует в snapshot. Это не AutoCAD-preview и не повторный расчет:
' пользователь видит актуальную геометрию последнего запуска и понятную
' подпись, почему цветовая раскраска Stress/Strain недоступна.
Public Sub LoadGeometryOnlyForMissingState(ByVal workbook As Object, _
        ByVal settings As CSystemSettingsReader, Optional ByVal noteText As String = vbNullString)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    ResetReaderState
    mLoadCase = ResolveLoadCase(workbook, settings.GetString("Plot.LoadCase", "Worst"))
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
End Sub

' Определяет, какой LC показывать на схеме.
' Значение Worst разворачивается в LC из верхней строки batch summary
' "Определяющее сочетание". Критерий выбора задает
' General.WorstCombinationCriterion, а схема только читает сохраненный результат.
Private Function ResolveLoadCase(ByVal workbook As Object, ByVal requested As String) As String
    Dim valueText As String
    valueText = Trim$(requested)
    If Len(valueText) = 0 Then valueText = "Worst"

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

    Dim rowOffset As Long
    For rowOffset = 11 To 1000
        If StrComp(Trim$(SafeText(anchor.Offset(rowOffset, 0).Value2)), loadCase, vbTextCompare) = 0 Then
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
    Dim data As Variant
    data = ReadAnchoredResultTable(workbook, "rngNDMSectionProperties")
    If Not HasRows(data) Then Err.Raise vbObjectError + 4703, "CSectionPlotDataReader", _
        "На листе Results нет rngNDMSectionProperties. Сначала выполните расчет."

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    ' В одной таблице лежат свойства ALL и свойства конкретных LC. Reader
    ' выбирает только общие строки и строки mLoadCase, остальные LC игнорируются.
    For rowIndex = 2 To UBound(data, 1)
        Dim parameter As String
        Dim loadCase As String
        parameter = Trim$(CStr(data(rowIndex, colParameter)))
        loadCase = Trim$(CStr(data(rowIndex, colLoadCase)))
        If Len(parameter) > 0 Then
            If StrComp(loadCase, "ALL", vbTextCompare) = 0 Or Len(loadCase) = 0 Then
                ReadCommonProperty parameter, data(rowIndex, colValue), data(rowIndex, ResultColumn(data, "Unit"))
            ElseIf StrComp(loadCase, mLoadCase, vbTextCompare) = 0 Then
                ReadLoadCaseProperty parameter, data(rowIndex, colValue), data(rowIndex, ResultColumn(data, "Unit"))
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
Private Sub ReadCommonProperty(ByVal parameter As String, ByVal value As Variant, ByVal unitValue As Variant)
    Dim unitText As String
    unitText = Trim$(SafeText(unitValue))
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
        Case "loadreferencex": mLoadReferenceX = OutputLengthToInternalByUnit(SafeDouble(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "loadreferencey": mLoadReferenceY = OutputLengthToInternalByUnit(SafeDouble(value), UnitOrDefault(unitText, mSnapshotLengthUnit, "mm"))
        Case "output.lengthunit": mSnapshotLengthUnit = CStr(value)
        Case "output.areaunit": mSnapshotAreaUnit = CStr(value)
        Case "output.stressunit": mStressUnit = CStr(value)
        Case "output.curvatureunit": mSnapshotCurvatureUnit = CStr(value)
    End Select
End Sub

' Запоминает свойства выбранного LC: плоскость деформаций и признаки
' допустимости найденного состояния. Точка приложения нагрузки читается из
' общего блока ALL, потому что она задается один раз для всего snapshot.
Private Sub ReadLoadCaseProperty(ByVal parameter As String, ByVal value As Variant, ByVal unitValue As Variant)
    Dim key As String
    Dim prefix As String
    key = LCase$(Trim$(parameter))
    prefix = "state." & LCase$(mStateType) & "."
    If Left$(key, Len(prefix)) <> prefix Then Exit Sub

    Dim unitText As String
    unitText = Trim$(SafeText(unitValue))
    Select Case Mid$(key, Len(prefix) + 1)
        Case "epsilon0": mEpsilon0 = SafeDouble(value)
        Case "kappax": mKappaX = OutputCurvatureToInternalByUnit(SafeDouble(value), UnitOrDefault(unitText, mSnapshotCurvatureUnit, "1/mm"))
        Case "kappay": mKappaY = OutputCurvatureToInternalByUnit(SafeDouble(value), UnitOrDefault(unitText, mSnapshotCurvatureUnit, "1/mm"))
        Case "extensionused": mExtensionUsed = SafeBoolean(value)
        Case "status": mStateStatus = Trim$(SafeText(value))
    End Select
End Sub

' Соединяет постоянную геометрию с LC-зависимыми результатами.
' rngNDMSectionGeometry хранит координаты и площади один раз, а
' rngNDMElementResults хранит Stress/Strain/PhysicalState по LoadCase.
Private Sub ReadGeometryAndResults(ByVal workbook As Object)
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."

    Dim results As Variant
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
            resultMap(CStr(results(rowIndex, colResID))) = Array(CDbl(results(rowIndex, colStrain)), _
                CDbl(results(rowIndex, colStress)), CStr(results(rowIndex, colState)))
        End If
    Next rowIndex
    If resultMap.Count = 0 Then Err.Raise vbObjectError + 4706, "CSectionPlotDataReader", _
        MissingStateMessage(mLoadCase, mStateType)

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colShape As Long: colShape = GeometryStatusColumn(geom)
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
        If Len(elementID) > 0 And resultMap.Exists(elementID) Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), OutputLengthToInternalByUnit(CDbl(geom(rowIndex, colX)), lengthUnit), _
                OutputLengthToInternalByUnit(CDbl(geom(rowIndex, colY)), lengthUnit), OutputAreaToInternalByUnit(CDbl(geom(rowIndex, colArea)), areaUnit), _
                OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colDiameter)), lengthUnit), CStr(geom(rowIndex, colShape)), _
                OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colWidth)), lengthUnit), OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colHeight)), lengthUnit), _
                SafeDouble(geom(rowIndex, colRotation)), resultMap(elementID)
        End If
    Next rowIndex
End Sub

' Читает только постоянную геометрию элементов и добавляет ее в модель схемы
' с нейтральным состоянием. Этот режим используется для AutoCAD-preview и для
' схемы без выбранного StateType, если геометрия уже есть в Results.
Private Sub ReadGeometryPreview(ByVal workbook As Object)
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colShape As Long: colShape = GeometryStatusColumn(geom)
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
        If Len(elementID) > 0 Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), OutputLengthToInternalByUnit(CDbl(geom(rowIndex, colX)), lengthUnit), _
                OutputLengthToInternalByUnit(CDbl(geom(rowIndex, colY)), lengthUnit), OutputAreaToInternalByUnit(CDbl(geom(rowIndex, colArea)), areaUnit), _
                OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colDiameter)), lengthUnit), CStr(geom(rowIndex, colShape)), _
                OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colWidth)), lengthUnit), OutputLengthToInternalByUnit(SafeDouble(geom(rowIndex, colHeight)), lengthUnit), _
                SafeDouble(geom(rowIndex, colRotation)), neutralResult
        End If
    Next rowIndex
End Sub

' Добавляет элемент схемы во внутренние массивы reader-а. Значение результата
' выбирается по текущему режиму визуализации: Stress или Strain.
Private Sub AppendElement(ByVal elementID As String, ByVal materialType As String, ByVal xCoord As Double, _
        ByVal yCoord As Double, ByVal area As Double, ByVal diameter As Double, _
        ByVal shapeType As String, ByVal width As Double, ByVal height As Double, _
        ByVal rotation As Double, ByVal resultData As Variant)
    ' Сохраняем только элементы, у которых есть и геометрия, и результат
    ' выбранного LC. Так схема не рисует “пустые” элементы после ошибочного LC.
    mCount = mCount + 1
    ReDim Preserve mElementID(1 To mCount)
    ReDim Preserve mMaterialType(1 To mCount)
    ReDim Preserve mX(1 To mCount)
    ReDim Preserve mY(1 To mCount)
    ReDim Preserve mArea(1 To mCount)
    ReDim Preserve mDiameter(1 To mCount)
    ReDim Preserve mShapeType(1 To mCount)
    ReDim Preserve mWidth(1 To mCount)
    ReDim Preserve mHeight(1 To mCount)
    ReDim Preserve mRotation(1 To mCount)
    ReDim Preserve mResultValue(1 To mCount)
    ReDim Preserve mPhysicalState(1 To mCount)

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

' Читает semantic-аннотации схемы.
' Если таблица отсутствует или пуста, схема все равно строится: для AutoCAD
' import это нормальный сценарий, потому что характерные грани могут быть неизвестны.
Private Sub ReadAnnotations(ByVal workbook As Object)
    On Error GoTo NoAnnotations
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
        annotationType = Trim$(CStr(data(rowIndex, colType)))
        If Len(annotationType) > 0 Then
            Dim lengthUnit As String
            lengthUnit = AnnotationLengthUnit(data, rowIndex, colUnit, mSnapshotLengthUnit)
            mAnnotationCount = mAnnotationCount + 1
            ReDim Preserve mAnnotationType(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationText(1 To mAnnotationCount)
            mAnnotationType(mAnnotationCount) = annotationType
            mAnnotationStartX(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colStartX)), lengthUnit)
            mAnnotationStartY(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colStartY)), lengthUnit)
            mAnnotationEndX(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colEndX)), lengthUnit)
            mAnnotationEndY(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colEndY)), lengthUnit)
            If StrComp(annotationType, "CONTOUR_ARC", vbTextCompare) = 0 Then
                mAnnotationNormalX(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colNormalX)), lengthUnit)
                mAnnotationNormalY(mAnnotationCount) = OutputLengthToInternalByUnit(SafeDouble(data(rowIndex, colNormalY)), lengthUnit)
            Else
                mAnnotationNormalX(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalX))
                mAnnotationNormalY(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalY))
            End If
            mAnnotationText(mAnnotationCount) = CStr(data(rowIndex, colText))
        End If
    Next rowIndex
NoAnnotations:
End Sub

' Сбрасывает внутренние массивы перед новым чтением Results.
' Reader может переиспользоваться тестами, поэтому нельзя оставлять элементы
' и аннотации от предыдущей загрузки.
Private Sub ResetReaderState()
    mGeometryOnly = False
    mGeometryOnlyTitle = vbNullString
    mGeometryOnlyNote = vbNullString
    mGeometryReferenceItemsEnabled = False
    mVisualizationQuantity = vbNullString
    mLoadCase = vbNullString
    mLoadCaseComment = vbNullString
    mProfileId = vbNullString
    mStateType = vbNullString
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
End Sub

' Читает таблицу Results от именованного якоря, не полагаясь на CurrentRegion.
' Это позволяет размещать человекочитаемый заголовок прямо над таблицей и не
' включать его в массив данных.
Private Function ReadAnchoredResultTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    Dim anchor As Object
    Set anchor = workbook.Names.Item(rangeName).RefersToRange

    Dim columnCount As Long
    columnCount = AnchoredColumnCount(anchor)
    If columnCount <= 0 Then Exit Function

    Dim rowCount As Long
    rowCount = AnchoredRowCount(anchor)
    If rowCount <= 0 Then rowCount = 1

    ReadAnchoredResultTable = anchor.Resize(rowCount, columnCount).Value2
End Function

' Определяет ширину таблицы по непрерывной строке заголовков от якоря.
Private Function AnchoredColumnCount(ByVal anchor As Object) As Long
    Dim colOffset As Long
    For colOffset = 0 To 255
        If Len(Trim$(CStr(anchor.Offset(0, colOffset).Value2))) = 0 Then Exit For
        AnchoredColumnCount = AnchoredColumnCount + 1
    Next colOffset
End Function

' Определяет высоту таблицы по первому столбцу. Пустой первый столбец означает
' конец данных; это соответствует формату всех Results-диапазонов.
Private Function AnchoredRowCount(ByVal anchor As Object) As Long
    Dim rowOffset As Long
    For rowOffset = 0 To 1048575 - anchor.Row
        If Len(Trim$(CStr(anchor.Offset(rowOffset, 0).Value2))) = 0 Then Exit For
        AnchoredRowCount = AnchoredRowCount + 1
    Next rowOffset
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

' Безопасно читает числовое поле Results, где часть колонок может быть пустой.
' Например Diameter у бетонных элементов не заполняется.
Private Function SafeDouble(ByVal value As Variant) As Double
    If IsError(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If IsNumeric(value) Then SafeDouble = CDbl(value)
End Function

' Читает логический флаг из Results. Snapshot может хранить True/False как
' настоящий Boolean или как текст после ручного просмотра/копирования.
Private Function SafeBoolean(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If VarType(value) = vbBoolean Then
        SafeBoolean = CBool(value)
    Else
        Select Case LCase$(Trim$(CStr(value)))
            Case "true", "yes", "да", "1"
                SafeBoolean = True
        End Select
    End If
End Function

' Превращает значение ячейки в строку без падения на Excel-ошибках.
Private Function SafeText(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    SafeText = CStr(value)
End Function

' Возвращает единицу измерения из строки аннотации. Если строка не содержит
' единицу или хранит служебный "-", используется единица длины всего snapshot.
Private Function AnnotationLengthUnit(ByRef data As Variant, ByVal rowIndex As Long, _
        ByVal colUnit As Long, ByVal defaultUnit As String) As String
    AnnotationLengthUnit = UnitOrDefault(SafeText(data(rowIndex, colUnit)), defaultUnit, "mm")
End Function

' Выбирает рабочую единицу для числового поля Results. Пустые и служебные
' значения не должны ломать старые snapshots, где часть Unit-колонки пустая.
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

' Читает GeometryInterpretationStatus; если этого столбца нет в сохраненном
' Results, допускает ShapeType как входной формат ранее записанного снимка.
Private Function GeometryStatusColumn(ByRef data As Variant) As Long
    On Error Resume Next
    GeometryStatusColumn = ResultColumn(data, "GeometryInterpretationStatus")
    If Err.Number = 0 And GeometryStatusColumn > 0 Then
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    GeometryStatusColumn = ResultColumn(data, "ShapeType")
    On Error GoTo 0
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
    OutputLengthToInternalByUnit = value * LengthFactorToMmByUnit(unitText)
End Function

' Переводит площадь из единицы, сохраненной в Results, во внутренние мм2.
Private Function OutputAreaToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    OutputAreaToInternalByUnit = value * AreaFactorToMm2ByUnit(unitText)
End Function

' Переводит кривизну из единицы Results во внутренние 1/мм, чтобы нейтральная
' линия строилась в той же системе, что и координаты схемы.
Private Function OutputCurvatureToInternalByUnit(ByVal value As Double, ByVal unitText As String) As Double
    Select Case LCase$(Trim$(unitText))
        Case "1/mm"
            OutputCurvatureToInternalByUnit = value
        Case "1/m"
            OutputCurvatureToInternalByUnit = value / 1000#
        Case Else
            Err.Raise vbObjectError + 4715, "CSectionPlotDataReader", _
                "В Results неизвестная единица кривизны: " & unitText
    End Select
End Function

' Возвращает множитель перевода длины в мм.
Private Function LengthFactorToMmByUnit(ByVal unitText As String) As Double
    Select Case LCase$(Trim$(unitText))
        Case "mm"
            LengthFactorToMmByUnit = 1#
        Case "cm"
            LengthFactorToMmByUnit = 10#
        Case "m"
            LengthFactorToMmByUnit = 1000#
        Case Else
            Err.Raise vbObjectError + 4716, "CSectionPlotDataReader", _
                "В Results неизвестная единица длины: " & unitText
    End Select
End Function

' Возвращает множитель перевода площади в мм2.
Private Function AreaFactorToMm2ByUnit(ByVal unitText As String) As Double
    Select Case LCase$(Trim$(unitText))
        Case "mm2"
            AreaFactorToMm2ByUnit = 1#
        Case "cm2"
            AreaFactorToMm2ByUnit = 100#
        Case "m2"
            AreaFactorToMm2ByUnit = 1000000#
        Case Else
            Err.Raise vbObjectError + 4717, "CSectionPlotDataReader", _
                "В Results неизвестная единица площади: " & unitText
    End Select
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

