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
' Reader не запускает расчет, не читает текущую геометрию Config и не
' пересчитывает единицы: схема показывает ровно то, что записано в Results.

Private mResultType As String ' Что сейчас визуализируется: Stress или Strain.
Private mLoadCase As String ' LC, выбранный для схемы после обработки Worst/конкретного ID.
Private mLoadCaseComment As String ' Комментарий LC для заголовка схемы.
Private mGeometryOnly As Boolean ' True для preview импортированной геометрии без расчетных результатов.
Private mResultUnit As String ' Единица выбранной величины из сохраненных Results.
Private mResultCaption As String ' Название выбранной величины для легенды.
Private mLengthUnit As String ' Единица длины, в которой записаны координаты Results.
Private mAreaUnit As String ' Единица площади, в которой записаны площади Results.
Private mStressUnit As String ' Единица напряжений, сохраненная при последнем расчете.
Private mCurvatureUnit As String ' Единица кривизн, сохраненная при последнем расчете.
Private mCount As Long ' Количество элементов, у которых есть геометрия и результат выбранного LC.
Private mElementID() As String ' ID элемента для связи геометрии, результатов и подписей.
Private mMaterialType() As String ' Concrete/Rebar; Plotter использует это для выбора маркера.
Private mX() As Double ' X элемента в output-единицах Results.
Private mY() As Double ' Y элемента в output-единицах Results.
Private mArea() As Double ' Площадь элемента в output-единицах Results.
Private mDiameter() As Double ' Диаметр арматуры; для бетона обычно 0.
Private mWidth() As Double ' Ширина эквивалентного бетонного маркера/области.
Private mHeight() As Double ' Высота эквивалентного бетонного маркера/области.
Private mResultValue() As Double ' Значение Stress или Strain для выбранного LC.
Private mPhysicalState() As String ' Compression/Tension/InactiveTensionConcrete/NearZero.
Private mMinX As Double ' Левая граница бетонной расчетной модели.
Private mMaxX As Double ' Правая граница бетонной расчетной модели.
Private mMinY As Double ' Нижняя граница бетонной расчетной модели.
Private mMaxY As Double ' Верхняя граница бетонной расчетной модели.
Private mCentroidX As Double ' X центра тяжести приведенного сечения.
Private mCentroidY As Double ' Y центра тяжести приведенного сечения.
Private mPrincipalAngle As Double ' Угол главных центральных осей, рад.
Private mEpsilon0 As Double ' Свободный член сохраненной плоскости деформаций LC.
Private mKappaX As Double ' Кривизна при y в сохраненной плоскости деформаций.
Private mKappaY As Double ' Кривизна при x в сохраненной плоскости деформаций.
Private mLoadReferenceX As Double ' X точки приложения нагрузки для выбранного LC.
Private mLoadReferenceY As Double ' Y точки приложения нагрузки для выбранного LC.
Private mExtensionUsed As Boolean ' True, если выбранный LC был найден в numerical extension.
Private mDirectStateStatus As String ' Сохраненный статус прямого НДС выбранного LC.
Private mAnnotationCount As Long ' Количество semantic-аннотаций, прочитанных из Results.
Private mAnnotationType() As String ' Dimension, RebarLabel, Outline и другие типы оформления.
Private mAnnotationStartX() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationStartY() As Double ' Начало линии/размера в координатах модели.
Private mAnnotationEndX() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationEndY() As Double ' Конец линии/размера в координатах модели.
Private mAnnotationNormalX() As Double ' Наружная нормаль semantic-аннотации.
Private mAnnotationNormalY() As Double ' Наружная нормаль semantic-аннотации.
Private mAnnotationOffset() As Double ' Конструктивный/semantic offset, если он задан builder-ом.
Private mAnnotationText() As String ' Готовый текст подписи или размера.

' Загружает все данные, необходимые для Excel-схемы.
' Метод намеренно читает только Results и настройки отображения Plot.*:
' геометрия и расчетные значения берутся из последнего выполненного расчета.
Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки Config не переданы."

    ResetReaderState
    mResultType = NormalizeResultType(settings.GetString("Plot.ResultType", "Stress"))
    mLoadCase = ResolveLoadCase(workbook, settings.GetString("Plot.LoadCase", "Worst"))

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
    mResultType = "Stress"
    mLoadCase = "AutoCAD import"

    ReadSectionProperties workbook
    mResultCaption = "Геометрия"
    mResultUnit = "-"
    ReadGeometryPreview workbook
    ReadAnnotations workbook
End Sub

' Определяет, какой LC показывать на схеме.
' Значение Worst разворачивается в LC из верхней строки batch summary
' "Определяющее сочетание по прочности". Эта строка формируется batch-слоем
' по прочностному правилу: минимальный CapacitySafetyFactor среди рассчитанных
' сочетаний. Общий MinSafetyFactor остается только справочной колонкой.
Private Function ResolveLoadCase(ByVal workbook As Object, ByVal requested As String) As String
    Dim valueText As String
    valueText = Trim$(requested)
    If Len(valueText) = 0 Then valueText = "Worst"

    If StrComp(valueText, "Worst", vbTextCompare) <> 0 Then
        ResolveLoadCase = valueText
        mLoadCaseComment = ReadBatchSummaryComment(workbook, ResolveLoadCase)
        Exit Function
    End If

    Dim data As Variant
    data = workbook.Names.Item("rngBatchSummary").RefersToRange.Value2
    If UBound(data, 1) >= 2 And UBound(data, 2) >= 5 Then
        ResolveLoadCase = Trim$(SafeText(data(2, 5)))
    End If
    If Len(ResolveLoadCase) = 0 Then ResolveLoadCase = FirstCalculatedLoadCaseFromResults(workbook)
    mLoadCaseComment = ReadBatchSummaryComment(workbook, ResolveLoadCase)
    If Len(ResolveLoadCase) = 0 Then Err.Raise vbObjectError + 4702, "CSectionPlotDataReader", _
        "В rngBatchSummary не найдено определяющее сочетание по прочности для Plot.LoadCase = Worst."
End Function

' Возвращает первый LC, для которого на Results реально есть результаты элементов.
' Это запасной путь только для режимов, где Plot.LoadCase = Worst не может быть
' развернут в прочностное определяющее сочетание, например DirectState без
' расчета несущей способности. Старый AutoCAD-preview этим путем не выбирается,
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

' Находит комментарий выбранного LC для заголовка схемы.
' Если комментария нет или batch summary еще пустой, возвращается пустая строка
' без ошибки: сама схема может строиться и без комментария.
Private Function ReadBatchSummaryComment(ByVal workbook As Object, ByVal loadCase As String) As String
    On Error GoTo Failed
    If Len(Trim$(loadCase)) = 0 Then Exit Function

    Dim data As Variant
    data = workbook.Names.Item("rngBatchSummary").RefersToRange.Value2

    Dim rowIndex As Long
    For rowIndex = 10 To UBound(data, 1)
        If StrComp(Trim$(SafeText(data(rowIndex, 2))), loadCase, vbTextCompare) = 0 Then
            ReadBatchSummaryComment = Trim$(SafeText(data(rowIndex, 3)))
            Exit Function
        End If
    Next rowIndex
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
                ReadCommonProperty parameter, data(rowIndex, colValue)
            ElseIf StrComp(loadCase, mLoadCase, vbTextCompare) = 0 Then
                ReadLoadCaseProperty parameter, data(rowIndex, colValue)
            End If
        End If
    Next rowIndex

    If Len(mLengthUnit) = 0 Then mLengthUnit = "mm"
    If Len(mAreaUnit) = 0 Then mAreaUnit = mLengthUnit & "2"
    If Len(mStressUnit) = 0 Then mStressUnit = "MPa"
    If Len(mCurvatureUnit) = 0 Then mCurvatureUnit = "1/mm"
    If StrComp(mResultType, "Strain", vbTextCompare) = 0 Then
        mResultUnit = "-"
        mResultCaption = "Strain"
    Else
        mResultUnit = mStressUnit
        mResultCaption = "Stress"
    End If
End Sub

' Запоминает общие свойства расчетного снимка Results.
' Эти данные одинаково подходят для Generated и AutoCAD-геометрии.
Private Sub ReadCommonProperty(ByVal parameter As String, ByVal value As Variant)
    Select Case LCase$(parameter)
        Case "bounds.minx": mMinX = CDbl(value)
        Case "bounds.maxx": mMaxX = CDbl(value)
        Case "bounds.miny": mMinY = CDbl(value)
        Case "bounds.maxy": mMaxY = CDbl(value)
        Case "centroidx": mCentroidX = CDbl(value)
        Case "centroidy": mCentroidY = CDbl(value)
        Case "principalangle": mPrincipalAngle = CDbl(value)
        Case "output.lengthunit": mLengthUnit = CStr(value)
        Case "output.areaunit": mAreaUnit = CStr(value)
        Case "output.stressunit": mStressUnit = CStr(value)
        Case "output.curvatureunit": mCurvatureUnit = CStr(value)
    End Select
End Sub

' Запоминает свойства выбранного LC: плоскость деформаций и точку нагрузки.
' Нейтральная линия затем строится из epsilon0/kappaX/kappaY без хранения
' отдельного A/B/C.
Private Sub ReadLoadCaseProperty(ByVal parameter As String, ByVal value As Variant)
    Select Case LCase$(parameter)
        Case "epsilon0": mEpsilon0 = CDbl(value)
        Case "kappax": mKappaX = CDbl(value)
        Case "kappay": mKappaY = CDbl(value)
        Case "loadreferencex": mLoadReferenceX = CDbl(value)
        Case "loadreferencey": mLoadReferenceY = CDbl(value)
        Case "extensionused": mExtensionUsed = SafeBoolean(value)
        Case "directstatestatus": mDirectStateStatus = Trim$(SafeText(value))
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
    Dim colResID As Long: colResID = ResultColumn(results, "ElementID")
    Dim colStrain As Long: colStrain = ResultColumn(results, "Strain")
    Dim colStress As Long: colStress = ResultColumn(results, "Stress")
    Dim colState As Long: colState = ResultColumn(results, "PhysicalState")

    Dim rowIndex As Long
    ' LC-зависимые результаты сначала собираются в словарь по ElementID. Это
    ' сохраняет разделение Results: геометрия лежит отдельно, а Stress/Strain
    ' выбираются только для нужного LoadCase.
    For rowIndex = 2 To UBound(results, 1)
        If StrComp(CStr(results(rowIndex, colResLC)), mLoadCase, vbTextCompare) = 0 Then
            resultMap(CStr(results(rowIndex, colResID))) = Array(CDbl(results(rowIndex, colStrain)), _
                CDbl(results(rowIndex, colStress)), CStr(results(rowIndex, colState)))
        End If
    Next rowIndex
    If resultMap.Count = 0 Then Err.Raise vbObjectError + 4706, "CSectionPlotDataReader", _
        "В Results нет данных для Plot.LoadCase: " & mLoadCase

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colDiameter As Long: colDiameter = ResultColumn(geom, "Diameter")
    Dim colWidth As Long: colWidth = ResultColumn(geom, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(geom, "Height")

    ' На втором проходе геометрия соединяется с найденным состоянием. Plotter после этого
    ' получает готовые элементы и не обращается к Results напрямую.
    For rowIndex = 2 To UBound(geom, 1)
        Dim elementID As String
        elementID = Trim$(CStr(geom(rowIndex, colID)))
        If Len(elementID) > 0 And resultMap.Exists(elementID) Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), CDbl(geom(rowIndex, colX)), _
                CDbl(geom(rowIndex, colY)), CDbl(geom(rowIndex, colArea)), _
                SafeDouble(geom(rowIndex, colDiameter)), SafeDouble(geom(rowIndex, colWidth)), _
                SafeDouble(geom(rowIndex, colHeight)), resultMap(elementID)
        End If
    Next rowIndex
End Sub

' Читает только постоянную геометрию элементов и добавляет ее в модель схемы
' с нейтральным состоянием. Этот режим нужен сразу после AutoCAD import.
Private Sub ReadGeometryPreview(ByVal workbook As Object)
    Dim geom As Variant
    geom = ReadAnchoredResultTable(workbook, "rngNDMSectionGeometry")
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет импортированной геометрии."

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colDiameter As Long: colDiameter = ResultColumn(geom, "Diameter")
    Dim colWidth As Long: colWidth = ResultColumn(geom, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(geom, "Height")

    Dim neutralResult As Variant
    neutralResult = Array(0#, 0#, "NearZero")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(geom, 1)
        Dim elementID As String
        elementID = Trim$(CStr(geom(rowIndex, colID)))
        If Len(elementID) > 0 Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), CDbl(geom(rowIndex, colX)), _
                CDbl(geom(rowIndex, colY)), CDbl(geom(rowIndex, colArea)), _
                SafeDouble(geom(rowIndex, colDiameter)), SafeDouble(geom(rowIndex, colWidth)), _
                SafeDouble(geom(rowIndex, colHeight)), neutralResult
        End If
    Next rowIndex
End Sub

Private Sub AppendElement(ByVal elementID As String, ByVal materialType As String, ByVal xCoord As Double, _
        ByVal yCoord As Double, ByVal area As Double, ByVal diameter As Double, _
        ByVal width As Double, ByVal height As Double, ByVal resultData As Variant)
    ' Сохраняем только элементы, у которых есть и геометрия, и результат
    ' выбранного LC. Так схема не рисует “пустые” элементы после ошибочного LC.
    mCount = mCount + 1
    ReDim Preserve mElementID(1 To mCount)
    ReDim Preserve mMaterialType(1 To mCount)
    ReDim Preserve mX(1 To mCount)
    ReDim Preserve mY(1 To mCount)
    ReDim Preserve mArea(1 To mCount)
    ReDim Preserve mDiameter(1 To mCount)
    ReDim Preserve mWidth(1 To mCount)
    ReDim Preserve mHeight(1 To mCount)
    ReDim Preserve mResultValue(1 To mCount)
    ReDim Preserve mPhysicalState(1 To mCount)

    mElementID(mCount) = elementID
    mMaterialType(mCount) = materialType
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mArea(mCount) = area
    mDiameter(mCount) = diameter
    mWidth(mCount) = width
    mHeight(mCount) = height
    If StrComp(mResultType, "Strain", vbTextCompare) = 0 Then
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
    Dim colOffset As Long: colOffset = ResultColumn(data, "Offset")
    Dim colText As Long: colText = ResultColumn(data, "Text")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Len(Trim$(CStr(data(rowIndex, colType)))) > 0 Then
            mAnnotationCount = mAnnotationCount + 1
            ReDim Preserve mAnnotationType(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationOffset(1 To mAnnotationCount)
            ReDim Preserve mAnnotationText(1 To mAnnotationCount)
            mAnnotationType(mAnnotationCount) = CStr(data(rowIndex, colType))
            mAnnotationStartX(mAnnotationCount) = SafeDouble(data(rowIndex, colStartX))
            mAnnotationStartY(mAnnotationCount) = SafeDouble(data(rowIndex, colStartY))
            mAnnotationEndX(mAnnotationCount) = SafeDouble(data(rowIndex, colEndX))
            mAnnotationEndY(mAnnotationCount) = SafeDouble(data(rowIndex, colEndY))
            mAnnotationNormalX(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalX))
            mAnnotationNormalY(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalY))
            mAnnotationOffset(mAnnotationCount) = SafeDouble(data(rowIndex, colOffset))
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
    mResultType = vbNullString
    mLoadCase = vbNullString
    mLoadCaseComment = vbNullString
    mResultUnit = vbNullString
    mResultCaption = vbNullString
    mLengthUnit = vbNullString
    mAreaUnit = vbNullString
    mStressUnit = vbNullString
    mCurvatureUnit = vbNullString
    mCount = 0
    mAnnotationCount = 0
    mMinX = 0#: mMaxX = 0#: mMinY = 0#: mMaxY = 0#
    mCentroidX = 0#: mCentroidY = 0#: mPrincipalAngle = 0#
    mEpsilon0 = 0#: mKappaX = 0#: mKappaY = 0#
    mLoadReferenceX = 0#: mLoadReferenceY = 0#
    mExtensionUsed = False
    mDirectStateStatus = vbNullString

    Erase mElementID
    Erase mMaterialType
    Erase mX
    Erase mY
    Erase mArea
    Erase mDiameter
    Erase mWidth
    Erase mHeight
    Erase mResultValue
    Erase mPhysicalState
    Erase mAnnotationType
    Erase mAnnotationStartX
    Erase mAnnotationStartY
    Erase mAnnotationEndX
    Erase mAnnotationEndY
    Erase mAnnotationNormalX
    Erase mAnnotationNormalY
    Erase mAnnotationOffset
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

Private Function NormalizeResultType(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "stress": NormalizeResultType = "Stress"
        Case "strain": NormalizeResultType = "Strain"
        Case Else: Err.Raise vbObjectError + 4707, "CSectionPlotDataReader", "Plot.ResultType должен быть Stress или Strain."
    End Select
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

Public Property Get Count() As Long: Count = mCount: End Property
Public Property Get ElementID(ByVal index As Long) As String: ElementID = mElementID(index): End Property
Public Property Get MaterialType(ByVal index As Long) As String: MaterialType = mMaterialType(index): End Property
Public Property Get X(ByVal index As Long) As Double: X = mX(index): End Property
Public Property Get Y(ByVal index As Long) As Double: Y = mY(index): End Property
Public Property Get Area(ByVal index As Long) As Double: Area = mArea(index): End Property
Public Property Get Diameter(ByVal index As Long) As Double: Diameter = mDiameter(index): End Property
Public Property Get Width(ByVal index As Long) As Double: Width = mWidth(index): End Property
Public Property Get Height(ByVal index As Long) As Double: Height = mHeight(index): End Property
Public Property Get ResultValue(ByVal index As Long) As Double: ResultValue = mResultValue(index): End Property
Public Property Get PhysicalState(ByVal index As Long) As String: PhysicalState = mPhysicalState(index): End Property
Public Property Get LoadCase() As String: LoadCase = mLoadCase: End Property
Public Property Get LoadCaseComment() As String: LoadCaseComment = mLoadCaseComment: End Property
Public Property Get GeometryOnly() As Boolean: GeometryOnly = mGeometryOnly: End Property
Public Property Get ResultType() As String: ResultType = mResultType: End Property
Public Property Get ResultUnit() As String: ResultUnit = mResultUnit: End Property
Public Property Get ResultCaption() As String: ResultCaption = mResultCaption: End Property
Public Property Get LengthUnit() As String: LengthUnit = mLengthUnit: End Property
Public Property Get AreaUnit() As String: AreaUnit = mAreaUnit: End Property
Public Property Get MinX() As Double: MinX = mMinX: End Property
Public Property Get MaxX() As Double: MaxX = mMaxX: End Property
Public Property Get MinY() As Double: MinY = mMinY: End Property
Public Property Get MaxY() As Double: MaxY = mMaxY: End Property
Public Property Get CentroidX() As Double: CentroidX = mCentroidX: End Property
Public Property Get CentroidY() As Double: CentroidY = mCentroidY: End Property
Public Property Get PrincipalAngle() As Double: PrincipalAngle = mPrincipalAngle: End Property
Public Property Get Epsilon0() As Double: Epsilon0 = mEpsilon0: End Property
Public Property Get KappaX() As Double: KappaX = mKappaX: End Property
Public Property Get KappaY() As Double: KappaY = mKappaY: End Property
Public Property Get LoadReferenceX() As Double: LoadReferenceX = mLoadReferenceX: End Property
Public Property Get LoadReferenceY() As Double: LoadReferenceY = mLoadReferenceY: End Property
Public Property Get ExtensionUsed() As Boolean: ExtensionUsed = mExtensionUsed: End Property
Public Property Get DirectStateStatus() As String: DirectStateStatus = mDirectStateStatus: End Property
Public Property Get PhysicalDiagramWarningRequired() As Boolean
    PhysicalDiagramWarningRequired = (mExtensionUsed Or StrComp(mDirectStateStatus, "FAIL", vbTextCompare) = 0)
End Property
Public Property Get AnnotationCount() As Long: AnnotationCount = mAnnotationCount: End Property
Public Property Get AnnotationType(ByVal index As Long) As String: AnnotationType = mAnnotationType(index): End Property
Public Property Get AnnotationStartX(ByVal index As Long) As Double: AnnotationStartX = mAnnotationStartX(index): End Property
Public Property Get AnnotationStartY(ByVal index As Long) As Double: AnnotationStartY = mAnnotationStartY(index): End Property
Public Property Get AnnotationEndX(ByVal index As Long) As Double: AnnotationEndX = mAnnotationEndX(index): End Property
Public Property Get AnnotationEndY(ByVal index As Long) As Double: AnnotationEndY = mAnnotationEndY(index): End Property
Public Property Get AnnotationNormalX(ByVal index As Long) As Double: AnnotationNormalX = mAnnotationNormalX(index): End Property
Public Property Get AnnotationNormalY(ByVal index As Long) As Double: AnnotationNormalY = mAnnotationNormalY(index): End Property
Public Property Get AnnotationOffset(ByVal index As Long) As Double: AnnotationOffset = mAnnotationOffset(index): End Property
Public Property Get AnnotationText(ByVal index As Long) As String: AnnotationText = mAnnotationText(index): End Property

