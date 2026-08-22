VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionPlotter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Excel-схема расчетного сечения
' ==========================================================================
' CSectionPlotter рисует схему по данным, которые уже прочитал
' CSectionPlotDataReader из листа Results. Он не запускает solver, не смотрит
' текущие исходные данные Config и не пересчитывает единицы: схема всегда
' показывает сохраненный результат последнего расчета.
'
' Расчетные объекты (бетон, арматура, нейтральная линия, главные оси и точка
' нагрузки) выводятся через XY-series. Shapes используются только для
' оформительских аннотаций: размеров, выносных линий, подписей арматуры и легенд.

Private Const CHART_NAME As String = "chtNDMSectionPlot"
Private Const SHAPE_PREFIX As String = "NDMPlot_"
Private Const XL_XY_SCATTER As Long = -4169
Private Const XL_XY_SCATTER_LINES As Long = 74
Private Const XL_XY_SCATTER_LINES_NO_MARKERS As Long = 75
Private Const XL_CATEGORY As Long = 1
Private Const XL_VALUE As Long = 2
Private Const XL_PRIMARY As Long = 1
Private Const XL_MARKER_SQUARE As Long = 1
Private Const XL_MARKER_DIAMOND As Long = 2
Private Const XL_MARKER_CIRCLE As Long = 8
Private Const XL_NONE As Long = -4142
Private Const XL_CENTER As Long = -4108
Private Const XL_VCENTER As Long = -4108
Private Const MSO_TEXT_ORIENTATION_HORIZONTAL As Long = 1
Private Const MSO_LINE_SOLID As Long = 1
Private Const MSO_LINE_DASH As Long = 4
Private Const MSO_LINE_DASH_DOT As Long = 5
Private Const MSO_ARROWHEAD_TRIANGLE As Long = 2
Private Const MSO_ARROWHEAD_STEALTH As Long = 3
Private Const MSO_ARROWHEAD_DIAMOND As Long = 4
Private Const MSO_ARROWHEAD_OVAL As Long = 5
Private Const MSO_ARROWHEAD_OPEN As Long = 6
Private Const MSO_ARROWHEAD_LENGTH_SHORT As Long = 1
Private Const MSO_ARROWHEAD_LENGTH_MEDIUM As Long = 2
Private Const MSO_ARROWHEAD_LENGTH_LONG As Long = 3
Private Const MSO_ARROWHEAD_WIDTH_NARROW As Long = 1
Private Const MSO_ARROWHEAD_WIDTH_MEDIUM As Long = 2
Private Const MSO_ARROWHEAD_WIDTH_WIDE As Long = 3
Private Const MSO_SHAPE_RECTANGLE As Long = 1
Private Const MSO_SHAPE_DIAMOND As Long = 4
Private Const MSO_SHAPE_OVAL As Long = 9

Private mChart As Object ' Excel Chart, внутри которого рисуются series и аннотационные Shapes.
Private mChartObject As Object ' Контейнер ChartObject на листе Расчет; пользователь может двигать/масштабировать его.
Private mMinX As Double ' Минимум оси X после добавления полей для аннотаций.
Private mMaxX As Double ' Максимум оси X после добавления полей для аннотаций.
Private mMinY As Double ' Минимум оси Y после добавления полей для аннотаций.
Private mMaxY As Double ' Максимум оси Y после добавления полей для аннотаций.
Private mResultMin As Double ' Минимальное значение выбранного ResultType среди рабочих элементов.
Private mResultMax As Double ' Максимальное значение выбранного ResultType среди рабочих элементов.
Private mCompressionMaxAbs As Double ' Максимальный модуль сжатия для общей легенды.
Private mCompressionValueAtMaxAbs As Double ' Знаковое значение, подписываемое у края шкалы сжатия.
Private mTensionMaxAbs As Double ' Максимальный модуль растяжения для общей легенды.
Private mTensionValueAtMaxAbs As Double ' Знаковое значение, подписываемое у края шкалы растяжения.
Private mResultMaxAbs As Double ' Резервный максимум по модулю, если состояние не разделено.
Private mHasCompression As Boolean ' Есть ли элементы в сжатии для общей легенды.
Private mHasTension As Boolean ' Есть ли элементы в растяжении для общей легенды.
Private mConcreteCompressionMaxAbs As Double ' Максимальный модуль сжатия бетона для отдельной легенды.
Private mConcreteCompressionValueAtMaxAbs As Double ' Подпись края бетонной шкалы сжатия.
Private mConcreteTensionMaxAbs As Double ' Максимальный модуль растяжения бетона, если оно учитывается.
Private mConcreteTensionValueAtMaxAbs As Double ' Подпись края бетонной шкалы растяжения.
Private mConcreteHasCompression As Boolean ' В расчетном состоянии есть сжатый бетон.
Private mConcreteHasTension As Boolean ' В расчетном состоянии есть растянутый работающий бетон.
Private mRebarCompressionMaxAbs As Double ' Максимальный модуль сжатия арматуры.
Private mRebarCompressionValueAtMaxAbs As Double ' Подпись края арматурной шкалы сжатия.
Private mRebarTensionMaxAbs As Double ' Максимальный модуль растяжения арматуры.
Private mRebarTensionValueAtMaxAbs As Double ' Подпись края арматурной шкалы растяжения.
Private mRebarHasCompression As Boolean ' В расчетном состоянии есть сжатая арматура.
Private mRebarHasTension As Boolean ' В расчетном состоянии есть растянутая арматура.
Private mResultGradientEnabled As Boolean ' Включена ли цветовая шкала по Stress/Strain.
Private mResultPrecision As Long ' Число знаков после запятой для подписей результатов.
Private mDimensionArrowType As String ' Тип стрелок размерных линий из Config.
Private mDimensionArrowSize As String ' Размер стрелок размерных линий из Config.
Private mLegendMode As String ' Separate или Common для легенд бетона/арматуры.
Private mRebarCompressionColor As Long ' Цвет максимального сжатия арматуры.
Private mRebarTensionColor As Long ' Цвет максимального растяжения арматуры.
Private mConcreteCompressionColor As Long ' Цвет максимального сжатия бетона.
Private mConcreteTensionColor As Long ' Цвет максимального растяжения бетона.

' Перестраивает схему по данным последнего расчета.
' Метод управляет порядком операций: подготовка chart, равный масштаб X/Y,
' series расчетных элементов, расчетные линии и Shape-аннотации.
Public Sub Draw(ByVal workbook As Object, ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4800, "CSectionPlotter", "Книга Excel не передана."
    If data Is Nothing Then Err.Raise vbObjectError + 4801, "CSectionPlotter", "Данные схемы не переданы."
    If settings Is Nothing Then Err.Raise vbObjectError + 4802, "CSectionPlotter", "Настройки Config не переданы."

    Set mChartObject = EnsureChart(workbook.Worksheets.Item("Расчет"))
    Set mChart = mChartObject.Chart

    mResultGradientEnabled = settings.GetBoolean("Plot.ResultGradient", True)
    mResultPrecision = settings.GetLong("Plot.ResultPrecision", 2)
    If mResultPrecision < 0 Then mResultPrecision = 0
    If mResultPrecision > 8 Then mResultPrecision = 8
    mDimensionArrowType = NormalizeArrowType(settings.GetString("Plot.Dimensions.ArrowType", "Triangle"))
    mDimensionArrowSize = NormalizeArrowSize(settings.GetString("Plot.Dimensions.ArrowSize", "Wide"))
    mLegendMode = NormalizeLegendMode(settings.GetString("Plot.LegendMode", "Separate"))
    mRebarCompressionColor = ParseRgb(settings.GetString("Plot.Color.RebarCompression", "30,80,220"), RGB(30, 80, 220))
    mRebarTensionColor = ParseRgb(settings.GetString("Plot.Color.RebarTension", "210,30,20"), RGB(210, 30, 20))
    mConcreteCompressionColor = ParseRgb(settings.GetString("Plot.Color.ConcreteCompression", "125,35,210"), RGB(125, 35, 210))
    mConcreteTensionColor = ParseRgb(settings.GetString("Plot.Color.ConcreteTension", "0,155,85"), RGB(0, 155, 85))

    PrepareChart data
    ConfigureEqualScale data, settings
    ClearGeneratedShapes
    CalculateResultRange data
    AddElementSeries data
    DrawAnalysisSeries data, settings
    DrawShapeDecorations data, settings
End Sub

' Возвращает существующий ChartObject или создает его при первом запуске.
' При последующих обновлениях контейнер не пересоздается, чтобы пользователь
' мог вручную перемещать и менять размер схемы на листе.
Private Function EnsureChart(ByVal sheet As Object) As Object
    Dim chartObject As Object
    On Error Resume Next
    Set chartObject = sheet.ChartObjects(CHART_NAME)
    On Error GoTo 0

    If chartObject Is Nothing Then
        Dim anchor As Object
        Set anchor = sheet.Range("AP1")
        Set chartObject = sheet.ChartObjects.Add(anchor.Left, anchor.Top, 820, 540)
        chartObject.Name = CHART_NAME
    End If
    Set EnsureChart = chartObject
End Function

' Подготавливает Chart к новой отрисовке, сохраняя сам ChartObject.
' Старые series удаляются, но размеры и положение окна схемы остаются как у пользователя.
Private Sub PrepareChart(ByVal data As CSectionPlotDataReader)
    mChart.ChartType = XL_XY_SCATTER
    mChart.HasLegend = False
    mChart.HasTitle = True
    mChart.ChartTitle.Text = PlotTitleText(data)
    mChart.ChartTitle.Font.Size = 12

    Do While mChart.SeriesCollection.Count > 0
        mChart.SeriesCollection(1).Delete
    Loop

    On Error Resume Next
    mChart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    mChart.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    mChart.PlotArea.Left = 28
    mChart.PlotArea.Top = 36
    mChart.PlotArea.Width = mChartObject.Width - 180
    mChart.PlotArea.Height = mChartObject.Height - 72
    On Error GoTo 0
End Sub

' Настраивает одинаковый физический масштаб X и Y.
' Одна единица модели должна занимать одинаковое расстояние по горизонтали и
' вертикали, иначе круг станет эллипсом, а расположение арматуры исказится.
Private Sub ConfigureEqualScale(ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    Dim rawMinX As Double, rawMaxX As Double, rawMinY As Double, rawMaxY As Double
    rawMinX = data.MinX
    rawMaxX = data.MaxX
    rawMinY = data.MinY
    rawMaxY = data.MaxY

    Dim dx As Double, dy As Double
    dx = rawMaxX - rawMinX
    dy = rawMaxY - rawMinY
    If dx <= 0# Then dx = 1#
    If dy <= 0# Then dy = 1#

    Dim annotationMargin As Double
    annotationMargin = AnnotationMarginModel(settings, data.LengthUnit)

    rawMinX = rawMinX - MaxDouble(dx * 0.16, annotationMargin)
    rawMaxX = rawMaxX + MaxDouble(dx * 0.24, annotationMargin)
    rawMinY = rawMinY - MaxDouble(dy * 0.18, annotationMargin)
    rawMaxY = rawMaxY + MaxDouble(dy * 0.16, annotationMargin)

    Dim plotW As Double, plotH As Double
    plotW = mChart.PlotArea.InsideWidth
    plotH = mChart.PlotArea.InsideHeight
    If plotW <= 0# Or plotH <= 0# Then
        plotW = mChartObject.Width * 0.72
        plotH = mChartObject.Height * 0.72
    End If

    Dim rangeX As Double, rangeY As Double
    rangeX = rawMaxX - rawMinX
    rangeY = rawMaxY - rawMinY

    Dim targetRangeX As Double, targetRangeY As Double
    If rangeX / plotW > rangeY / plotH Then
        targetRangeX = rangeX
        targetRangeY = rangeX * plotH / plotW
    Else
        targetRangeY = rangeY
        targetRangeX = rangeY * plotW / plotH
    End If

    Dim centerX As Double, centerY As Double
    centerX = 0.5 * (rawMinX + rawMaxX)
    centerY = 0.5 * (rawMinY + rawMaxY)
    mMinX = centerX - targetRangeX / 2#
    mMaxX = centerX + targetRangeX / 2#
    mMinY = centerY - targetRangeY / 2#
    mMaxY = centerY + targetRangeY / 2#

    mChart.Axes(XL_CATEGORY).MinimumScale = mMinX
    mChart.Axes(XL_CATEGORY).MaximumScale = mMaxX
    mChart.Axes(XL_VALUE).MinimumScale = mMinY
    mChart.Axes(XL_VALUE).MaximumScale = mMaxY
    HideChartAxes
End Sub

Private Function AnnotationMarginModel(ByVal settings As CSystemSettingsReader, ByVal lengthUnit As String) As Double
    If settings Is Nothing Then Exit Function

    Dim marginMm As Double
    If settings.GetBoolean("Plot.Dimensions.Enabled", True) Then
        marginMm = MaxDouble(marginMm, settings.GetDouble("Plot.Dimensions.Offset", 100#) + _
            settings.GetDouble("Plot.Dimensions.TextGap", 15#) + _
            settings.GetDouble("Plot.Dimensions.TextHeight", 25#))
    End If
    If settings.GetBoolean("Plot.RebarLabels.Enabled", True) Then
        marginMm = MaxDouble(marginMm, settings.GetDouble("Plot.RebarLabels.Offset", 60#) + _
            settings.GetDouble("Plot.RebarLabels.TextGap", 15#) + _
            settings.GetDouble("Plot.RebarLabels.TextHeight", 25#))
    End If

    AnnotationMarginModel = marginMm / LengthFactorToMm(lengthUnit)
End Function

Private Sub HideChartAxes()
    On Error Resume Next
    With mChart.Axes(XL_CATEGORY)
        .TickLabelPosition = XL_NONE
        .MajorTickMark = XL_NONE
        .MinorTickMark = XL_NONE
        .Format.Line.Visible = False
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    With mChart.Axes(XL_VALUE)
        .TickLabelPosition = XL_NONE
        .MajorTickMark = XL_NONE
        .MinorTickMark = XL_NONE
        .Format.Line.Visible = False
        .HasMajorGridlines = False
        .HasMinorGridlines = False
    End With
    On Error GoTo 0
End Sub

Private Sub CalculateResultRange(ByVal data As CSectionPlotDataReader)
    Dim found As Boolean
    Dim i As Long
    mHasCompression = False
    mHasTension = False
    mCompressionMaxAbs = 0#
    mTensionMaxAbs = 0#
    mResultMaxAbs = 0#
    ResetMaterialLegendRanges
    For i = 1 To data.Count
        If StrComp(data.PhysicalState(i), "InactiveTensionConcrete", vbTextCompare) <> 0 Then
            If Not found Then
                mResultMin = data.ResultValue(i)
                mResultMax = data.ResultValue(i)
                found = True
            Else
                mResultMin = MinDouble(mResultMin, data.ResultValue(i))
                mResultMax = MaxDouble(mResultMax, data.ResultValue(i))
            End If

            If StrComp(data.PhysicalState(i), "Compression", vbTextCompare) = 0 Then
                RegisterResultMagnitude data.ResultValue(i), mCompressionMaxAbs, mCompressionValueAtMaxAbs, mHasCompression
                RegisterMaterialLegendMagnitude data.MaterialType(i), True, data.ResultValue(i)
                mResultMaxAbs = MaxDouble(mResultMaxAbs, Abs(data.ResultValue(i)))
            ElseIf StrComp(data.PhysicalState(i), "Tension", vbTextCompare) = 0 Then
                RegisterResultMagnitude data.ResultValue(i), mTensionMaxAbs, mTensionValueAtMaxAbs, mHasTension
                RegisterMaterialLegendMagnitude data.MaterialType(i), False, data.ResultValue(i)
                mResultMaxAbs = MaxDouble(mResultMaxAbs, Abs(data.ResultValue(i)))
            End If
        End If
    Next i
    If Not found Or Abs(mResultMax - mResultMin) < 0.000000000001 Then
        mResultMin = -1#
        mResultMax = 1#
    End If
End Sub

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Private Sub ResetMaterialLegendRanges()
    mConcreteCompressionMaxAbs = 0#
    mConcreteTensionMaxAbs = 0#
    mConcreteHasCompression = False
    mConcreteHasTension = False
    mRebarCompressionMaxAbs = 0#
    mRebarTensionMaxAbs = 0#
    mRebarHasCompression = False
    mRebarHasTension = False
End Sub

Private Sub RegisterMaterialLegendMagnitude(ByVal materialType As String, ByVal compression As Boolean, ByVal value As Double)
    If StrComp(materialType, "Concrete", vbTextCompare) = 0 Then
        If compression Then
            RegisterResultMagnitude value, mConcreteCompressionMaxAbs, mConcreteCompressionValueAtMaxAbs, mConcreteHasCompression
        Else
            RegisterResultMagnitude value, mConcreteTensionMaxAbs, mConcreteTensionValueAtMaxAbs, mConcreteHasTension
        End If
    ElseIf StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        If compression Then
            RegisterResultMagnitude value, mRebarCompressionMaxAbs, mRebarCompressionValueAtMaxAbs, mRebarHasCompression
        Else
            RegisterResultMagnitude value, mRebarTensionMaxAbs, mRebarTensionValueAtMaxAbs, mRebarHasTension
        End If
    End If
End Sub

Private Sub RegisterResultMagnitude(ByVal value As Double, ByRef maxAbs As Double, _
        ByRef valueAtMaxAbs As Double, ByRef found As Boolean)
    Dim absValue As Double
    absValue = Abs(value)
    If Not found Or absValue > maxAbs Then
        maxAbs = absValue
        valueAtMaxAbs = value
        found = True
    End If
End Sub

' Группирует расчетные элементы в серии Excel по типу маркера, размеру и цвету.
' Это сильно ускоряет схему: сотни бетонных волокон не создают сотни отдельных
' series, но визуально остаются отдельными точками.
Private Sub AddElementSeries(ByVal data As CSectionPlotDataReader)
    Dim buckets As Object
    Set buckets = CreateObject("Scripting.Dictionary")
    buckets.CompareMode = vbTextCompare

    Dim i As Long
    ' Пространственная сетка подписей не зависит от ElementID. В каждой ячейке выбирается
    ' ближайшее бетонное волокно к центру ячейки, чтобы подписи распределялись равномерно.
    For i = 1 To data.Count
        Dim key As String
        key = ElementBucketKey(data, i)
        Dim indexList As Collection
        If buckets.Exists(key) Then
            Set indexList = buckets.Item(key)
        Else
            Set indexList = New Collection
            buckets.Add key, indexList
        End If
        indexList.Add i
    Next i

    Dim keyVariant As Variant
    For Each keyVariant In buckets.Keys
        AddBucketSeriesByIndexes data, CStr(keyVariant), buckets(keyVariant)
    Next keyVariant
End Sub

Private Function ElementBucketKey(ByVal data As CSectionPlotDataReader, ByVal index As Long) As String
    Dim markerStyle As Long
    Dim markerSize As Long
    If StrComp(data.MaterialType(index), "Rebar", vbTextCompare) = 0 Then
        markerStyle = XL_MARKER_CIRCLE
        markerSize = RoundToLong(ClampDouble(data.Diameter(index) * ChartScale(), 4#, 42#))
    Else
        markerStyle = XL_MARKER_SQUARE
        markerSize = RoundToLong(ClampDouble(0.7 * ConcreteMarkerLength(data, index) * ChartScale(), 2#, 48#))
    End If

    Dim colorValue As Long
    colorValue = ElementColor(data, index)
    ElementBucketKey = CStr(markerStyle) & "|" & CStr(markerSize) & "|" & CStr(colorValue) & "|" & _
        data.MaterialType(index) & "|" & data.PhysicalState(index)
End Function

' Создает одну XY-series для группы одинаково оформленных элементов.
' Координаты элементов остаются model/output-координатами, а Excel сам
' размещает маркеры по настроенным осям.
Private Sub AddBucketSeriesByIndexes(ByVal data As CSectionPlotDataReader, ByVal bucketKey As String, ByVal indexes As Collection)
    If indexes Is Nothing Then Exit Sub
    If indexes.Count <= 0 Then Exit Sub

    Dim parts As Variant
    parts = Split(bucketKey, "|")
    Dim markerStyle As Long: markerStyle = CLng(parts(0))
    Dim markerSize As Long: markerSize = CLng(parts(1))
    Dim colorValue As Long: colorValue = CLng(parts(2))

    Dim xValues() As Double
    Dim yValues() As Double
    ReDim xValues(1 To indexes.Count)
    ReDim yValues(1 To indexes.Count)

    Dim i As Long, index As Long
    For i = 1 To indexes.Count
        index = CLng(indexes.Item(i))
        xValues(i) = data.X(index)
        yValues(i) = data.Y(index)
    Next i

    Dim series As Object
    Set series = mChart.SeriesCollection.NewSeries
    series.Name = parts(3) & " " & parts(4)
    series.XValues = xValues
    series.Values = yValues
    series.MarkerStyle = markerStyle
    series.MarkerSize = markerSize
    series.Format.Line.Visible = False
    If UBound(parts) >= 3 And StrComp(CStr(parts(3)), "Rebar", vbTextCompare) = 0 Then
        series.MarkerForegroundColor = RGB(0, 0, 0)
    Else
        series.MarkerForegroundColor = colorValue
    End If
    series.MarkerBackgroundColor = colorValue
End Sub

' Выбирает цвет одного расчетного элемента.
' Цвет берется из PhysicalState и выбранной легенды, а не из пользовательского
' знака Stress/Strain, потому что знак вывода может быть настроен пользователем.
Private Function ElementColor(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Long
    If StrComp(data.PhysicalState(index), "InactiveTensionConcrete", vbTextCompare) = 0 Then
        ElementColor = RGB(218, 218, 218)
    ElseIf StrComp(data.PhysicalState(index), "NearZero", vbTextCompare) = 0 Then
        ElementColor = RGB(150, 150, 150)
    ElseIf mResultGradientEnabled Then
        ElementColor = MaterialResultColor(data.MaterialType(index), data.PhysicalState(index), data.ResultValue(index))
    ElseIf StrComp(data.MaterialType(index), "Rebar", vbTextCompare) = 0 Then
        If StrComp(data.PhysicalState(index), "Compression", vbTextCompare) = 0 Then
            ElementColor = LegendCompressionColor("Rebar")
        Else
            ElementColor = LegendTensionColor("Rebar")
        End If
    ElseIf StrComp(data.PhysicalState(index), "Compression", vbTextCompare) = 0 Then
        ElementColor = LegendCompressionColor("Concrete")
    Else
        ElementColor = LegendTensionColor("Concrete")
    End If
End Function

' Возвращает градиентный цвет внутри отдельного состояния материала.
' Ноль идет к нейтральному серому, максимум по модулю - к цвету сжатия или
' растяжения для бетона/арматуры.
Private Function MaterialResultColor(ByVal materialType As String, ByVal state As String, ByVal value As Double) As Long
    Dim t As Double
    Dim maxAbs As Double
    maxAbs = MaterialStateMaxAbs(materialType, state)
    If maxAbs <= 0.000000000001 Then
        t = 1#
    Else
        t = ClampDouble(Abs(value) / maxAbs, 0#, 1#)
    End If

    If StrComp(state, "Compression", vbTextCompare) = 0 Then
        MaterialResultColor = BlendColor(RGB(218, 218, 218), LegendCompressionColor(materialType), t)
    ElseIf StrComp(state, "Tension", vbTextCompare) = 0 Then
        MaterialResultColor = BlendColor(RGB(218, 218, 218), LegendTensionColor(materialType), t)
    Else
        MaterialResultColor = RGB(150, 150, 150)
    End If
End Function

' Возвращает максимум по модулю для выбранного материала и физического состояния.
' При общей легенде используется общий диапазон, при раздельной - свой диапазон
' для бетона и арматуры.
Private Function MaterialStateMaxAbs(ByVal materialType As String, ByVal state As String) As Double
    If IsCommonLegendMode() Then
        If StrComp(state, "Compression", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mCompressionMaxAbs
        ElseIf StrComp(state, "Tension", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mTensionMaxAbs
        End If
        If MaterialStateMaxAbs <= 0# Then MaterialStateMaxAbs = mResultMaxAbs
        Exit Function
    End If

    If StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        If StrComp(state, "Compression", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mRebarCompressionMaxAbs
        ElseIf StrComp(state, "Tension", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mRebarTensionMaxAbs
        End If
    ElseIf StrComp(materialType, "Concrete", vbTextCompare) = 0 Then
        If StrComp(state, "Compression", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mConcreteCompressionMaxAbs
        ElseIf StrComp(state, "Tension", vbTextCompare) = 0 Then
            MaterialStateMaxAbs = mConcreteTensionMaxAbs
        End If
    End If
    If MaterialStateMaxAbs <= 0# Then MaterialStateMaxAbs = mResultMaxAbs
End Function

' Старый непрерывный градиент для общей шкалы значений.
' Оставлен как утилита цвета; основная раскраска сейчас идет через
' MaterialResultColor, чтобы не смешивать физический смысл сжатия/растяжения.
Private Function GradientColor(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Long
    Dim t As Double
    t = (value - minValue) / (maxValue - minValue)
    t = ClampDouble(t, 0#, 1#)

    If t < 0.25 Then
        GradientColor = BlendColor(RGB(210, 30, 20), RGB(255, 140, 20), t / 0.25)
    ElseIf t < 0.5 Then
        GradientColor = BlendColor(RGB(255, 140, 20), RGB(255, 230, 30), (t - 0.25) / 0.25)
    ElseIf t < 0.75 Then
        GradientColor = BlendColor(RGB(255, 230, 30), RGB(50, 190, 190), (t - 0.5) / 0.25)
    Else
        GradientColor = BlendColor(RGB(50, 190, 190), RGB(30, 80, 220), (t - 0.75) / 0.25)
    End If
End Function

' Линейно смешивает два RGB-цвета.
Private Function BlendColor(ByVal c1 As Long, ByVal c2 As Long, ByVal t As Double) As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    r1 = c1 Mod 256
    g1 = (c1 \ 256) Mod 256
    b1 = (c1 \ 65536) Mod 256
    r2 = c2 Mod 256
    g2 = (c2 \ 256) Mod 256
    b2 = (c2 \ 65536) Mod 256
    BlendColor = RGB(CLng(r1 + (r2 - r1) * t), CLng(g1 + (g2 - g1) * t), CLng(b1 + (b2 - b1) * t))
End Function

' Добавляет расчетные линии и точку нагрузки поверх маркеров элементов.
' Нейтральная линия и главные оси остаются XY-series, а не Shapes.
Private Sub DrawAnalysisSeries(ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    If settings.GetBoolean("Plot.NeutralLineEnabled", True) Then DrawNeutralLine data
    If settings.GetBoolean("Plot.PrincipalAxesEnabled", True) Then DrawPrincipalAxes data
    If settings.GetBoolean("Plot.LoadApplicationPointEnabled", True) Then DrawPointMarker data.LoadReferenceX, data.LoadReferenceY, 10, RGB(255, 128, 0), XL_MARKER_DIAMOND, "LoadPoint"
End Sub

' Рисует Shape-слой схемы: контур, размеры, групповые подписи и легенды.
' Shapes используются только для оформления; расчетные объекты остаются series.
Private Sub DrawShapeDecorations(ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    DrawConcreteContour data
    DrawShapeAnnotations data, settings, settings.GetBoolean("Plot.Dimensions.Enabled", True), _
        settings.GetBoolean("Plot.RebarLabels.Enabled", True)
    If settings.GetBoolean("Plot.ResultLabelsEnabled", False) Then DrawResultLabels data, settings.GetDouble("Plot.ResultLabelSpacing", 100#)
    If settings.GetBoolean("Plot.LegendEnabled", True) Then DrawLegend data
End Sub

' Рисует контур бетона из semantic-аннотаций Results.
' Generated-геометрия обычно передает точный контур; AutoCAD import может не
' иметь таких аннотаций, и тогда контур не угадывается насильно.
Private Sub DrawConcreteContour(ByVal data As CSectionPlotDataReader)
    Dim i As Long
    For i = 1 To data.AnnotationCount
        If StrComp(data.AnnotationType(i), "CONTOUR_LINE", vbTextCompare) = 0 Then
            DrawModelLine data.AnnotationStartX(i), data.AnnotationStartY(i), _
                data.AnnotationEndX(i), data.AnnotationEndY(i), RGB(25, 25, 25), 1.7, "Contour"
        ElseIf StrComp(data.AnnotationType(i), "CONTOUR_CIRCLE", vbTextCompare) = 0 Then
            DrawModelOval data.AnnotationStartX(i) - data.AnnotationEndX(i), _
                data.AnnotationStartY(i) - data.AnnotationEndX(i), _
                data.AnnotationStartX(i) + data.AnnotationEndX(i), _
                data.AnnotationStartY(i) + data.AnnotationEndX(i), RGB(25, 25, 25), 1.7, "ContourCircle", False
        End If
    Next i
End Sub

' Добавляет ребро прямоугольного элемента в набор границ.
' Если соседний элемент уже добавил обратное ребро, оно удаляется: внутренние
' границы сетки не должны становиться наружным контуром.
Private Sub AddBoundaryEdge(ByVal edges As Object, ByVal x1 As Double, ByVal y1 As Double, _
        ByVal x2 As Double, ByVal y2 As Double)
    Dim key As String, reverseKey As String
    key = EdgeKey(x1, y1, x2, y2)
    reverseKey = EdgeKey(x2, y2, x1, y1)
    If edges.Exists(reverseKey) Then
        edges.Remove reverseKey
    ElseIf Not edges.Exists(key) Then
        edges.Add key, Array(x1, y1, x2, y2)
    End If
End Sub

' Формирует устойчивый текстовый ключ ребра с округлением координат.
Private Function EdgeKey(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As String
    EdgeKey = FormatNumberInvariant(RoundTo(x1, 6)) & ":" & FormatNumberInvariant(RoundTo(y1, 6)) & "|" & _
        FormatNumberInvariant(RoundTo(x2, 6)) & ":" & FormatNumberInvariant(RoundTo(y2, 6))
End Function

' Строит нейтральную линию по сохраненной плоскости деформаций LC.
' Линия обрезается по semantic-контуру, а если контура нет - по условному
' контуру бетонных элементов.
Private Sub DrawNeutralLine(ByVal data As CSectionPlotDataReader)
    Dim a As Double, b As Double, c As Double
    a = data.KappaY
    b = data.KappaX
    c = data.Epsilon0
    If Abs(a) + Abs(b) < 0.000000000001 Then Exit Sub

    Dim dx As Double, dy As Double
    dx = -b
    dy = a
    Dim norm As Double
    norm = Sqr(dx * dx + dy * dy)
    If norm <= 0# Then Exit Sub
    dx = dx / norm
    dy = dy / norm

    Dim x0 As Double, y0 As Double
    If Abs(b) > Abs(a) Then
        x0 = 0#
        y0 = -c / b
    Else
        y0 = 0#
        x0 = -c / a
    End If

    Dim capacity As Long
    capacity = data.AnnotationCount * 2 + 8
    If capacity < 8 Then capacity = 8
    Dim values() As Double
    ReDim values(1 To capacity)
    Dim count As Long
    CollectContourIntersections data, a, b, c, x0, y0, dx, dy, values, count
    If count < 2 Then CollectConcreteCenterBoundsIntersections data, a, b, c, x0, y0, dx, dy, values, count
    If count < 2 Then Exit Sub

    Dim minT As Double, maxT As Double, i As Long
    minT = values(1)
    maxT = values(1)
    For i = 2 To count
        minT = MinDouble(minT, values(i))
        maxT = MaxDouble(maxT, values(i))
    Next i
    If maxT <= minT Then Exit Sub

    DrawModelLineSeries x0 + minT * dx, y0 + minT * dy, x0 + maxT * dx, y0 + maxT * dy, _
        RGB(105, 105, 105), 0.7, "NeutralAxis"
End Sub

' Собирает пересечения нейтральной линии с semantic-контуром.
' Для Generated-сечений это основной и наиболее точный путь обрезки линии.
Private Sub CollectContourIntersections(ByVal data As CSectionPlotDataReader, ByVal a As Double, _
        ByVal b As Double, ByVal c As Double, ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByRef values() As Double, ByRef count As Long)
    Dim i As Long
    For i = 1 To data.AnnotationCount
        If StrComp(data.AnnotationType(i), "CONTOUR_LINE", vbTextCompare) = 0 Then
            AddSegmentLineIntersectionT a, b, c, x0, y0, dirX, dirY, _
                data.AnnotationStartX(i), data.AnnotationStartY(i), _
                data.AnnotationEndX(i), data.AnnotationEndY(i), values, count
        ElseIf StrComp(data.AnnotationType(i), "CONTOUR_CIRCLE", vbTextCompare) = 0 Then
            AddCircleLineIntersectionT x0, y0, dirX, dirY, data.AnnotationStartX(i), _
                data.AnnotationStartY(i), data.AnnotationEndX(i), values, count
        End If
    Next i
End Sub

' Резервно обрезает линию по габариту центров бетонных элементов.
' Используется, когда точного semantic-контура нет, например после AutoCAD import.
Private Sub CollectConcreteCenterBoundsIntersections(ByVal data As CSectionPlotDataReader, ByVal a As Double, _
        ByVal b As Double, ByVal c As Double, ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByRef values() As Double, ByRef count As Long)
    Dim found As Boolean
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    Dim i As Long
    For i = 1 To data.Count
        If StrComp(data.MaterialType(i), "Concrete", vbTextCompare) = 0 Then
            If Not found Then
                minX = data.X(i): maxX = data.X(i)
                minY = data.Y(i): maxY = data.Y(i)
                found = True
            Else
                minX = MinDouble(minX, data.X(i))
                maxX = MaxDouble(maxX, data.X(i))
                minY = MinDouble(minY, data.Y(i))
                maxY = MaxDouble(maxY, data.Y(i))
            End If
        End If
    Next i
    If Not found Or maxX <= minX Or maxY <= minY Then Exit Sub

    AddSegmentLineIntersectionT a, b, c, x0, y0, dirX, dirY, minX, minY, maxX, minY, values, count
    AddSegmentLineIntersectionT a, b, c, x0, y0, dirX, dirY, maxX, minY, maxX, maxY, values, count
    AddSegmentLineIntersectionT a, b, c, x0, y0, dirX, dirY, maxX, maxY, minX, maxY, values, count
    AddSegmentLineIntersectionT a, b, c, x0, y0, dirX, dirY, minX, maxY, minX, minY, values, count
End Sub

' Добавляет параметр пересечения бесконечной расчетной линии с одним отрезком контура.
Private Sub AddSegmentLineIntersectionT(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
        ByVal x0 As Double, ByVal y0 As Double, ByVal dirX As Double, ByVal dirY As Double, _
        ByVal sx As Double, ByVal sy As Double, ByVal ex As Double, ByVal ey As Double, _
        ByRef values() As Double, ByRef count As Long)
    Dim denom As Double
    denom = a * (ex - sx) + b * (ey - sy)
    If Abs(denom) < 0.000000000001 Then Exit Sub

    Dim u As Double
    u = -(a * sx + b * sy + c) / denom
    If u < -0.0000001 Or u > 1.0000001 Then Exit Sub

    Dim x As Double, y As Double
    x = sx + u * (ex - sx)
    y = sy + u * (ey - sy)
    AddUniqueLineT (x - x0) * dirX + (y - y0) * dirY, values, count
End Sub

' Добавляет две точки пересечения расчетной линии с круговым контуром.
Private Sub AddCircleLineIntersectionT(ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal centerX As Double, _
        ByVal centerY As Double, ByVal radius As Double, ByRef values() As Double, ByRef count As Long)
    If radius <= 0# Then Exit Sub

    Dim vx As Double, vy As Double, tCenter As Double, d2 As Double, halfChord As Double
    vx = centerX - x0
    vy = centerY - y0
    tCenter = vx * dirX + vy * dirY
    d2 = vx * vx + vy * vy - tCenter * tCenter
    If d2 > radius * radius + 0.0000001 Then Exit Sub
    If d2 < 0# Then d2 = 0#
    halfChord = Sqr(MaxDouble(radius * radius - d2, 0#))
    AddUniqueLineT tCenter - halfChord, values, count
    AddUniqueLineT tCenter + halfChord, values, count
End Sub

' Добавляет параметр t без дублей, чтобы вершины контура не создавали
' нулевые или повторные участки линии.
Private Sub AddUniqueLineT(ByVal t As Double, ByRef values() As Double, ByRef count As Long)
    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - t) < 0.000001 Then Exit Sub
    Next i
    If count >= UBound(values) Then ReDim Preserve values(1 To UBound(values) + 16)
    count = count + 1
    values(count) = t
End Sub

' Находит интервал прохождения линии через эквивалентный прямоугольник элемента.
' Этот путь нужен только для приближенной обрезки по сетке.
Private Sub CollectRectLineInterval(ByVal data As CSectionPlotDataReader, ByVal index As Long, _
        ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal dirX As Double, ByVal dirY As Double, _
        ByVal intervals As Collection)
    Dim w As Double, h As Double
    w = ConcreteElementWidth(data, index)
    h = ConcreteElementHeight(data, index)

    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
    x1 = data.X(index) - w / 2#
    x2 = data.X(index) + w / 2#
    y1 = data.Y(index) - h / 2#
    y2 = data.Y(index) + h / 2#

    Dim values(1 To 8) As Double
    Dim count As Long
    AddRectLineIntersectionValue a, b, c, dirX, dirY, x1, y1, x2, y1, values, count
    AddRectLineIntersectionValue a, b, c, dirX, dirY, x2, y1, x2, y2, values, count
    AddRectLineIntersectionValue a, b, c, dirX, dirY, x2, y2, x1, y2, values, count
    AddRectLineIntersectionValue a, b, c, dirX, dirY, x1, y2, x1, y1, values, count
    If count < 2 Then Exit Sub

    Dim minT As Double, maxT As Double, i As Long
    minT = values(1)
    maxT = values(1)
    For i = 2 To count
        minT = MinDouble(minT, values(i))
        maxT = MaxDouble(maxT, values(i))
    Next i
    If maxT > minT Then intervals.Add Array(minT, maxT)
End Sub

' Добавляет одну координату пересечения линии со стороной прямоугольника элемента.
Private Sub AddRectLineIntersectionValue(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal sx As Double, ByVal sy As Double, _
        ByVal ex As Double, ByVal ey As Double, ByRef values() As Double, ByRef count As Long)
    Dim denom As Double
    denom = a * (ex - sx) + b * (ey - sy)
    If Abs(denom) < 0.000000000001 Then Exit Sub

    Dim u As Double
    u = -(a * sx + b * sy + c) / denom
    If u < -0.0000001 Or u > 1.0000001 Then Exit Sub

    Dim x As Double, y As Double, t As Double
    x = sx + u * (ex - sx)
    y = sy + u * (ey - sy)
    t = x * dirX + y * dirY

    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - t) < 0.0000001 Then Exit Sub
    Next i
    count = count + 1
    values(count) = t
End Sub

' Рисует заранее найденные интервалы расчетной линии.
' Близкие интервалы объединяются, чтобы линия выглядела цельной на грубой сетке.
Private Sub DrawLineIntervals(ByVal intervals As Collection, ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal colorValue As Long, ByVal weight As Double, _
        ByVal suffix As String)
    If intervals Is Nothing Then Exit Sub
    If intervals.Count <= 0 Then Exit Sub

    Dim starts() As Double, ends() As Double
    ReDim starts(1 To intervals.Count)
    ReDim ends(1 To intervals.Count)

    Dim i As Long
    For i = 1 To intervals.Count
        Dim item As Variant
        item = intervals.Item(i)
        starts(i) = CDbl(item(0))
        ends(i) = CDbl(item(1))
    Next i
    SortIntervals starts, ends

    Dim gapTol As Double
    gapTol = MaxDouble(mMaxX - mMinX, mMaxY - mMinY) * 0.004

    Dim curStart As Double, curEnd As Double
    curStart = starts(1)
    curEnd = ends(1)
    For i = 2 To UBound(starts)
        If starts(i) <= curEnd + gapTol Then
            curEnd = MaxDouble(curEnd, ends(i))
        Else
            DrawModelLine x0 + curStart * dirX, y0 + curStart * dirY, _
                x0 + curEnd * dirX, y0 + curEnd * dirY, colorValue, weight, suffix
            curStart = starts(i)
            curEnd = ends(i)
        End If
    Next i
    DrawModelLine x0 + curStart * dirX, y0 + curStart * dirY, _
        x0 + curEnd * dirX, y0 + curEnd * dirY, colorValue, weight, suffix
End Sub

' Сортирует интервалы по начальной координате вдоль расчетной линии.
Private Sub SortIntervals(ByRef starts() As Double, ByRef ends() As Double)
    Dim i As Long, j As Long
    For i = LBound(starts) To UBound(starts) - 1
        For j = i + 1 To UBound(starts)
            If starts(j) < starts(i) Then
                SwapDouble starts(i), starts(j)
                SwapDouble ends(i), ends(j)
            End If
        Next j
    Next i
End Sub

' Меняет местами два числа при сортировке интервалов.
Private Sub SwapDouble(ByRef a As Double, ByRef b As Double)
    Dim t As Double
    t = a
    a = b
    b = t
End Sub

' Собирает min/max параметр линии внутри прямоугольного элемента.
Private Sub CollectRectLineIntersections(ByVal data As CSectionPlotDataReader, ByVal index As Long, _
        ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal dirX As Double, ByVal dirY As Double, _
        ByRef minT As Double, ByRef maxT As Double, ByRef found As Boolean)
    Dim w As Double, h As Double
    w = ConcreteElementWidth(data, index)
    h = ConcreteElementHeight(data, index)

    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
    x1 = data.X(index) - w / 2#
    x2 = data.X(index) + w / 2#
    y1 = data.Y(index) - h / 2#
    y2 = data.Y(index) + h / 2#

    AddRectLineIntersection a, b, c, dirX, dirY, x1, y1, x2, y1, minT, maxT, found
    AddRectLineIntersection a, b, c, dirX, dirY, x2, y1, x2, y2, minT, maxT, found
    AddRectLineIntersection a, b, c, dirX, dirY, x2, y2, x1, y2, minT, maxT, found
    AddRectLineIntersection a, b, c, dirX, dirY, x1, y2, x1, y1, minT, maxT, found
End Sub

' Обновляет min/max параметр линии по пересечению с одной стороной прямоугольника.
Private Sub AddRectLineIntersection(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal sx As Double, ByVal sy As Double, _
        ByVal ex As Double, ByVal ey As Double, ByRef minT As Double, ByRef maxT As Double, ByRef found As Boolean)
    Dim denom As Double
    denom = a * (ex - sx) + b * (ey - sy)
    If Abs(denom) < 0.000000000001 Then Exit Sub

    Dim u As Double
    u = -(a * sx + b * sy + c) / denom
    If u < -0.0000001 Or u > 1.0000001 Then Exit Sub

    Dim x As Double, y As Double, t As Double
    x = sx + u * (ex - sx)
    y = sy + u * (ey - sy)
    t = x * dirX + y * dirY
    If t < minT Then minT = t
    If t > maxT Then maxT = t
    found = True
End Sub

' Рисует главные центральные оси приведенного сечения как XY-series.
' Это дает корректный порядок поверх бетонных элементов и сохраняет единый
' model-to-chart масштаб.
Private Sub DrawPrincipalAxes(ByVal data As CSectionPlotDataReader)
    Dim lengthValue As Double
    lengthValue = 0.65 * Sqr((data.MaxX - data.MinX) * (data.MaxX - data.MinX) + _
        (data.MaxY - data.MinY) * (data.MaxY - data.MinY))
    Dim ca As Double, sa As Double
    ca = Cos(data.PrincipalAngle)
    sa = Sin(data.PrincipalAngle)
    DrawModelLineSeries data.CentroidX - lengthValue * ca, data.CentroidY - lengthValue * sa, _
        data.CentroidX + lengthValue * ca, data.CentroidY + lengthValue * sa, RGB(0, 0, 0), 0.7, "Principal1", MSO_LINE_DASH
    DrawModelLineSeries data.CentroidX + lengthValue * sa, data.CentroidY - lengthValue * ca, _
        data.CentroidX - lengthValue * sa, data.CentroidY + lengthValue * ca, RGB(0, 0, 0), 0.7, "Principal2", MSO_LINE_DASH
End Sub

' Передает semantic-аннотации Results в CPlotAnnotationLayout.
' Здесь нет логики конкретной формы: если аннотация есть, layout размещает ее;
' если import не дал semantic-данных, Plotter ничего не угадывает.
Private Sub DrawShapeAnnotations(ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader, _
        ByVal dimensionsEnabled As Boolean, ByVal rebarLabelsEnabled As Boolean)
    If Not dimensionsEnabled And Not rebarLabelsEnabled Then Exit Sub

    Dim layout As CPlotAnnotationLayout
    Set layout = New CPlotAnnotationLayout
    layout.ConfigureFromSettings settings
    layout.Initialize mMinX, mMaxX, mMinY, mMaxY, mChart.PlotArea.InsideLeft, mChart.PlotArea.InsideTop, _
        mChart.PlotArea.InsideWidth, mChart.PlotArea.InsideHeight, data.LengthUnit

    Dim i As Long
    If rebarLabelsEnabled Then
        For i = 1 To data.AnnotationCount
            If StrComp(data.AnnotationType(i), "REBAR_ANNOTATION", vbTextCompare) = 0 Then
                layout.AddRebarLabel data.AnnotationStartX(i), data.AnnotationStartY(i), _
                    data.AnnotationEndX(i), data.AnnotationEndY(i), data.AnnotationNormalX(i), _
                    data.AnnotationNormalY(i), data.AnnotationOffset(i), data.AnnotationText(i)
            End If
        Next i
    End If

    If dimensionsEnabled Then
        For i = 1 To data.AnnotationCount
            If StrComp(data.AnnotationType(i), "DIMENSION", vbTextCompare) = 0 Then
                layout.AddDimension data.AnnotationStartX(i), data.AnnotationStartY(i), _
                    data.AnnotationEndX(i), data.AnnotationEndY(i), data.AnnotationNormalX(i), _
                    data.AnnotationNormalY(i), data.AnnotationOffset(i), data.AnnotationText(i)
            End If
        Next i
    End If

    RenderAnnotationLayout layout
End Sub

' Создает Chart.Shapes по готовому layout.
' Dimension lines получают стрелки, тексты и выносные линии уже с рассчитанными
' координатами внутри Chart.
Private Sub RenderAnnotationLayout(ByVal layout As CPlotAnnotationLayout)
    Dim i As Long
    For i = 1 To layout.Count
        If StrComp(layout.ItemKind(i), "LINE", vbTextCompare) = 0 Then
            Dim lineShape As Object
            Set lineShape = DrawChartLine(layout.X1(i), layout.Y1(i), layout.X2(i), layout.Y2(i), _
                layout.Color(i), layout.Weight(i), "AnnotationLine")
            If layout.Arrowheads(i) Then
                With lineShape.Line
                    .BeginArrowheadStyle = DimensionArrowStyle()
                    .EndArrowheadStyle = DimensionArrowStyle()
                    .BeginArrowheadLength = DimensionArrowLength()
                    .EndArrowheadLength = DimensionArrowLength()
                    .BeginArrowheadWidth = DimensionArrowWidth()
                    .EndArrowheadWidth = DimensionArrowWidth()
                End With
            End If
        ElseIf StrComp(layout.ItemKind(i), "TEXT", vbTextCompare) = 0 Then
            AddChartText layout.X1(i), layout.Y1(i), layout.X2(i), layout.Y2(i), _
                layout.Text(i), layout.Rotation(i), layout.Color(i), layout.Bold(i), True, layout.FontSize(i)
        End If
    Next i
End Sub

' Нормализует пользовательский тип стрелки размерной линии.
Private Function NormalizeArrowType(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "stealth": NormalizeArrowType = "Stealth"
        Case "diamond": NormalizeArrowType = "Diamond"
        Case "oval": NormalizeArrowType = "Oval"
        Case "open": NormalizeArrowType = "Open"
        Case Else: NormalizeArrowType = "Triangle"
    End Select
End Function

' Нормализует пользовательский размер стрелки размерной линии.
Private Function NormalizeArrowSize(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "small": NormalizeArrowSize = "Small"
        Case "medium": NormalizeArrowSize = "Medium"
        Case Else: NormalizeArrowSize = "Wide"
    End Select
End Function

' Преобразует текстовую настройку стрелки в константу Office.
Private Function DimensionArrowStyle() As Long
    Select Case mDimensionArrowType
        Case "Stealth": DimensionArrowStyle = MSO_ARROWHEAD_STEALTH
        Case "Diamond": DimensionArrowStyle = MSO_ARROWHEAD_DIAMOND
        Case "Oval": DimensionArrowStyle = MSO_ARROWHEAD_OVAL
        Case "Open": DimensionArrowStyle = MSO_ARROWHEAD_OPEN
        Case Else: DimensionArrowStyle = MSO_ARROWHEAD_TRIANGLE
    End Select
End Function

' Преобразует размер стрелки в длину наконечника Office.
Private Function DimensionArrowLength() As Long
    Select Case mDimensionArrowSize
        Case "Small": DimensionArrowLength = MSO_ARROWHEAD_LENGTH_SHORT
        Case "Medium": DimensionArrowLength = MSO_ARROWHEAD_LENGTH_MEDIUM
        Case Else: DimensionArrowLength = MSO_ARROWHEAD_LENGTH_LONG
    End Select
End Function

' Преобразует размер стрелки в ширину наконечника Office.
Private Function DimensionArrowWidth() As Long
    Select Case mDimensionArrowSize
        Case "Small": DimensionArrowWidth = MSO_ARROWHEAD_WIDTH_NARROW
        Case "Medium": DimensionArrowWidth = MSO_ARROWHEAD_WIDTH_MEDIUM
        Case Else: DimensionArrowWidth = MSO_ARROWHEAD_WIDTH_WIDE
    End Select
End Function

' Расставляет пространственные численные подписи Stress/Strain по бетонным элементам.
' Подписи не ставятся на неработающий растянутый бетон и выбираются по сетке
' spacing, а не по порядковому ElementID.
Private Sub DrawResultLabels(ByVal data As CSectionPlotDataReader, ByVal spacing As Double)
    If spacing <= 0# Then spacing = MaxDouble(mMaxX - mMinX, mMaxY - mMinY) / 12#
    Dim bestIndexByCell As Object
    Set bestIndexByCell = CreateObject("Scripting.Dictionary")
    Dim bestDistanceByCell As Object
    Set bestDistanceByCell = CreateObject("Scripting.Dictionary")

    ' Пространственные подписи выбираются не по ElementID, а по расчетной сетке ячеек.
    ' Для каждой ячейки оставляем ближайшее бетонное волокно к ее центру, поэтому подписи
    ' распределяются по площади сечения равномернее и не зависят от порядка записи элементов.
    Dim i As Long
    For i = 1 To data.Count
        If StrComp(data.MaterialType(i), "Concrete", vbTextCompare) = 0 And _
                StrComp(data.PhysicalState(i), "InactiveTensionConcrete", vbTextCompare) <> 0 Then
            Dim cellX As Long
            Dim cellY As Long
            cellX = CLng(Int((data.X(i) - mMinX) / spacing))
            cellY = CLng(Int((data.Y(i) - mMinY) / spacing))

            Dim key As String
            key = CStr(cellX) & ":" & CStr(cellY)

            Dim targetX As Double
            Dim targetY As Double
            targetX = mMinX + (CDbl(cellX) + 0.5) * spacing
            targetY = mMinY + (CDbl(cellY) + 0.5) * spacing

            Dim distance2 As Double
            distance2 = (data.X(i) - targetX) * (data.X(i) - targetX) + _
                (data.Y(i) - targetY) * (data.Y(i) - targetY)

            If Not bestIndexByCell.Exists(key) Then
                bestIndexByCell.Add key, i
                bestDistanceByCell.Add key, distance2
            ElseIf distance2 < CDbl(bestDistanceByCell(key)) Then
                bestIndexByCell(key) = i
                bestDistanceByCell(key) = distance2
            End If
        End If
    Next i

    Dim cellKey As Variant
    For Each cellKey In bestIndexByCell.Keys
        i = CLng(bestIndexByCell(cellKey))
        DrawTextAt data.X(i), data.Y(i), FormatPlotValue(data.ResultValue(i)), 1#, 0#, RGB(60, 60, 60), False
    Next cellKey
End Sub

' Рисует одну общую или две раздельные легенды результата.
' Раздельный режим считает диапазоны бетона и арматуры независимо, поэтому
' цвета материалов не “слипаются” из-за разных уровней напряжений.
Private Sub DrawLegend(ByVal data As CSectionPlotDataReader)
    Dim leftPos As Double, topPos As Double, barW As Double, barH As Double
    leftPos = mChartObject.Width - 142
    topPos = 88
    barW = 18
    barH = 96

    ' Легенда использует те же диапазоны, что и цвет элементов. В режиме Separate бетон и
    ' арматура получают независимые шкалы, поэтому их экстремумы не смешиваются визуально.
    If IsCommonLegendMode() Then
        DrawChartLegendScale data, leftPos, topPos, barW, barH, "Rebar", _
            mHasCompression, mCompressionValueAtMaxAbs, mHasTension, mTensionValueAtMaxAbs
    Else
        DrawChartLegendScale data, leftPos, topPos, barW, barH, "Rebar", _
            mRebarHasCompression, mRebarCompressionValueAtMaxAbs, mRebarHasTension, mRebarTensionValueAtMaxAbs
        DrawChartLegendScale data, leftPos, topPos + barH + 62, barW, barH, "Concrete", _
            mConcreteHasCompression, mConcreteCompressionValueAtMaxAbs, mConcreteHasTension, mConcreteTensionValueAtMaxAbs
    End If
End Sub

' Рисует вертикальную цветовую шкалу для одного материала.
' Верх шкалы соответствует сжатию, середина - нулю/серому цвету, низ - растяжению.
Private Sub DrawChartLegendScale(ByVal data As CSectionPlotDataReader, ByVal leftPos As Double, _
        ByVal topPos As Double, ByVal barW As Double, ByVal barH As Double, _
        ByVal materialType As String, ByVal hasCompression As Boolean, ByVal compressionValue As Double, _
        ByVal hasTension As Boolean, ByVal tensionValue As Double)
    AddChartText leftPos - 6, topPos - 28, 132, 22, data.ResultCaption & ", " & data.ResultUnit, 0#, RGB(30, 30, 30), True

    Dim i As Long
    For i = 0 To 11
        Dim t As Double
        t = CDbl(i) / 11#
        AddLegendRect leftPos, topPos + i * barH / 12#, barW, barH / 12# + 1, MaterialLegendColor(materialType, t)
    Next i
    AddChartText leftPos + 24, topPos - 12, 88, 14, LegendCompressionText(), 0#, RGB(40, 40, 40), True
    AddChartText leftPos + 24, topPos + 4, 84, 16, LegendValueText(hasCompression, compressionValue), 0#, RGB(40, 40, 40), False
    AddChartText leftPos + 24, topPos + barH / 2# - 8, 84, 16, "0", 0#, RGB(40, 40, 40), False
    AddChartText leftPos + 24, topPos + barH - 10, 84, 16, LegendValueText(hasTension, tensionValue), 0#, RGB(40, 40, 40), False
    AddChartText leftPos + 24, topPos + barH + 4, 88, 14, LegendTensionText(), 0#, RGB(40, 40, 40), True
End Sub

' Возвращает русскую подпись "Сжатие" без зависимости от кодировки VBE.
Private Function LegendCompressionText() As String
    LegendCompressionText = ChrW$(&H421) & ChrW$(&H436) & ChrW$(&H430) & ChrW$(&H442) & ChrW$(&H438) & ChrW$(&H435)
End Function

' Возвращает русскую подпись "Растяжение" без зависимости от кодировки VBE.
Private Function LegendTensionText() As String
    LegendTensionText = ChrW$(&H420) & ChrW$(&H430) & ChrW$(&H441) & ChrW$(&H442) & ChrW$(&H44F) & ChrW$(&H436) & ChrW$(&H435) & ChrW$(&H43D) & ChrW$(&H438) & ChrW$(&H435)
End Function

' Возвращает цвет легенды в относительной точке шкалы 0..1.
Private Function MaterialLegendColor(ByVal materialType As String, ByVal position As Double) As Long
    Dim t As Double
    t = ClampDouble(position, 0#, 1#)
    If t < 0.5 Then
        MaterialLegendColor = BlendColor(LegendCompressionColor(materialType), RGB(218, 218, 218), t / 0.5)
    Else
        MaterialLegendColor = BlendColor(RGB(218, 218, 218), LegendTensionColor(materialType), (t - 0.5) / 0.5)
    End If
End Function

' Выбирает цвет максимального сжатия для материала или общей легенды.
Private Function LegendCompressionColor(ByVal materialType As String) As Long
    If IsCommonLegendMode() Or StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        LegendCompressionColor = mRebarCompressionColor
    Else
        LegendCompressionColor = mConcreteCompressionColor
    End If
End Function

' Выбирает цвет максимального растяжения для материала или общей легенды.
Private Function LegendTensionColor(ByVal materialType As String) As Long
    If IsCommonLegendMode() Or StrComp(materialType, "Rebar", vbTextCompare) = 0 Then
        LegendTensionColor = mRebarTensionColor
    Else
        LegendTensionColor = mConcreteTensionColor
    End If
End Function

' True, если пользователь выбрал одну общую легенду вместо отдельных шкал.
Private Function IsCommonLegendMode() As Boolean
    IsCommonLegendMode = (StrComp(mLegendMode, "Common", vbTextCompare) = 0)
End Function

' Старый цветовой переход общей физической шкалы.
' Оставлен как совместимая утилита, но основная легенда сейчас строится
' через MaterialLegendColor.
Private Function PhysicalLegendColor(ByVal position As Double) As Long
    Dim t As Double
    t = ClampDouble(position, 0#, 1#)
    If t < 0.5 Then
        PhysicalLegendColor = BlendColor(RGB(210, 30, 20), RGB(255, 238, 120), t / 0.5)
    Else
        PhysicalLegendColor = BlendColor(RGB(210, 245, 250), RGB(30, 80, 220), (t - 0.5) / 0.5)
    End If
End Function

' Формирует подпись min/max легенды; если состояния нет, показывает "нет".
Private Function LegendValueText(ByVal hasValue As Boolean, ByVal value As Double) As String
    If hasValue Then
        LegendValueText = FormatPlotValue(value)
    Else
        LegendValueText = "нет"
    End If
End Function

' Добавляет один цветной прямоугольник вертикальной легенды.
Private Sub AddLegendRect(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal colorValue As Long)
    Dim marker As Object
    Set marker = mChart.Shapes.AddShape(MSO_SHAPE_RECTANGLE, leftPos, topPos, widthValue, heightValue)
    marker.Name = SHAPE_PREFIX & "LegendGradient" & CStr(mChart.Shapes.Count)
    marker.Fill.ForeColor.RGB = colorValue
    marker.Line.ForeColor.RGB = colorValue
End Sub

' Добавляет строку легенды с цветным маркером и текстом.
' Сейчас почти вся легенда строится градиентными прямоугольниками, но строка
' оставлена для компактных пояснений.
Private Sub AddLegendRow(ByVal leftPos As Double, ByVal topPos As Double, ByVal colorValue As Long, ByVal textValue As String)
    Dim marker As Object
    Set marker = mChart.Shapes.AddShape(MSO_SHAPE_RECTANGLE, leftPos, topPos + 3, 10, 10)
    marker.Name = SHAPE_PREFIX & "LegendMarker" & CStr(mChart.Shapes.Count)
    marker.Fill.ForeColor.RGB = colorValue
    marker.Line.ForeColor.RGB = colorValue
    AddChartText leftPos + 14, topPos, 118, 28, textValue, 0#, RGB(50, 50, 50), False
End Sub

' Рисует Shape-линию по координатам модели.
' Используется для контуров и аннотационных линий, когда нужен именно Chart.Shape.
Private Function DrawModelLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String) As Object
    Set DrawModelLine = DrawChartLine(ModelToChartX(x1), ModelToChartY(y1), ModelToChartX(x2), ModelToChartY(y2), _
        colorValue, weight, suffix)
End Function

' Рисует линию в координатах Chart/PlotArea.
' Все model-to-chart преобразования должны быть сделаны до вызова этого метода.
Private Function DrawChartLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String) As Object
    Dim lineShape As Object
    Set lineShape = mChart.Shapes.AddLine(x1, y1, x2, y2)
    lineShape.Name = SHAPE_PREFIX & suffix & CStr(mChart.Shapes.Count)
    lineShape.Line.ForeColor.RGB = colorValue
    lineShape.Line.Weight = weight
    Set DrawChartLine = lineShape
End Function

' Рисует расчетную линию как XY-series.
' Такой способ используется для нейтральной линии и главных осей, чтобы они
' оставались поверх расчетных элементов и масштабировались вместе с осями.
Private Sub DrawModelLineSeries(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String, Optional ByVal dashStyle As Long = MSO_LINE_SOLID)
    Dim xValues(1 To 2) As Double
    Dim yValues(1 To 2) As Double
    xValues(1) = x1
    xValues(2) = x2
    yValues(1) = y1
    yValues(2) = y2

    Dim series As Object
    Set series = mChart.SeriesCollection.NewSeries
    series.Name = SHAPE_PREFIX & suffix & CStr(mChart.SeriesCollection.Count)
    series.ChartType = XL_XY_SCATTER_LINES_NO_MARKERS
    series.XValues = xValues
    series.Values = yValues
    series.MarkerStyle = XL_NONE
    series.Format.Line.Visible = True
    series.Format.Line.ForeColor.RGB = colorValue
    series.Format.Line.Weight = weight
    On Error Resume Next
    series.Border.LineStyle = dashStyle
    series.Format.Line.DashStyle = dashStyle
    series.Smooth = False
    On Error GoTo 0
End Sub

' Рисует овал/круг в координатах модели как Shape.
' Нужен для semantic-контура круга; сами арматурные стержни рисуются series.
Private Function DrawModelOval(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String, ByVal filled As Boolean) As Object
    Dim leftPos As Double, topPos As Double, widthValue As Double, heightValue As Double
    leftPos = ModelToChartX(MinDouble(x1, x2))
    topPos = ModelToChartY(MaxDouble(y1, y2))
    widthValue = Abs(ModelToChartX(x2) - ModelToChartX(x1))
    heightValue = Abs(ModelToChartY(y2) - ModelToChartY(y1))

    Dim ovalShape As Object
    Set ovalShape = mChart.Shapes.AddShape(MSO_SHAPE_OVAL, leftPos, topPos, widthValue, heightValue)
    ovalShape.Name = SHAPE_PREFIX & suffix & CStr(mChart.Shapes.Count)
    ovalShape.Line.ForeColor.RGB = colorValue
    ovalShape.Line.Weight = weight
    If filled Then
        ovalShape.Fill.ForeColor.RGB = colorValue
    Else
        ovalShape.Fill.Visible = False
    End If
    Set DrawModelOval = ovalShape
End Function

' Рисует одиночную Shape-точку.
' Сейчас расчетные точки предпочтительно рисуются series; метод оставлен для
' специальных декоративных объектов, если им нужен Shape-контракт.
Private Sub DrawPoint(ByVal xCoord As Double, ByVal yCoord As Double, ByVal sizeValue As Double, _
        ByVal colorValue As Long, ByVal labelText As String, ByVal shapeType As Long)
    Dim leftPos As Double, topPos As Double
    leftPos = ModelToChartX(xCoord) - sizeValue / 2#
    topPos = ModelToChartY(yCoord) - sizeValue / 2#
    Dim pointShape As Object
    Set pointShape = mChart.Shapes.AddShape(shapeType, leftPos, topPos, sizeValue, sizeValue)
    pointShape.Name = SHAPE_PREFIX & "Point" & labelText
    pointShape.Fill.ForeColor.RGB = colorValue
    pointShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    If Len(labelText) > 0 Then DrawTextAt xCoord, yCoord, labelText, 1#, 0#, RGB(30, 30, 30), True
End Sub

' Рисует одиночный маркер как XY-series.
' Так выводится точка приложения нагрузки: она остается в координатах модели,
' а не в абсолютных координатах листа.
Private Sub DrawPointMarker(ByVal xCoord As Double, ByVal yCoord As Double, ByVal sizeValue As Double, _
        ByVal colorValue As Long, ByVal markerStyle As Long, ByVal seriesName As String)
    Dim xValues(1 To 1) As Double
    Dim yValues(1 To 1) As Double
    xValues(1) = xCoord
    yValues(1) = yCoord

    Dim series As Object
    Set series = mChart.SeriesCollection.NewSeries
    series.Name = SHAPE_PREFIX & seriesName
    series.XValues = xValues
    series.Values = yValues
    series.MarkerStyle = markerStyle
    series.MarkerSize = sizeValue
    series.Format.Line.Visible = False
    series.MarkerForegroundColor = RGB(0, 0, 0)
    series.MarkerBackgroundColor = colorValue
End Sub

' Размещает короткую подпись около точки модели.
' Используется для численных подписей результатов внутри бетонных элементов.
Private Sub DrawTextAt(ByVal xCoord As Double, ByVal yCoord As Double, ByVal textValue As String, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal colorValue As Long, ByVal isBold As Boolean)
    Dim leftPos As Double, topPos As Double
    leftPos = ModelToChartX(xCoord) - 32
    topPos = ModelToChartY(yCoord) - 9
    Dim rotation As Double
    If Abs(dirY) > Abs(dirX) Then rotation = 90#
    AddChartText leftPos, topPos, 64, 18, textValue, rotation, colorValue, isBold, True
End Sub

' Создает текстовый Shape внутри Chart.
' Важно: это не Worksheet.Shape, поэтому при перемещении Chart все подписи
' остаются внутри единого графического объекта.
Private Function AddChartText(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal textValue As String, ByVal rotation As Double, _
        ByVal colorValue As Long, ByVal isBold As Boolean, Optional ByVal centered As Boolean = False, _
        Optional ByVal fontSize As Double = 8#) As Object
    Dim textShape As Object
    Set textShape = mChart.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, leftPos, topPos, widthValue, heightValue)
    textShape.Name = SHAPE_PREFIX & "Text" & CStr(mChart.Shapes.Count)
    textShape.TextFrame.Characters().Text = textValue
    If fontSize <= 0# Then fontSize = 8#
    textShape.TextFrame.Characters().Font.Size = fontSize
    textShape.TextFrame.Characters().Font.Color = colorValue
    textShape.TextFrame.Characters().Font.Bold = isBold
    If centered Then
        textShape.TextFrame.HorizontalAlignment = XL_CENTER
        textShape.TextFrame.VerticalAlignment = XL_VCENTER
        textShape.TextFrame.MarginLeft = 0
        textShape.TextFrame.MarginRight = 0
        textShape.TextFrame.MarginTop = 0
        textShape.TextFrame.MarginBottom = 0
    End If
    textShape.Line.Visible = False
    textShape.Fill.Visible = False
    If rotation <> 0# Then textShape.Rotation = rotation
    Set AddChartText = textShape
End Function

Private Function ModelToChartX(ByVal xCoord As Double) As Double
    ModelToChartX = mChart.PlotArea.InsideLeft + (xCoord - mMinX) / (mMaxX - mMinX) * mChart.PlotArea.InsideWidth
End Function

Private Function ModelToChartY(ByVal yCoord As Double) As Double
    ModelToChartY = mChart.PlotArea.InsideTop + mChart.PlotArea.InsideHeight - _
        (yCoord - mMinY) / (mMaxY - mMinY) * mChart.PlotArea.InsideHeight
End Function

Private Function ChartScale() As Double
    Dim sx As Double, sy As Double
    sx = mChart.PlotArea.InsideWidth / (mMaxX - mMinX)
    sy = mChart.PlotArea.InsideHeight / (mMaxY - mMinY)
    ChartScale = MinDouble(sx, sy)
End Function

' Очищает накопленное состояние перед новым расчетом или повторным формированием вывода.
Private Sub ClearGeneratedShapes()
    Dim i As Long
    For i = mChart.Shapes.Count To 1 Step -1
        If Left$(mChart.Shapes.Item(i).Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            mChart.Shapes.Item(i).Delete
        End If
    Next i
End Sub

Private Function ConcreteElementWidth(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Double
    ConcreteElementWidth = data.Width(index)
    If ConcreteElementWidth <= 0# Then ConcreteElementWidth = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
End Function

Private Function ConcreteElementHeight(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Double
    ConcreteElementHeight = data.Height(index)
    If ConcreteElementHeight <= 0# Then ConcreteElementHeight = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
End Function

Private Function ConcreteMarkerLength(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Double
    If data.Width(index) > 0# And data.Height(index) > 0# Then
        ConcreteMarkerLength = Sqr(data.Width(index) * data.Height(index))
    ElseIf data.Width(index) > 0# Then
        ConcreteMarkerLength = data.Width(index)
    ElseIf data.Height(index) > 0# Then
        ConcreteMarkerLength = data.Height(index)
    Else
        ConcreteMarkerLength = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
    End If
End Function

Private Function EquivalentLengthFromArea(ByVal areaValue As Double, ByVal areaUnit As String, ByVal lengthUnit As String) As Double
    If areaValue <= 0# Then
        EquivalentLengthFromArea = 0#
    Else
        EquivalentLengthFromArea = Sqr(areaValue * AreaFactorToMm2(areaUnit)) / LengthFactorToMm(lengthUnit)
    End If
End Function

Private Function LengthFactorToMm(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "mm": LengthFactorToMm = 1#
        Case "cm": LengthFactorToMm = 10#
        Case "m": LengthFactorToMm = 1000#
        Case Else: LengthFactorToMm = 1#
    End Select
End Function

Private Function AreaFactorToMm2(ByVal unitName As String) As Double
    Select Case LCase$(Trim$(unitName))
        Case "mm2": AreaFactorToMm2 = 1#
        Case "cm2": AreaFactorToMm2 = 100#
        Case "m2": AreaFactorToMm2 = 1000000#
        Case Else: AreaFactorToMm2 = LengthFactorToMm(Replace$(LCase$(Trim$(unitName)), "2", "")) ^ 2
    End Select
End Function

Private Function PlotTitlePrefix() As String
    PlotTitlePrefix = ChrW$(&H421) & ChrW$(&H445) & ChrW$(&H435) & ChrW$(&H43C) & ChrW$(&H430) & " " & _
        ChrW$(&H441) & ChrW$(&H435) & ChrW$(&H447) & ChrW$(&H435) & ChrW$(&H43D) & ChrW$(&H438) & ChrW$(&H44F) & " " & _
        ChrW$(&H43F) & ChrW$(&H440) & ChrW$(&H438) & " " & ChrW$(&H437) & ChrW$(&H430) & ChrW$(&H433) & _
        ChrW$(&H440) & ChrW$(&H443) & ChrW$(&H436) & ChrW$(&H435) & ChrW$(&H43D) & ChrW$(&H438) & ChrW$(&H438) & " "
End Function

Private Function PlotTitleText(ByVal data As CSectionPlotDataReader) As String
    PlotTitleText = PlotTitlePrefix() & data.LoadCase
    If Len(Trim$(data.LoadCaseComment)) > 0 Then PlotTitleText = PlotTitleText & " (" & data.LoadCaseComment & ")"
End Function

Private Function NormalizeLegendMode(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "common": NormalizeLegendMode = "Common"
        Case "separate": NormalizeLegendMode = "Separate"
        Case Else: NormalizeLegendMode = "Separate"
    End Select
End Function

Private Function ParseRgb(ByVal value As String, ByVal defaultColor As Long) As Long
    Dim parts As Variant
    parts = Split(Replace$(Trim$(value), ";", ","), ",")
    If UBound(parts) <> 2 Then
        ParseRgb = defaultColor
        Exit Function
    End If

    On Error GoTo InvalidColor
    Dim r As Long, g As Long, b As Long
    r = CLng(Trim$(CStr(parts(0))))
    g = CLng(Trim$(CStr(parts(1))))
    b = CLng(Trim$(CStr(parts(2))))
    If r < 0 Or r > 255 Or g < 0 Or g > 255 Or b < 0 Or b > 255 Then GoTo InvalidColor
    ParseRgb = RGB(r, g, b)
    Exit Function

InvalidColor:
    ParseRgb = defaultColor
End Function

Private Function FormatPlotValue(ByVal value As Double) As String
    Dim pattern As String
    If mResultPrecision <= 0 Then
        pattern = "0"
    Else
        pattern = "0." & String$(mResultPrecision, "0")
    End If
    FormatPlotValue = Replace$(Format$(value, pattern), ",", ".")
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.######"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatNumberInvariant = textValue
End Function

Private Function RoundTo(ByVal value As Double, ByVal digits As Long) As Double
    RoundTo = Round(value, digits)
End Function

Private Function RoundToLong(ByVal value As Double) As Long
    RoundToLong = CLng(value + 0.5)
End Function

Private Function ClampDouble(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    ClampDouble = value
    If ClampDouble < minValue Then ClampDouble = minValue
    If ClampDouble > maxValue Then ClampDouble = maxValue
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
