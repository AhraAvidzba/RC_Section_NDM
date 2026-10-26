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
' Отрисовка расчётного сечения на диаграмме Excel
' ==========================================================================
' Рисует схему по сохранённым данным Results, прочитанным CSectionPlotDataReader.
' Из Config получает оформление; геометрию и нагрузки берёт из сохранённого снимка Results.
' На обычных и средних сетках выводит элементы через Chart.Shapes с сохранением
' их размеров. Для очень больших сеток применяет ряды диаграммы с группировкой
' по цвету, чтобы ограничить объём отрисовки.

Private Const CHART_NAME As String = "chtNDMSectionPlot" ' Имя диаграммы схемы на листе Excel.
Private Const SHAPE_PREFIX As String = "NDMPlot_" ' Начало имён фигур, созданных при отрисовке схемы.
Private Const XL_XY_SCATTER As Long = -4169 ' Код Excel: точечная диаграмма.
Private Const XL_XY_SCATTER_SMOOTH As Long = 72 ' Код Excel: сглаженная точечная диаграмма с маркерами.
Private Const XL_XY_SCATTER_SMOOTH_NO_MARKERS As Long = 73 ' Код Excel: сглаженная точечная диаграмма без маркеров.
Private Const XL_XY_SCATTER_LINES As Long = 74 ' Код Excel: точечная диаграмма с прямыми и маркерами.
Private Const XL_XY_SCATTER_LINES_NO_MARKERS As Long = 75 ' Код Excel: точечная диаграмма с прямыми без маркеров.
Private Const XL_CATEGORY As Long = 1 ' Код горизонтальной оси Excel.
Private Const XL_VALUE As Long = 2 ' Код вертикальной оси Excel.
Private Const XL_PRIMARY As Long = 1 ' Код основной группы осей Excel.
Private Const XL_MARKER_SQUARE As Long = 1 ' Код квадратного маркера.
Private Const XL_MARKER_CIRCLE As Long = 8 ' Код круглого маркера.
Private Const XL_NONE As Long = -4142 ' Код отсутствующего элемента оформления.
Private Const XL_CENTER As Long = -4108 ' Код горизонтального центрирования.
Private Const XL_VCENTER As Long = -4108 ' Код вертикального центрирования.
Private Const MSO_TEXT_ORIENTATION_HORIZONTAL As Long = 1 ' Код горизонтального текста.
Private Const MSO_TEXT_ORIENTATION_UPWARD As Long = 2 ' Код текста с чтением снизу вверх.
Private Const MSO_TEXT_ORIENTATION_DOWNWARD As Long = 3 ' Код текста с чтением сверху вниз.
Private Const MSO_SEGMENT_LINE As Long = 0 ' Код прямолинейного участка фигуры.
Private Const MSO_SEGMENT_CURVE As Long = 1 ' Код кубической кривой Безье в свободной фигуре Excel.
Private Const MSO_EDITING_CORNER As Long = 1 ' Контрольные точки дуги задаются явно, без автоматического сглаживания соседних ребер.
Private Const MSO_EDITING_AUTO As Long = 0 ' Код автоматического соединения узлов фигуры.
Private Const MSO_ARROWHEAD_TRIANGLE As Long = 2 ' Код треугольной стрелки.
Private Const MSO_ARROWHEAD_STEALTH As Long = 3 ' Код стрелки с выемкой.
Private Const MSO_ARROWHEAD_DIAMOND As Long = 4 ' Код ромбовидной стрелки.
Private Const MSO_ARROWHEAD_OVAL As Long = 5 ' Код овальной стрелки.
Private Const MSO_ARROWHEAD_OPEN As Long = 6 ' Код открытой стрелки.
Private Const MSO_ARROWHEAD_LENGTH_SHORT As Long = 1 ' Код короткой стрелки.
Private Const MSO_ARROWHEAD_LENGTH_MEDIUM As Long = 2 ' Код средней длины стрелки.
Private Const MSO_ARROWHEAD_LENGTH_LONG As Long = 3 ' Код длинной стрелки.
Private Const MSO_ARROWHEAD_WIDTH_NARROW As Long = 1 ' Код узкой стрелки.
Private Const MSO_ARROWHEAD_WIDTH_MEDIUM As Long = 2 ' Код средней ширины стрелки.
Private Const MSO_ARROWHEAD_WIDTH_WIDE As Long = 3 ' Код широкой стрелки.
Private Const MSO_BRING_TO_FRONT As Long = 0 ' Команда перемещения фигуры на передний план.
Private Const MSO_SHAPE_RECTANGLE As Long = 1 ' Код прямоугольной фигуры.
Private Const MSO_SHAPE_DIAMOND As Long = 4 ' Код ромбовидной фигуры.
Private Const MSO_SHAPE_OVAL As Long = 9 ' Код овальной фигуры.
Private Const EXTENSION_WARNING_TEXT As String = "ВНЕ ФИЗИЧЕСКОЙ ДИАГРАММЫ МАТЕРИАЛА" ' Текст предупреждения об использовании продолжения диаграммы.
Private Const NUMERICAL_STATE_WARNING_TEXT As String = "ПРЯМОЕ НДС НЕ СОШЛОСЬ" ' Текст предупреждения о недостигнутом равновесии.
Private Const RESULT_COLOR_BUCKET_COUNT As Long = 16 ' Число цветовых ступеней для расчетных маркеров Excel Chart.
Private Const CONCRETE_MARKER_SIZE_BUCKET As Long = 3 ' Шаг группировки бетонных маркеров, чтобы Chart не превысил лимит series.
Private Const ELEMENT_SHAPE_RENDER_MAX_COUNT As Long = 5000 ' До этого числа элементов геометрию рисуем поэлементно Shape-ами.
Private Const ELEMENT_SERIES_MAX_COUNT As Long = 255 ' Безопасный предел рядов XY-диаграммы; при большем числе оформлений используем отдельные фигуры.
Private Const CONCRETE_SHAPE_FILL_RATIO As Double = 0.7 ' Визуальный размер бетонной ячейки: оставляет зазор между соседними элементами.
Private Const REFERENCE_AXIS_WEIGHT As Double = 0.56 ' Толщина сплошных главных осей; немного тоньше нейтральной линии.
Private Const REFERENCE_BAND_MIN_THICKNESS As Double = 0.4 ' Минимальная толщина Freeform-полосы, чтобы тонкие оси оставались видимыми.
Private Const PLOT_TITLE_TOP As Double = 6# ' Верхний отступ собственного Shape-заголовка схемы внутри Chart.
Private Const PLOT_TITLE_HEIGHT As Double = 22# ' Высота собственного Shape-заголовка схемы.
Private Const PLOT_FRAME_LEFT As Double = 28# ' Левый край рабочей области схемы внутри Chart.
Private Const PLOT_FRAME_TOP As Double = 36# ' Верхний край рабочей области схемы внутри Chart.
Private Const PLOT_FRAME_WIDTH_MARGIN As Double = 180# ' Запас справа для легенд и подписей.
Private Const PLOT_FRAME_HEIGHT_MARGIN As Double = 72# ' Запас снизу для предупреждений и подписи режима.
Private Const AXIS_LABEL_MARGIN As Double = 6# ' Отступ подписей осей от внутренней рамки построения.
Private Const AXIS_LABEL_X_WIDTH As Double = 48# ' Ширина горизонтальной подписи "Ось X".
Private Const AXIS_LABEL_Y_HEIGHT As Double = 48# ' Высота вертикальной подписи "Ось Y".
Private Const AXIS_LABEL_TEXT_HEIGHT As Double = 14# ' Толщина текстового блока подписи оси.

Private mChart As Object ' Excel Chart, внутри которого рисуются series и аннотационные Shapes.
Private mChartObject As Object ' Контейнер ChartObject на листе Расчет; пользователь может двигать/масштабировать его.
Private mMinX As Double ' Минимум оси X после добавления полей для аннотаций.
Private mMaxX As Double ' Максимум оси X после добавления полей для аннотаций.
Private mMinY As Double ' Минимум оси Y после добавления полей для аннотаций.
Private mMaxY As Double ' Максимум оси Y после добавления полей для аннотаций.
Private mPlotLeft As Double ' Зафиксированная левая граница PlotArea для всех Shape-объектов текущей схемы.
Private mPlotTop As Double ' Зафиксированная верхняя граница PlotArea для всех Shape-объектов текущей схемы.
Private mPlotWidth As Double ' Зафиксированная ширина PlotArea; не перечитывается во время отрисовки.
Private mPlotHeight As Double ' Зафиксированная высота PlotArea; не перечитывается во время отрисовки.
Private mPlotMetricsCaptured As Boolean ' True после снимка PlotArea для текущего прохода отрисовки.
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
Private mDimensionArrowType As String ' Тип стрелок размерных линий из Config.
Private mDimensionArrowSize As String ' Размер стрелок размерных линий из Config.
Private mLegendMode As String ' Separate или Common для легенд бетона/арматуры.
Private mRebarCompressionColor As Long ' Цвет максимального сжатия арматуры.
Private mRebarTensionColor As Long ' Цвет максимального растяжения арматуры.
Private mConcreteCompressionColor As Long ' Цвет максимального сжатия бетона.
Private mConcreteTensionColor As Long ' Цвет максимального растяжения бетона.
Private mElementShapeLayerEnabled As Boolean ' True - расчетные элементы выводятся Shape-ами без укрупнения.
Private mElementGroupingEnabled As Boolean ' True - цвета большой сетки разбиты на ступени; сохраняется и при резервной отрисовке фигурами.
Private mAxisLabelsEnabled As Boolean ' Нужно ли выводить маленькие подписи глобальных осей X/Y.
Private mAxisLabelFontSize As Double ' Размер шрифта подписей глобальных осей X/Y в пунктах Excel.
Private mContourEnabled As Boolean ' Контур нужен как расчетной, так и geometry-only схеме.
Private mContourLineWeight As Double ' Общая толщина наружных границ и отверстий, в пунктах Excel.
Private mNeutralLineEnabled As Boolean ' Нейтральная линия доступна только для сохраненного НДС.
Private mLoadPointEnabled As Boolean ' Точка нагрузки для активного reference-слоя.
Private mResultLabelsEnabled As Boolean ' Подписи готовых Stress/Strain, без повторного solve.
Private mLegendEnabled As Boolean ' Цветовая легенда расчетного состояния.
Private mResultLabelSpacing As Double ' Расстояние между ячейками выбора численных подписей в координатах сечения, мм.
Private mPrincipalAxesMode As String ' Проверенный режим reference-осей текущего Draw.
Private mAnnotationLayout As CPlotAnnotationLayout ' Проверенные настройки и размещение аннотаций одного прохода Draw.
Private mAnnotationOffsetScale As Double ' Один коэффициент для обоих Offset: max(1; наибольший габарит / 1000 мм).
Private mUnits As CUnitSystem ' Единицы переданной модели схемы, без повторного применения пользовательских знаков.

' Создаёт общий перевод единиц для графических длин и площадей.
' Физические значения уже прочитаны из Results; отдельная таблица
' коэффициентов в построителе схемы не нужна.
Private Sub Class_Initialize()
    Set mUnits = New CUnitSystem
End Sub

' Перестраивает схему по данным сохранённого сечения.
' Метод управляет порядком операций: подготовка chart, равный масштаб X/Y,
' расчетные элементы, Shape-аннотации и верхний служебный Shape-оверлей.
Public Sub Draw(ByVal workbook As Object, ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4800, "CSectionPlotter", "Книга Excel не передана."
    If data Is Nothing Then Err.Raise vbObjectError + 4801, "CSectionPlotter", "Данные схемы не переданы."
    If settings Is Nothing Then Err.Raise vbObjectError + 4802, "CSectionPlotter", "Настройки Config не переданы."

    Set mAnnotationLayout = New CPlotAnnotationLayout ' Проверяем активное оформление до очистки существующей схемы.
    mAnnotationLayout.ConfigureFromSettings settings
    If mAnnotationLayout.DimensionsEnabled Then
        mDimensionArrowType = settings.GetRequiredChoice("Plot.Dimensions.ArrowType", Array("Triangle", "Stealth", "Diamond", "Oval", "Open"))
        mDimensionArrowSize = settings.GetRequiredChoice("Plot.Dimensions.ArrowSize", Array("Small", "Medium", "Wide"))
    End If
    mResultGradientEnabled = settings.GetRequiredBoolean("Plot.ResultGradient")
    mLegendMode = settings.GetRequiredChoice("Plot.LegendMode", Array("Separate", "Common"))
    mRebarCompressionColor = settings.GetRequiredRgbColor("Plot.Color.RebarCompression")
    mRebarTensionColor = settings.GetRequiredRgbColor("Plot.Color.RebarTension")
    If mLegendMode = "Separate" Then
        mConcreteCompressionColor = settings.GetRequiredRgbColor("Plot.Color.ConcreteCompression")
        mConcreteTensionColor = settings.GetRequiredRgbColor("Plot.Color.ConcreteTension")
    End If
    mAxisLabelsEnabled = settings.GetRequiredBoolean("Plot.AxisLabelsEnabled")
    If mAxisLabelsEnabled Then
        mAxisLabelFontSize = settings.GetRequiredDouble("Plot.AxisLabelsFontSize")
        If mAxisLabelFontSize <= 0# Then Err.Raise vbObjectError + 4812, "CSectionPlotter", _
            settings.InputErrorMessage("Plot.AxisLabelsFontSize", "Высота подписей осей должна быть больше нуля.", "Введите положительное число в пунктах Excel.")
    End If
    mContourEnabled = settings.GetRequiredBoolean("Plot.ContourEnabled")
    If mContourEnabled Then
        mContourLineWeight = settings.GetRequiredDouble("Plot.ContourLineWeight")
        If mContourLineWeight <= 0# Then Err.Raise vbObjectError + 4812, "CSectionPlotter", _
            settings.InputErrorMessage("Plot.ContourLineWeight", "Толщина линий контура должна быть больше нуля.", "Введите положительное число в пунктах Excel.")
    End If
    mNeutralLineEnabled = False: mResultLabelsEnabled = False: mLegendEnabled = False
    mLoadPointEnabled = False: mPrincipalAxesMode = "None"
    If Not data.GeometryOnly Then
        mNeutralLineEnabled = settings.GetRequiredBoolean("Plot.NeutralLineEnabled")
        mResultLabelsEnabled = settings.GetRequiredBoolean("Plot.ResultLabelsEnabled")
        mLegendEnabled = settings.GetRequiredBoolean("Plot.LegendEnabled")
        If mResultLabelsEnabled Then
            mResultLabelSpacing = settings.GetRequiredDouble("Plot.ResultLabelSpacing")
            If mResultLabelSpacing <= 0# Then Err.Raise vbObjectError + 4812, "CSectionPlotter", _
                settings.InputErrorMessage("Plot.ResultLabelSpacing", "Шаг подписей должен быть больше нуля.", "Введите положительное расстояние в мм.")
        End If
    End If
    If Not data.GeometryOnly Or data.GeometryReferenceItemsEnabled Then
        mPrincipalAxesMode = PlotPrincipalAxesMode(settings)
        mLoadPointEnabled = settings.GetRequiredBoolean("Plot.LoadApplicationPointEnabled")
    End If

    Set mChartObject = EnsureChart(workbook.Worksheets.Item("Расчет"))
    Set mChart = mChartObject.Chart

    mElementShapeLayerEnabled = ShouldDrawElementShapeLayer(data)
    mElementGroupingEnabled = Not mElementShapeLayerEnabled
    PrepareChart data
    ConfigureEqualScale data, settings
    ClearGeneratedShapes
    CalculateResultRange data
    AddElementSeries data, settings
    DrawShapeDecorations data
    If data.GeometryOnly Then
        DrawGeometryReferenceOverlay data
    Else
        DrawAnalysisOverlay data
    End If
    If mAxisLabelsEnabled Then DrawAxisNameLabels
    DrawPlotTitleOverlay PlotTitleText(data)
End Sub

' Очищает прежние элементы, подписи и заголовок, сохраняя саму диаграмму,
' её положение и размер. Устаревший расчёт не должен выглядеть текущим.
Public Sub ClearExisting(ByVal workbook As Object, ByVal titleText As String)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4800, "CSectionPlotter", "Книга Excel не передана."

    Set mChartObject = EnsureChart(workbook.Worksheets.Item("Расчет"))
    Set mChart = mChartObject.Chart


    mChart.HasLegend = False
    mChart.HasTitle = False

    Do While mChart.SeriesCollection.Count > 0
        mChart.SeriesCollection(1).Delete
    Loop
    ClearGeneratedShapes
    DrawPlotTitleOverlay titleText
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
    Select Case chartObject.Chart.ChartType ' Существующая XY-схема уже имеет две числовые оси. Повторное присвоение ChartType заполненной Shape-схеме сбивает подсветку ссылок Excel. Новый контейнер или вручную изменённый иной тип переводим в XY один раз.
        Case XL_XY_SCATTER, XL_XY_SCATTER_SMOOTH, XL_XY_SCATTER_SMOOTH_NO_MARKERS, _
                XL_XY_SCATTER_LINES, XL_XY_SCATTER_LINES_NO_MARKERS
        Case Else
            chartObject.Chart.ChartType = XL_XY_SCATTER
    End Select
    Set EnsureChart = chartObject
End Function

' Подготавливает Chart к новой отрисовке, сохраняя сам ChartObject.
' Старые series удаляются, но размеры и положение окна схемы остаются как у пользователя.
Private Sub PrepareChart(ByVal data As CSectionPlotDataReader)
    mPlotMetricsCaptured = False
    mPlotLeft = 0#
    mPlotTop = 0#
    mPlotWidth = 0#
    mPlotHeight = 0#

    mChart.HasLegend = False
    mChart.HasTitle = False

    Do While mChart.SeriesCollection.Count > 0
        mChart.SeriesCollection(1).Delete
    Loop

    On Error Resume Next
    mChart.ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    mChart.PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    mChart.PlotArea.Left = PLOT_FRAME_LEFT
    mChart.PlotArea.Top = PLOT_FRAME_TOP
    mChart.PlotArea.Width = StablePlotFrameWidth()
    mChart.PlotArea.Height = StablePlotFrameHeight()
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
    mAnnotationOffsetScale = MaxDouble(1#, MaxDouble(dx, dy) / 1000#)

    Dim annotationMargin As Double
    annotationMargin = AnnotationMarginModel(settings, data.LengthUnit, dx, dy)

    rawMinX = rawMinX - MaxDouble(dx * 0.16, annotationMargin)
    rawMaxX = rawMaxX + MaxDouble(dx * 0.24, annotationMargin)
    rawMinY = rawMinY - MaxDouble(dy * 0.18, annotationMargin)
    rawMaxY = rawMaxY + MaxDouble(dy * 0.16, annotationMargin)

    Dim plotW As Double, plotH As Double
    plotW = PlotCoordinateWidth()
    plotH = PlotCoordinateHeight()
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

' Оценивает запас координатной области под размерные линии и подписи арматуры.
' Значение считается в единицах модели, чтобы оси графика сразу оставляли
' место для внешних аннотаций.
Private Function AnnotationMarginModel(ByVal settings As CSystemSettingsReader, ByVal lengthUnit As String, _
        ByVal rangeX As Double, ByVal rangeY As Double) As Double
    If settings Is Nothing Then Exit Function

    Dim marginModel As Double
    If mAnnotationLayout.DimensionsEnabled Then
        marginModel = MaxDouble(marginModel, _
            LengthMmToModel(settings.GetDouble("Plot.Dimensions.Offset", 100#) * mAnnotationOffsetScale, lengthUnit) + _
            AnnotationTextMetricModel(settings.GetDouble("Plot.Dimensions.TextGap", 9#), _
                settings.GetString("Plot.Dimensions.TextUnits", "pt"), lengthUnit, rangeX, rangeY) + _
            AnnotationTextMetricModel(settings.GetDouble("Plot.Dimensions.TextHeight", 13#), _
                settings.GetString("Plot.Dimensions.TextUnits", "pt"), lengthUnit, rangeX, rangeY))
    End If
    If mAnnotationLayout.RebarLabelsEnabled Then
        marginModel = MaxDouble(marginModel, _
            LengthMmToModel(settings.GetDouble("Plot.RebarLabels.Offset", 30#) * mAnnotationOffsetScale, lengthUnit) + _
            AnnotationTextMetricModel(settings.GetDouble("Plot.RebarLabels.TextGap", 9#), _
                settings.GetString("Plot.RebarLabels.TextUnits", "pt"), lengthUnit, rangeX, rangeY) + _
            AnnotationTextMetricModel(settings.GetDouble("Plot.RebarLabels.TextHeight", 13#), _
                settings.GetString("Plot.RebarLabels.TextUnits", "pt"), lengthUnit, rangeX, rangeY))
    End If

    AnnotationMarginModel = marginModel
End Function

' Переводит заданный в мм отступ в единицы координат схемы.
' Отступ относится к геометрии сечения, а не к экранным пикселям.
Private Function LengthMmToModel(ByVal valueMm As Double, ByVal lengthUnit As String) As Double
    LengthMmToModel = mUnits.InternalLengthToOutput(valueMm, lengthUnit)
End Function

' Предварительно оценивает размеры текста и зазора в координатах модели
' для расширения границ графика. Точные экранные размеры вычисляются
' позже по фактической области построения PlotArea.
Private Function AnnotationTextMetricModel(ByVal value As Double, ByVal unitName As String, _
        ByVal lengthUnit As String, ByVal rangeX As Double, ByVal rangeY As Double) As Double
    If StrComp(Trim$(unitName), "pt", vbTextCompare) <> 0 Then
        AnnotationTextMetricModel = LengthMmToModel(value, lengthUnit)
        Exit Function
    End If

    Dim plotW As Double, plotH As Double
    plotW = PlotCoordinateWidth()
    plotH = PlotCoordinateHeight()
    If plotW <= 0# Or plotH <= 0# Then
        plotW = mChartObject.Width * 0.72
        plotH = mChartObject.Height * 0.72
    End If

    Dim scaleValue As Double
    scaleValue = MinDouble(plotW / MaxDouble(rangeX, 1#), plotH / MaxDouble(rangeY, 1#))
    If scaleValue <= 0# Then Exit Function
    AnnotationTextMetricModel = value / scaleValue
End Function

' Прячет штатные оси Excel-графика: схема использует собственные осевые линии,
' размерные аннотации и подписи.
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

' Определяет диапазон значений Stress/Strain для окраски элементов и легенды.
' Неактивный растянутый бетон исключается, чтобы "нет" не расширяло шкалу.
Private Sub CalculateResultRange(ByVal data As CSectionPlotDataReader)
    Dim i As Long
    mHasCompression = False
    mHasTension = False
    mCompressionMaxAbs = 0#
    mTensionMaxAbs = 0#
    mResultMaxAbs = 0#
    ResetMaterialLegendRanges
    For i = 1 To data.Count
        If StrComp(data.PhysicalState(i), "InactiveTensionConcrete", vbTextCompare) <> 0 Then
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
End Sub

' Очищает накопленные пределы раздельной легенды бетона и арматуры.
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

' Добавляет значение элемента в соответствующую подшкалу легенды материала:
' бетон/арматура и сжатие/растяжение хранятся отдельно.
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

' Запоминает значение с максимальным модулем для одной группы легенды.
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

' Рисует элементы с их фактическими размерами отдельными фигурами.
' Только для очень больших сеток использует группы маркеров диаграммы
' одинакового оформления. Если групп больше лимита Excel, выводит фигуры,
' сохраняя ступени цветов большой сетки и координаты каждого элемента.
Private Sub AddElementSeries(ByVal data As CSectionPlotDataReader, ByVal settings As CSystemSettingsReader)
    If mElementShapeLayerEnabled Then
        CapturePlotAreaMetrics
        AddElementShapes data
        Exit Sub
    End If

    Dim buckets As Object
    Set buckets = CreateObject("Scripting.Dictionary")
    buckets.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To data.Count ' Объединяем одинаково оформленные элементы в одну Chart series; координаты и расчетные значения отдельных элементов не меняются.
        Dim key As String
        key = ElementBucketKey(data, i)
        Dim indexList As Collection
        If buckets.Exists(key) Then
            Set indexList = buckets.Item(key)
        Else
            Set indexList = New Collection
            buckets.Add key, indexList
            If buckets.Count > ELEMENT_SERIES_MAX_COUNT Then
                mElementShapeLayerEnabled = True ' До создания первого ряда переключаем всю сетку, чтобы не оставлять частично нарисованные элементы.
                ConfigureEqualScale data, settings ' Фигуры используют устойчивую рамку ChartObject; заново согласуем с ней одинаковый масштаб X/Y.
                CapturePlotAreaMetrics
                AddElementShapes data
                Exit Sub
            End If
        End If
        indexList.Add i
    Next i

    Dim keyVariant As Variant
    For Each keyVariant In buckets.Keys
        AddBucketSeriesByIndexes data, CStr(keyVariant), buckets(keyVariant)
    Next keyVariant

    CapturePlotAreaMetrics
End Sub

' Проверяет, можно ли рисовать расчетные элементы отдельными Shapes.
' Это обход лимита Excel Chart по числу series: геометрия остается поэлементной
' на средних сетках, а укрупнение включается только при действительно тяжелом
' количестве объектов.
Private Function ShouldDrawElementShapeLayer(ByVal data As CSectionPlotDataReader) As Boolean
    ShouldDrawElementShapeLayer = (data.Count <= ELEMENT_SHAPE_RENDER_MAX_COUNT)
    If ShouldDrawElementShapeLayer Then Exit Function
    Dim i As Long
    For i = 1 To data.Count ' Аналитическую границу нельзя заменить квадратным маркером большой сетки.
        If data.ConcreteHasBoundary(i) Then ShouldDrawElementShapeLayer = True: Exit Function
    Next i
End Function

' Рисует каждый бетонный элемент и стержень отдельной фигурой без усреднения размеров.
' Цвет выбирает по тем же правилам, что при групповой отрисовке больших сеток.
Private Sub AddElementShapes(ByVal data As CSectionPlotDataReader)
    Dim fallbackSquareRotation As Double
    fallbackSquareRotation = AverageKnownConcreteElementRotation(data)

    Dim groups As Object: Set groups = CreateObject("Scripting.Dictionary")
    Dim itemIndex As Long, colorValue As Long, key As String, shapeObject As Object, indexes As Collection
    Dim shapeIndex As Long: shapeIndex = mChart.Shapes.Count
    For itemIndex = 1 To data.Count
        Set shapeObject = Nothing
        colorValue = ElementColor(data, itemIndex)
        If StrComp(data.MaterialType(itemIndex), "Rebar", vbTextCompare) = 0 Then
            Set shapeObject = AddRebarElementShape(data, itemIndex, colorValue)
            key = "R:" & CStr(colorValue)
        Else
            Set shapeObject = AddConcreteElementShape(data, itemIndex, fallbackSquareRotation, colorValue)
            key = "C:" & CStr(colorValue)
            If data.ConcreteHasBoundary(itemIndex) Then key = "F:" & CStr(colorValue)
        End If
        If Not shapeObject Is Nothing Then
            shapeIndex = shapeIndex + 1
            If groups.Exists(key) Then
                Set indexes = groups.Item(key)
            Else
                Set indexes = New Collection: groups.Add key, indexes
            End If
            indexes.Add shapeIndex
        End If
    Next itemIndex

    Dim groupKey As Variant, lineColor As Long, weight As Double
    For Each groupKey In groups.Keys ' Координаты и порядок фигур индивидуальны; одинаковое оформление применяем сразу ко всему набору, сокращая обращения к Excel.
        colorValue = CLng(Mid$(CStr(groupKey), 3))
        If Left$(CStr(groupKey), 1) = "R" Then
            lineColor = RGB(0, 0, 0): weight = 0.55
        Else
            lineColor = colorValue: weight = 0.15
        End If
        FormatElementShapes groups.Item(groupKey), colorValue, lineColor, weight
        If Left$(CStr(groupKey), 1) = "F" Then HideElementShapeLines groups.Item(groupKey) ' Двойные нулевые перемычки составной фигуры не должны рисовать цветные швы через отверстия.
    Next groupKey
End Sub

' Убирает только обводку заполненных составных областей: отверстия задаются
' противоположным обходом, а двойные мосты соединяют loops без дополнительной площади.
Private Sub HideElementShapeLines(ByVal indexes As Collection)
    Dim values() As Variant, i As Long
    ReDim values(1 To indexes.Count)
    For i = 1 To indexes.Count: values(i) = CLng(indexes(i)): Next i
    mChart.Shapes.Range(values).Line.Visible = False
End Sub

' Применяет прежние заливку и обводку ко всем фигурам одного оформления.
' ShapeRange не объединяет фигуры и не меняет их имена, геометрию или порядок.
Private Sub FormatElementShapes(ByVal indexes As Collection, ByVal fillColor As Long, ByVal lineColor As Long, ByVal weight As Double)
    Dim values() As Variant, i As Long, shapes As Object
    ReDim values(1 To indexes.Count)
    For i = 1 To indexes.Count: values(i) = CLng(indexes(i)): Next i
    Set shapes = mChart.Shapes.Range(values)
    shapes.Fill.ForeColor.RGB = fillColor
    shapes.Line.ForeColor.RGB = lineColor
    shapes.Line.Weight = weight
End Sub

' Рисует бетонный элемент по сохранённым ширине, высоте и углу.
' Если при импорте известна только площадь, использует эквивалентный квадрат
' с сохранённой средней ориентацией сетки.
Private Function AddConcreteElementShape(ByVal data As CSectionPlotDataReader, ByVal index As Long, _
        ByVal fallbackSquareRotation As Double, ByVal colorValue As Long) As Object
    If data.ConcreteHasBoundary(index) Then
        Dim geometry As CConcreteRegion, query As CSectionGeometryQuery
        Dim cx As Double, cy As Double, rectangleWidth As Double, rectangleHeight As Double, rectangleRotation As Double
        Set geometry = data.ConcreteGeometry(index): Set query = New CSectionGeometryQuery
        If query.TryRectangle(geometry, cx, cy, rectangleWidth, rectangleHeight, rectangleRotation) Then
            Set AddConcreteElementShape = DrawModelRectangleCentered(cx, cy, rectangleWidth * CONCRETE_SHAPE_FILL_RATIO, _
                rectangleHeight * CONCRETE_SHAPE_FILL_RATIO, rectangleRotation, colorValue, colorValue, 0.15, "ElementConcrete", True)
        Else
            Set AddConcreteElementShape = DrawConcreteRegionShape(geometry, data.X(index), data.Y(index))
        End If
        Exit Function
    End If
    Dim widthValue As Double
    Dim heightValue As Double
    widthValue = ConcreteElementWidth(data, index) * CONCRETE_SHAPE_FILL_RATIO
    heightValue = ConcreteElementHeight(data, index) * CONCRETE_SHAPE_FILL_RATIO
    If widthValue <= 0# Or heightValue <= 0# Then Exit Function

    Dim rotationValue As Double
    rotationValue = ConcreteElementVisualRotation(data, index, fallbackSquareRotation)
    Set AddConcreteElementShape = DrawModelRectangleCentered(data.X(index), data.Y(index), widthValue, heightValue, _
        rotationValue, colorValue, colorValue, 0.15, "ElementConcrete", True)
End Function

' Выводит всю область одним заполненным Shape с визуальным коэффициентом 0,7.
' Противоположные обходы отверстий соединены двойными мостами нулевой площади;
' расчетные границы, центры и свойства при этом не изменяются.
Private Function DrawConcreteRegionShape(ByVal region As CConcreteRegion, ByVal centerX As Double, ByVal centerY As Double) As Object
    If region Is Nothing Then Err.Raise vbObjectError + 4813, "CSectionPlotter", "Не передана граница бетонного элемента."
    Dim edges() As TRegionEdge, count As Long, i As Long, previousLoop As Long
    region.CopyEdges edges, count
    If count = 0 Then Err.Raise vbObjectError + 4813, "CSectionPlotter", "Бетонный элемент имеет пустую границу."
    Dim anchorX As Double, anchorY As Double, builder As Object
    anchorX = edges(1).X1: anchorY = edges(1).Y1
    Set builder = mChart.Shapes.BuildFreeform(MSO_EDITING_CORNER, _
        RegionChartX(anchorX, centerX), RegionChartY(anchorY, centerY))
    For i = 1 To count
        If edges(i).LoopID <> previousLoop Then
            If previousLoop > 0 Then
                builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_CORNER, RegionChartX(anchorX, centerX), RegionChartY(anchorY, centerY)
                builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_CORNER, RegionChartX(edges(i).X1, centerX), RegionChartY(edges(i).Y1, centerY)
            End If
            previousLoop = edges(i).LoopID
        End If
        AppendRegionShapeEdge builder, edges(i), centerX, centerY
    Next i
    If region.LoopCount > 1 Then builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_CORNER, RegionChartX(anchorX, centerX), RegionChartY(anchorY, centerY)
    Dim shape As Object: Set shape = builder.ConvertToShape
    shape.Name = SHAPE_PREFIX & "ElementConcrete" & CStr(mChart.Shapes.Count)
    Set DrawConcreteRegionShape = shape
End Function

' Передает прямую либо кубические участки аналитической дуги в Freeform.
' Каждый участок не длиннее четверти окружности; аппроксимация касается только рисунка.
Private Sub AppendRegionShapeEdge(ByVal builder As Object, ByRef edge As TRegionEdge, _
        ByVal centerX As Double, ByVal centerY As Double)
    If Abs(edge.Sweep) <= 0.000000000001 Then
        builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_CORNER, RegionChartX(edge.X2, centerX), RegionChartY(edge.Y2, centerY)
        Exit Sub
    End If
    Dim pieces As Long, i As Long, radius As Double, startAngle As Double, angle As Double, following As Double
    Dim sweep As Double, tangent As Double, x As Double, y As Double, nextX As Double, nextY As Double
    pieces = CLng(-Int(-Abs(edge.Sweep) / (GEOM_PI / 2#)))
    If pieces < 1 Then pieces = 1
    radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
    startAngle = Atn2Local(edge.Y1 - edge.CenterY, edge.X1 - edge.CenterX)
    sweep = edge.Sweep / CDbl(pieces): tangent = 4# / 3# * Tan(sweep / 4#)
    For i = 0 To pieces - 1
        angle = startAngle + CDbl(i) * sweep: following = angle + sweep
        x = edge.CenterX + radius * Cos(angle): y = edge.CenterY + radius * Sin(angle)
        nextX = edge.CenterX + radius * Cos(following): nextY = edge.CenterY + radius * Sin(following)
        If i = pieces - 1 Then nextX = edge.X2: nextY = edge.Y2
        builder.AddNodes MSO_SEGMENT_CURVE, MSO_EDITING_CORNER, _
            RegionChartX(x - radius * tangent * Sin(angle), centerX), RegionChartY(y + radius * tangent * Cos(angle), centerY), _
            RegionChartX(nextX + radius * tangent * Sin(following), centerX), RegionChartY(nextY - radius * tangent * Cos(following), centerY), _
            RegionChartX(nextX, centerX), RegionChartY(nextY, centerY)
    Next i
End Sub

' Масштабирует только рисунок относительно центра элемента, затем переводит X в Chart.
Private Function RegionChartX(ByVal x As Double, ByVal centerX As Double) As Double
    RegionChartX = ModelToChartX(centerX + (x - centerX) * CONCRETE_SHAPE_FILL_RATIO)
End Function

' Масштабирует только рисунок относительно центра элемента, затем переводит Y в Chart.
Private Function RegionChartY(ByVal y As Double, ByVal centerY As Double) As Double
    RegionChartY = ModelToChartY(centerY + (y - centerY) * CONCRETE_SHAPE_FILL_RATIO)
End Function

' Рисует арматурный стержень по диаметру из Results; если диаметр не сохранен,
' используется эквивалентный круг по площади.
Private Function AddRebarElementShape(ByVal data As CSectionPlotDataReader, ByVal index As Long, ByVal colorValue As Long) As Object
    Dim diameterValue As Double
    diameterValue = data.Diameter(index)
    If diameterValue <= 0# Then diameterValue = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
    If diameterValue <= 0# Then Exit Function

    Set AddRebarElementShape = DrawModelOval(data.X(index) - diameterValue / 2#, data.Y(index) - diameterValue / 2#, _
        data.X(index) + diameterValue / 2#, data.Y(index) + diameterValue / 2#, _
        RGB(0, 0, 0), 0.55, "ElementRebar", True, colorValue, True)
End Function

' Формирует ключ группировки точек диаграммы: элементы с одинаковым маркером,
' размером, цветом, материалом и физическим состоянием можно рисовать одной
' chart-серией, что сильно уменьшает число Excel-объектов.
Private Function ElementBucketKey(ByVal data As CSectionPlotDataReader, ByVal index As Long) As String
    Dim markerStyle As Long
    Dim markerSize As Long
    markerStyle = ElementMarkerStyle(data, index)
    markerSize = ElementMarkerSize(data, index)
    If StrComp(data.MaterialType(index), "Rebar", vbTextCompare) <> 0 Then markerSize = BucketConcreteMarkerSize(markerSize)

    Dim colorValue As Long
    colorValue = ElementColor(data, index)
    ElementBucketKey = CStr(markerStyle) & "|" & CStr(markerSize) & "|" & CStr(colorValue) & "|" & _
        data.MaterialType(index) & "|" & data.PhysicalState(index)
End Function

' Возвращает тип маркера Excel Chart для бетонного волокна или стержня.
Private Function ElementMarkerStyle(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Long
    If StrComp(data.MaterialType(index), "Rebar", vbTextCompare) = 0 Then
        ElementMarkerStyle = XL_MARKER_CIRCLE
    Else
        ElementMarkerStyle = XL_MARKER_SQUARE
    End If
End Function

' Возвращает размер маркера grouped-series в points.
Private Function ElementMarkerSize(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Long
    If StrComp(data.MaterialType(index), "Rebar", vbTextCompare) = 0 Then
        ElementMarkerSize = RoundToLong(ClampDouble(data.Diameter(index) * ChartScale(), 4#, 42#))
    Else
        ElementMarkerSize = RoundToLong(ClampDouble(0.7 * ConcreteMarkerLength(data, index) * ChartScale(), 2#, 48#))
    End If
End Function

' Создаёт один ряд XY-диаграммы для одинаково оформленных элементов.
' Координаты сохраняются в единицах схемы; Excel размещает маркеры по осям.
Private Sub AddBucketSeriesByIndexes(ByVal data As CSectionPlotDataReader, ByVal bucketKey As String, ByVal indexes As Collection)
    If indexes Is Nothing Then Exit Sub
    If indexes.Count <= 0 Then Exit Sub

    Dim parts As Variant
    parts = Split(bucketKey, "|")
    If UBound(parts) < 4 Then _
        Err.Raise vbObjectError + 4735, "CSectionPlotter", "Группа элементов схемы не содержит всех параметров оформления."
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
    If StrComp(CStr(parts(3)), "Rebar", vbTextCompare) = 0 Then
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
    t = BucketResultColorRatio(t)

    If StrComp(state, "Compression", vbTextCompare) = 0 Then
        MaterialResultColor = BlendColor(RGB(218, 218, 218), LegendCompressionColor(materialType), t)
    ElseIf StrComp(state, "Tension", vbTextCompare) = 0 Then
        MaterialResultColor = BlendColor(RGB(218, 218, 218), LegendTensionColor(materialType), t)
    Else
        MaterialResultColor = RGB(150, 150, 150)
    End If
End Function

' Укрупняет бетонные маркеры до нескольких близких размеров. На мелкой сетке
' это визуально почти не меняет схему, но резко снижает число Excel series.
Private Function BucketConcreteMarkerSize(ByVal markerSize As Long) As Long
    If markerSize <= 2 Then
        BucketConcreteMarkerSize = 2
    Else
        BucketConcreteMarkerSize = RoundToLong(markerSize / CONCRETE_MARKER_SIZE_BUCKET) * CONCRETE_MARKER_SIZE_BUCKET
        BucketConcreteMarkerSize = RoundToLong(ClampDouble(BucketConcreteMarkerSize, 2#, 48#))
    End If
End Function

' Переводит непрерывный градиент в ограниченное число цветовых ступеней.
' Численное значение элемента остается точным в Results; дискретизация нужна
' только для устойчивой отрисовки большой расчетной сетки средствами Excel Chart.
Private Function BucketResultColorRatio(ByVal t As Double) As Double
    If Not mElementGroupingEnabled Then
        BucketResultColorRatio = ClampDouble(t, 0#, 1#)
    ElseIf RESULT_COLOR_BUCKET_COUNT <= 1 Then
        BucketResultColorRatio = ClampDouble(t, 0#, 1#)
    Else
        BucketResultColorRatio = RoundToLong(ClampDouble(t, 0#, 1#) * RESULT_COLOR_BUCKET_COUNT) / RESULT_COLOR_BUCKET_COUNT
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

' Добавляет верхний служебный оверлей поверх элементов и аннотаций.
' Нейтральная линия, главные оси и точка нагрузки рисуются Chart.Shape-ами
' через тот же model-to-chart mapper, что и элементы, поэтому их координаты
' остаются согласованы, а слой можно поднять поверх сетки.
Private Sub DrawAnalysisOverlay(ByVal data As CSectionPlotDataReader)
    If mNeutralLineEnabled Then DrawNeutralLine data
    If PrincipalAxesModeDraws(mPrincipalAxesMode) Then DrawPrincipalAxes data, mPrincipalAxesMode
    If mLoadPointEnabled Then DrawPointMarker data.LoadReferenceX, data.LoadReferenceY, 10, RGB(255, 128, 0), MSO_SHAPE_DIAMOND, "LoadPoint"
End Sub

' Рисует доступные справочные оси и точку нагрузки при просмотре геометрии
' без НДС. Координаты берёт из сохранённых свойств, оформление — из настроек схемы.
Private Sub DrawGeometryReferenceOverlay(ByVal data As CSectionPlotDataReader)
    If Not data.GeometryReferenceItemsEnabled Then Exit Sub
    If PrincipalAxesModeDraws(mPrincipalAxesMode) Then DrawPrincipalAxes data, mPrincipalAxesMode
    If mLoadPointEnabled Then DrawPointMarker data.LoadReferenceX, data.LoadReferenceY, 10, RGB(255, 128, 0), MSO_SHAPE_DIAMOND, "LoadPoint"
End Sub

' Рисует контуры, размеры, подписи и легенды фигурами диаграммы.
' Расчётные значения берёт только из сохранённых Results.
Private Sub DrawShapeDecorations(ByVal data As CSectionPlotDataReader)
    If mContourEnabled Then DrawConcreteContour data
    DrawShapeAnnotations data, mAnnotationLayout.DimensionsEnabled, mAnnotationLayout.RebarLabelsEnabled
    If data.GeometryOnly Then
        DrawGeometryOnlyNote data
        Exit Sub
    End If
    If mResultLabelsEnabled Then DrawResultLabels data, mResultLabelSpacing
    If mLegendEnabled Then DrawLegend data
    DrawStateWarning data
End Sub

' Пишет под схемой короткое пояснение, почему у geometry-only схемы нет
' цветовой шкалы, нейтральной линии и поэлементных Stress/Strain.
Private Sub DrawGeometryOnlyNote(ByVal data As CSectionPlotDataReader)
    If Len(Trim$(data.GeometryOnlyNote)) = 0 Then Exit Sub

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthValue As Double
    EnsurePlotAreaMetrics
    leftPos = mPlotLeft
    topPos = mPlotTop + mPlotHeight + 16
    widthValue = mPlotWidth
    AddChartText leftPos, topPos, widthValue, 62, data.GeometryOnlyNote, 0#, RGB(160, 40, 40), True, True, 8.5
End Sub

' Рисует единственный заголовок схемы как Chart.Shape. Shape-заголовок живет
' в той же системе координат Chart, что и элементы, оси и аннотации.
Private Sub DrawPlotTitleOverlay(ByVal titleText As String)
    If Len(Trim$(titleText)) = 0 Then Exit Sub

    Dim titleWidth As Double
    titleWidth = mChartObject.Width - 16#
    If titleWidth <= 0# Then titleWidth = StablePlotFrameWidth()

    Dim titleShape As Object
    Set titleShape = AddChartText(8#, PLOT_TITLE_TOP, titleWidth, PLOT_TITLE_HEIGHT, _
        titleText, 0#, RGB(20, 20, 20), True, True, 12#)
    titleShape.Name = SHAPE_PREFIX & "Title"
End Sub

' Добавляет маленькие курсивные подписи глобальных осей X/Y внутри рабочей
' области схемы. Подписи живут в Chart.Shapes, поэтому двигаются вместе с
' окном схемы и не попадают в правую область легенд.
Private Sub DrawAxisNameLabels()
    EnsurePlotAreaMetrics
    If mPlotWidth <= 0# Or mPlotHeight <= 0# Then Exit Sub

    Dim textHeight As Double
    Dim xWidth As Double
    Dim yHeight As Double
    textHeight = MaxDouble(AXIS_LABEL_TEXT_HEIGHT, mAxisLabelFontSize + 6#)
    xWidth = MaxDouble(AXIS_LABEL_X_WIDTH, mAxisLabelFontSize * 7#)
    yHeight = MaxDouble(AXIS_LABEL_Y_HEIGHT, mAxisLabelFontSize * 7#)

    Dim xLeft As Double
    Dim xTop As Double
    xLeft = mPlotLeft + mPlotWidth - xWidth - AXIS_LABEL_MARGIN
    xTop = mPlotTop + mPlotHeight - textHeight - AXIS_LABEL_MARGIN
    If xLeft < mPlotLeft + AXIS_LABEL_MARGIN Then xLeft = mPlotLeft + AXIS_LABEL_MARGIN
    If xTop < mPlotTop + AXIS_LABEL_MARGIN Then xTop = mPlotTop + AXIS_LABEL_MARGIN

    Dim xLabel As Object
    Set xLabel = AddChartText(xLeft, xTop, xWidth, textHeight, _
        "Ось X", 0#, RGB(80, 80, 80), False, True, mAxisLabelFontSize, False, True)
    xLabel.Name = SHAPE_PREFIX & "AxisLabelX"
    BringShapeToFront xLabel

    Dim yLeft As Double
    Dim yTop As Double
    yLeft = mPlotLeft + AXIS_LABEL_MARGIN
    yTop = mPlotTop + AXIS_LABEL_MARGIN

    Dim yLabel As Object
    Set yLabel = AddChartText(yLeft, yTop, textHeight, yHeight, _
        "Ось Y", 90#, RGB(80, 80, 80), False, True, mAxisLabelFontSize, True, True)
    yLabel.Name = SHAPE_PREFIX & "AxisLabelY"
    BringShapeToFront yLabel
End Sub

' Показывает под схемой короткое красное предупреждение по сохраненному
' DirectStateStatus. Для FAIL состояние найдено
' вне физической диаграммы. Для NumFail отдельно предупреждаем, что на схеме
' может быть последняя итерация, а не корректное равновесное состояние.
Private Sub DrawStateWarning(ByVal data As CSectionPlotDataReader)
    Dim warningText As String
    Dim policy As CResultStatusPolicy
    Set policy = New CResultStatusPolicy

    If data.PhysicalDiagramWarningRequired Then
        warningText = EXTENSION_WARNING_TEXT
    ElseIf policy.ToUserStatus(data.DirectStateStatus) = policy.NumFail Then
        warningText = NUMERICAL_STATE_WARNING_TEXT
    End If
    If Len(warningText) = 0 Then Exit Sub

    Dim leftPos As Double
    Dim topPos As Double
    Dim widthValue As Double
    EnsurePlotAreaMetrics
    leftPos = mPlotLeft
    topPos = mPlotTop + mPlotHeight + 8
    widthValue = mPlotWidth
    If widthValue <= 0# Then widthValue = mChartObject.Width - 40

    AddChartText leftPos, topPos, widthValue, 18, warningText, 0#, RGB(192, 0, 0), True, True, 10#
End Sub

' Рисует отрезки и дуги контура бетона, прочитанные из Results.
' Если при импорте точный контур не был задан, отсутствующие границы не угадываются.
Private Sub DrawConcreteContour(ByVal data As CSectionPlotDataReader)
    Dim i As Long
    For i = 1 To data.AnnotationCount
        If StrComp(data.AnnotationType(i), "CONTOUR_LINE", vbTextCompare) = 0 Then
            DrawModelLine data.AnnotationStartX(i), data.AnnotationStartY(i), _
                data.AnnotationEndX(i), data.AnnotationEndY(i), RGB(25, 25, 25), mContourLineWeight, "Contour"
        ElseIf StrComp(data.AnnotationType(i), "CONTOUR_CIRCLE", vbTextCompare) = 0 Then
            DrawModelOval data.AnnotationStartX(i) - data.AnnotationEndX(i), _
                data.AnnotationStartY(i) - data.AnnotationEndX(i), _
                data.AnnotationStartX(i) + data.AnnotationEndX(i), _
                data.AnnotationStartY(i) + data.AnnotationEndX(i), RGB(25, 25, 25), mContourLineWeight, "ContourCircle", False
        ElseIf StrComp(data.AnnotationType(i), "CONTOUR_ARC", vbTextCompare) = 0 Then
            DrawContourArc data.AnnotationNormalX(i), data.AnnotationNormalY(i), _
                data.AnnotationStartX(i), data.AnnotationStartY(i), _
                data.AnnotationEndX(i), data.AnnotationEndY(i), _
                data.AnnotationSweepAngle(i), RGB(25, 25, 25), mContourLineWeight
        End If
    Next i
End Sub

' Рисует сохранённую аналитическую дугу ломаной для Excel Chart.
' Шаг не превышает 5 градусов, число отрезков не меньше шести;
' дискретизация влияет только на рисунок, а не на расчётную геометрию.
Private Sub DrawContourArc(ByVal centerX As Double, ByVal centerY As Double, _
        ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal sweepAngle As Double, ByVal colorValue As Long, ByVal weight As Double)
    Dim radius As Double
    radius = Sqr((startX - centerX) * (startX - centerX) + (startY - centerY) * (startY - centerY))
    If radius <= 0# Or Abs(sweepAngle) <= 0.000000000001 Then
        DrawModelLine startX, startY, endX, endY, colorValue, weight, "ContourArcLine"
        Exit Sub
    End If

    Dim startAngle As Double
    startAngle = Atn2Local(startY - centerY, startX - centerX)

    Dim segmentCount As Long
    segmentCount = CLng(Int(Abs(sweepAngle) / (GEOM_PI / 36#)))
    If Abs(sweepAngle) > segmentCount * (GEOM_PI / 36#) + 0.000000000001 Then segmentCount = segmentCount + 1
    If segmentCount < 6 Then segmentCount = 6

    Dim i As Long
    Dim nextX As Double, nextY As Double
    Dim angleValue As Double
    Dim builder As Object, arcShape As Object
    Set builder = mChart.Shapes.BuildFreeform(MSO_EDITING_AUTO, ModelToChartX(startX), ModelToChartY(startY))
    For i = 1 To segmentCount
        angleValue = startAngle + sweepAngle * CDbl(i) / CDbl(segmentCount)
        If i = segmentCount Then
            nextX = endX
            nextY = endY
        Else
            nextX = centerX + radius * Cos(angleValue)
            nextY = centerY + radius * Sin(angleValue)
        End If
        builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_AUTO, ModelToChartX(nextX), ModelToChartY(nextY)
    Next i
    Set arcShape = builder.ConvertToShape
    arcShape.Name = SHAPE_PREFIX & "ContourArcLine" & CStr(mChart.Shapes.Count)
    arcShape.Fill.Visible = False
    arcShape.Line.Visible = True
    arcShape.Line.ForeColor.RGB = colorValue
    arcShape.Line.Weight = weight
End Sub

' Строит нейтральную линию по сохранённому полю деформаций сочетания.
' Обрезает её по сохранённому контуру сечения. Если контура нет,
' определяет границы по бетонным элементам.
Private Sub DrawNeutralLine(ByVal data As CSectionPlotDataReader)
    Dim a As Double, b As Double, c As Double
    a = data.KappaY ' Нулевая деформация записана как a*x + b*y + c = 0; коэффициенты переставлены относительно kappaX/kappaY.
    b = data.KappaX
    c = data.Epsilon0
    If Abs(a) + Abs(b) < 0.000000000001 Then Exit Sub

    Dim dx As Double, dy As Double
    dx = -b ' Касательная к нейтральной линии перпендикулярна нормали (a, b).
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

    DrawModelReferenceLine x0 + minT * dx, y0 + minT * dy, x0 + maxT * dx, y0 + maxT * dy, _
        RGB(105, 105, 105), 0.7, "NeutralAxis"
End Sub

' Находит пересечения нейтральной линии с сохранённым контуром сечения.
' Использует точные исходные границы независимо от способа формирования сечения.
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
        ElseIf StrComp(data.AnnotationType(i), "CONTOUR_ARC", vbTextCompare) = 0 Then
            AddArcLineIntersectionT x0, y0, dirX, dirY, data.AnnotationNormalX(i), data.AnnotationNormalY(i), _
                data.AnnotationStartX(i), data.AnnotationStartY(i), data.AnnotationSweepAngle(i), values, count
        End If
    Next i
End Sub

' Резервно обрезает линию по габариту центров бетонных элементов.
' Используется при отсутствии точного наружного контура, в том числе у импортированной сетки.
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
    tCenter = vx * dirX + vy * dirY ' Проекция центра на единичное направление линии, мм.
    d2 = vx * vx + vy * vy - tCenter * tCenter ' Квадрат расстояния от центра окружности до линии, мм2.
    If d2 > radius * radius + 0.0000001 Then Exit Sub
    If d2 < 0# Then d2 = 0# ' Отрицательный остаток возможен из-за округления разности почти равных квадратов.
    halfChord = Sqr(MaxDouble(radius * radius - d2, 0#))
    AddUniqueLineT tCenter - halfChord, values, count
    AddUniqueLineT tCenter + halfChord, values, count
End Sub

' Добавляет пересечения расчетной линии с дуговым сегментом контура.
' Сначала решается пересечение с полной окружностью, затем точки фильтруются
' по знаковому углу обхода дуги от её начальной точки.
Private Sub AddArcLineIntersectionT(ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal centerX As Double, _
        ByVal centerY As Double, ByVal arcStartX As Double, ByVal arcStartY As Double, _
        ByVal sweepAngle As Double, ByRef values() As Double, ByRef count As Long)
    Dim radius As Double
    radius = Sqr((arcStartX - centerX) * (arcStartX - centerX) + (arcStartY - centerY) * (arcStartY - centerY))
    If radius <= 0# Or Abs(sweepAngle) <= 0.000000000001 Then Exit Sub

    Dim vx As Double, vy As Double, tCenter As Double, d2 As Double, halfChord As Double
    vx = centerX - x0
    vy = centerY - y0
    tCenter = vx * dirX + vy * dirY ' Проекция центра на единичное направление линии, мм.
    d2 = vx * vx + vy * vy - tCenter * tCenter ' Квадрат расстояния от центра окружности до линии, мм2.
    If d2 > radius * radius + 0.0000001 Then Exit Sub
    If d2 < 0# Then d2 = 0# ' Отрицательный остаток возможен из-за округления разности почти равных квадратов.
    halfChord = Sqr(MaxDouble(radius * radius - d2, 0#))
    AddArcIntersectionCandidate x0, y0, dirX, dirY, centerX, centerY, arcStartX, arcStartY, _
        sweepAngle, tCenter - halfChord, values, count
    AddArcIntersectionCandidate x0, y0, dirX, dirY, centerX, centerY, arcStartX, arcStartY, _
        sweepAngle, tCenter + halfChord, values, count
End Sub

' Проверяет вычисленную точку пересечения с окружностью
' и сохраняет её, только если она лежит на нужной части дуги.
Private Sub AddArcIntersectionCandidate(ByVal x0 As Double, ByVal y0 As Double, _
        ByVal dirX As Double, ByVal dirY As Double, ByVal centerX As Double, _
        ByVal centerY As Double, ByVal arcStartX As Double, ByVal arcStartY As Double, _
        ByVal sweepAngle As Double, ByVal tValue As Double, ByRef values() As Double, ByRef count As Long)
    Dim x As Double, y As Double
    x = x0 + tValue * dirX
    y = y0 + tValue * dirY
    If IsPointOnArcSweep(x, y, centerX, centerY, arcStartX, arcStartY, sweepAngle) Then
        AddUniqueLineT tValue, values, count
    End If
End Sub

' True, если точка находится на signed sweep дуги от ее начального радиуса.
Private Function IsPointOnArcSweep(ByVal x As Double, ByVal y As Double, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal arcStartX As Double, _
        ByVal arcStartY As Double, ByVal sweepAngle As Double) As Boolean
    Dim startAngle As Double
    Dim pointAngle As Double
    startAngle = Atn2Local(arcStartY - centerY, arcStartX - centerX)
    pointAngle = Atn2Local(y - centerY, x - centerX)

    Dim delta As Double
    If sweepAngle >= 0# Then
        delta = NormalizePositiveAngle(pointAngle - startAngle)
        IsPointOnArcSweep = (delta <= sweepAngle + 0.000001)
    Else
        delta = NormalizePositiveAngle(startAngle - pointAngle)
        IsPointOnArcSweep = (delta <= -sweepAngle + 0.000001)
    End If
End Function

' Нормализует угол к диапазону [0; 2*pi).
Private Function NormalizePositiveAngle(ByVal angleValue As Double) As Double
    Do While angleValue < 0#
        angleValue = angleValue + 2# * GEOM_PI
    Loop
    Do While angleValue >= 2# * GEOM_PI
        angleValue = angleValue - 2# * GEOM_PI
    Loop
    NormalizePositiveAngle = angleValue
End Function

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

' Рисует выбранные главные центральные оси: бетонного или приведенного сечения.
Private Sub DrawPrincipalAxes(ByVal data As CSectionPlotDataReader, ByVal axesMode As String)
    Dim lengthValue As Double
    lengthValue = 0.65 * Sqr((data.MaxX - data.MinX) * (data.MaxX - data.MinX) + _
        (data.MaxY - data.MinY) * (data.MaxY - data.MinY))
    Dim ca As Double, sa As Double
    Dim angle As Double
    Dim centroidX As Double
    Dim centroidY As Double
    angle = data.PrincipalAxesAngle(axesMode)
    centroidX = data.PrincipalAxesCentroidX(axesMode)
    centroidY = data.PrincipalAxesCentroidY(axesMode)
    ca = Cos(angle)
    sa = Sin(angle)
    DrawModelReferenceLine centroidX - lengthValue * ca, centroidY - lengthValue * sa, _
        centroidX + lengthValue * ca, centroidY + lengthValue * sa, RGB(0, 0, 0), REFERENCE_AXIS_WEIGHT, "Principal1"
    DrawModelReferenceLine centroidX + lengthValue * sa, centroidY - lengthValue * ca, _
        centroidX - lengthValue * sa, centroidY + lengthValue * ca, RGB(0, 0, 0), REFERENCE_AXIS_WEIGHT, "Principal2"
End Sub

' Читает обязательный режим главных осей до очистки существующего Chart.
' Поддержанные текстовые обозначения приводятся к одному из трех вариантов.
Private Function PlotPrincipalAxesMode(ByVal settings As CSystemSettingsReader) As String
    Dim rawValue As String
    rawValue = settings.GetRequiredString("Plot.PrincipalAxesMode")
    PlotPrincipalAxesMode = NormalizePrincipalAxesMode(rawValue, settings)
End Function

' Проверяет выбранный режим осей и возвращает принятое написание.
' Для неизвестного значения сообщает актуальный адрес Config
' и допустимые варианты без подстановки другого режима.
Private Function NormalizePrincipalAxesMode(ByVal rawValue As String, ByVal settings As CSystemSettingsReader) As String
    Select Case LCase$(Trim$(rawValue))
        Case "transformed", "приведенное", "приведенное сечение"
            NormalizePrincipalAxesMode = "Transformed"
        Case "concrete", "бетон", "бетонное", "бетонное сечение"
            NormalizePrincipalAxesMode = "Concrete"
        Case "none", "no", "off", "нет", "не выводить"
            NormalizePrincipalAxesMode = "None"
        Case Else
            Err.Raise vbObjectError + 4718, "CSectionPlotter", _
                settings.InputErrorMessage("Plot.PrincipalAxesMode", "Режим главных осей не распознан.", "Выберите Transformed, Concrete или None.")
    End Select
End Function

' True, если выбранный режим требует вывода осей.
Private Function PrincipalAxesModeDraws(ByVal axesMode As String) As Boolean
    PrincipalAxesModeDraws = (StrComp(axesMode, "None", vbTextCompare) <> 0)
End Function

' Передаёт сохранённые размеры и подписи в CPlotAnnotationLayout.
' Общий класс размещает имеющиеся аннотации независимо от формы сечения;
' отсутствующие при импорте данные самостоятельно не добавляются.
Private Sub DrawShapeAnnotations(ByVal data As CSectionPlotDataReader, _
        ByVal dimensionsEnabled As Boolean, ByVal rebarLabelsEnabled As Boolean)
    If Not dimensionsEnabled And Not rebarLabelsEnabled Then Exit Sub

    Dim layout As CPlotAnnotationLayout
    Set layout = mAnnotationLayout
    EnsurePlotAreaMetrics
    layout.Initialize mMinX, mMaxX, mMinY, mMaxY, mPlotLeft, mPlotTop, _
        mPlotWidth, mPlotHeight, data.LengthUnit, mAnnotationOffsetScale

    Dim i As Long
    If rebarLabelsEnabled Then
        For i = 1 To data.AnnotationCount
            If StrComp(data.AnnotationType(i), "REBAR_ANNOTATION", vbTextCompare) = 0 Or _
                    StrComp(data.AnnotationType(i), "REBAR_GLOBAL_ANNOTATION", vbTextCompare) = 0 Then
                layout.AddRebarLabel data.AnnotationStartX(i), data.AnnotationStartY(i), _
                    data.AnnotationEndX(i), data.AnnotationEndY(i), data.AnnotationNormalX(i), _
                    data.AnnotationNormalY(i), data.AnnotationText(i), data.AnnotationBoundaryDistance(i)
            End If
        Next i
    End If

    If dimensionsEnabled Then
        For i = 1 To data.AnnotationCount
            If StrComp(data.AnnotationType(i), "DIMENSION", vbTextCompare) = 0 Then
                layout.AddDimension data.AnnotationStartX(i), data.AnnotationStartY(i), _
                    data.AnnotationEndX(i), data.AnnotationEndY(i), data.AnnotationNormalX(i), _
                    data.AnnotationNormalY(i), data.AnnotationText(i)
            ElseIf StrComp(data.AnnotationType(i), "MEASUREMENT_LABEL", vbTextCompare) = 0 Or _
                    StrComp(data.AnnotationType(i), "TEXT_ANNOTATION", vbTextCompare) = 0 Then
                layout.AddPointLabel data.AnnotationStartX(i), data.AnnotationStartY(i), data.AnnotationText(i)
            End If
        Next i
    End If

    RenderAnnotationLayout layout
End Sub

' Создаёт фигуры диаграммы по готовому размещению: размерные линии,
' стрелки, выносные линии и текст с уже вычисленными координатами.
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
                layout.Text(i), layout.Rotation(i), layout.Color(i), layout.Bold(i), True, layout.FontSize(i), True
        End If
    Next i
End Sub

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

' Размещает подписи напряжений или деформаций у выбранных бетонных элементов.
' Неработающий растянутый бетон не подписывает. Частоту и расположение
' определяет шаг ячеек, а не порядковый номер элемента.
Private Sub DrawResultLabels(ByVal data As CSectionPlotDataReader, ByVal spacing As Double)
    Dim bestIndexByCell As Object
    Set bestIndexByCell = CreateObject("Scripting.Dictionary")
    Dim bestDistanceByCell As Object
    Set bestDistanceByCell = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To data.Count ' Пространственные подписи выбираются не по ElementID, а по расчетной сетке ячеек. Для каждой ячейки оставляем ближайшее бетонное волокно к ее центру, поэтому подписи распределяются по площади сечения равномернее и не зависят от порядка записи элементов.
        If StrComp(data.MaterialType(i), "Concrete", vbTextCompare) = 0 And _
                StrComp(data.PhysicalState(i), "InactiveTensionConcrete", vbTextCompare) <> 0 Then
            Dim cellX As Double ' Номера ячеек выбора подписей могут быть больше диапазона Long. Мелкий положительный шаг оформления допустим и не меняет расчётные Results.
            Dim cellY As Double
            cellX = Int((data.X(i) - mMinX) / spacing)
            cellY = Int((data.Y(i) - mMinY) / spacing)

            Dim key As String
            key = CStr(cellX) & ":" & CStr(cellY)

            Dim targetX As Double
            Dim targetY As Double
            targetX = mMinX + (cellX + 0.5) * spacing
            targetY = mMinY + (cellY + 0.5) * spacing

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
        DrawTextAt data.X(i), data.Y(i), FormatPlotValue(data.ResultValue(i), data.ResultPrecision), 1#, 0#, RGB(60, 60, 60), False
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

    If IsCommonLegendMode() Then ' Легенда использует те же диапазоны, что и цвет элементов. В режиме Separate бетон и арматура получают независимые шкалы, поэтому их экстремумы не смешиваются визуально.
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
    AddChartText leftPos + 24, topPos + 4, 84, 16, LegendValueText(hasCompression, compressionValue, data.ResultPrecision), 0#, RGB(40, 40, 40), False
    AddChartText leftPos + 24, topPos + barH / 2# - 8, 84, 16, "0", 0#, RGB(40, 40, 40), False
    AddChartText leftPos + 24, topPos + barH - 10, 84, 16, LegendValueText(hasTension, tensionValue, data.ResultPrecision), 0#, RGB(40, 40, 40), False
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

' Формирует подпись min/max легенды; если состояния нет, показывает "нет".
Private Function LegendValueText(ByVal hasValue As Boolean, ByVal value As Double, ByVal precision As Long) As String
    If hasValue Then
        LegendValueText = FormatPlotValue(value, precision)
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

' Рисует Shape-линию по координатам модели.
' Используется для контуров и аннотационных линий, когда нужен именно Chart.Shape.
Private Function DrawModelLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String) As Object
    Set DrawModelLine = DrawChartLine(ModelToChartX(x1), ModelToChartY(y1), ModelToChartX(x2), ModelToChartY(y2), _
        colorValue, weight, suffix)
End Function

' Рисует расчётную опорную линию одной тонкой сплошной Freeform-полосой.
' Для главных осей и нейтральной линии не используем Excel Line/Connector:
' у диагональных линий Excel может нестабильно трактовать рамку объекта.
Private Sub DrawModelReferenceLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String)
    Dim cx1 As Double
    Dim cy1 As Double
    Dim cx2 As Double
    Dim cy2 As Double
    cx1 = ModelToChartX(x1)
    cy1 = ModelToChartY(y1)
    cx2 = ModelToChartX(x2)
    cy2 = ModelToChartY(y2)

    If Not ClipChartLineToPlotFrame(cx1, cy1, cx2, cy2, MaxDouble(weight, REFERENCE_BAND_MIN_THICKNESS) / 2#) Then Exit Sub
    DrawChartReferenceBand cx1, cy1, cx2, cy2, colorValue, weight, suffix
End Sub

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

' Рисует служебную линию внутри ChartObject как тонкую Freeform-полосу.
' Такая форма движется вместе с окном схемы, но не страдает от геометрии
' стандартного Excel Line/Connector на диагональных направлениях.
Private Function DrawChartReferenceBand(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String) As Object
    Dim dx As Double
    Dim dy As Double
    Dim lengthValue As Double
    dx = x2 - x1
    dy = y2 - y1
    lengthValue = Sqr(dx * dx + dy * dy)
    If lengthValue <= 0.0000001 Then Exit Function

    Dim thickness As Double
    thickness = MaxDouble(weight, REFERENCE_BAND_MIN_THICKNESS)
    Dim nx As Double
    Dim ny As Double
    nx = -dy / lengthValue * thickness / 2#
    ny = dx / lengthValue * thickness / 2#

    Dim builder As Object
    Set builder = mChart.Shapes.BuildFreeform(MSO_EDITING_AUTO, x1 + nx, y1 + ny)
    builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_AUTO, x2 + nx, y2 + ny
    builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_AUTO, x2 - nx, y2 - ny
    builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_AUTO, x1 - nx, y1 - ny
    builder.AddNodes MSO_SEGMENT_LINE, MSO_EDITING_AUTO, x1 + nx, y1 + ny

    Dim bandShape As Object
    Set bandShape = builder.ConvertToShape
    bandShape.Name = SHAPE_PREFIX & suffix & CStr(mChart.Shapes.Count)
    bandShape.Fill.ForeColor.RGB = colorValue
    bandShape.Line.Visible = False
    BringShapeToFront bandShape
    Set DrawChartReferenceBand = bandShape
End Function

' Обрезает линию по текущей координатной рамке схемы.
' Это не меняет расчетную геометрию, а только не дает опорным осям выходить
' за видимую область построения.
Private Function ClipChartLineToPlotFrame(ByRef x1 As Double, ByRef y1 As Double, _
        ByRef x2 As Double, ByRef y2 As Double, ByVal inset As Double) As Boolean
    EnsurePlotAreaMetrics

    Dim leftBound As Double
    Dim rightBound As Double
    Dim topBound As Double
    Dim bottomBound As Double
    leftBound = mPlotLeft + inset
    rightBound = mPlotLeft + mPlotWidth - inset
    topBound = mPlotTop + inset
    bottomBound = mPlotTop + mPlotHeight - inset
    If rightBound <= leftBound Or bottomBound <= topBound Then Exit Function

    Dim dx As Double
    Dim dy As Double
    dx = x2 - x1
    dy = y2 - y1

    Dim tMin As Double
    Dim tMax As Double
    tMin = 0#
    tMax = 1#
    If Not ClipLineParameter(-dx, x1 - leftBound, tMin, tMax) Then Exit Function
    If Not ClipLineParameter(dx, rightBound - x1, tMin, tMax) Then Exit Function
    If Not ClipLineParameter(-dy, y1 - topBound, tMin, tMax) Then Exit Function
    If Not ClipLineParameter(dy, bottomBound - y1, tMin, tMax) Then Exit Function
    If tMax < tMin Then Exit Function

    Dim originalX1 As Double
    Dim originalY1 As Double
    originalX1 = x1
    originalY1 = y1
    x1 = originalX1 + dx * tMin
    y1 = originalY1 + dy * tMin
    x2 = originalX1 + dx * tMax
    y2 = originalY1 + dy * tMax
    ClipChartLineToPlotFrame = True
End Function

' Обновляет параметры отсечения Liang-Barsky для одной стороны рамки.
Private Function ClipLineParameter(ByVal p As Double, ByVal q As Double, _
        ByRef tMin As Double, ByRef tMax As Double) As Boolean
    Const CLIP_TOLERANCE As Double = 0.0000001
    If Abs(p) <= CLIP_TOLERANCE Then
        ClipLineParameter = (q >= -CLIP_TOLERANCE)
        Exit Function
    End If

    Dim r As Double
    r = q / p
    If p < 0# Then
        If r > tMax Then Exit Function
        If r > tMin Then tMin = r
    Else
        If r < tMin Then Exit Function
        If r < tMax Then tMax = r
    End If
    ClipLineParameter = True
End Function

' Возвращает Shape в заданный центр после поворота.
' Excel может пересчитать рамку повернутого прямоугольника, поэтому центр
' фиксируем после установки Rotation, а не только при создании объекта.
Private Sub SetShapeCenter(ByVal shapeObject As Object, ByVal centerX As Double, ByVal centerY As Double)
    If shapeObject Is Nothing Then Exit Sub
    shapeObject.Left = centerX - CDbl(shapeObject.Width) / 2#
    shapeObject.Top = centerY - CDbl(shapeObject.Height) / 2#
End Sub

' Поднимает служебные Shape-объекты поверх сетки и аннотаций.
' Если конкретная версия Excel не поддержит ZOrder для объекта, отрисовка
' остается рабочей, просто без дополнительного изменения слоя.
Private Sub BringShapeToFront(ByVal shapeObject As Object)
    If shapeObject Is Nothing Then Exit Sub
    On Error Resume Next
    shapeObject.ZOrder MSO_BRING_TO_FRONT
    On Error GoTo 0
End Sub

' Рисует прямоугольник с центром и поворотом в координатах модели.
' Excel Shape вращается уже в координатах экрана, поэтому знак угла меняется:
' в расчетной модели Y направлена вверх, а в координатах листа Excel - вниз.
Private Function DrawModelRectangleCentered(ByVal centerX As Double, ByVal centerY As Double, _
        ByVal widthModel As Double, ByVal heightModel As Double, ByVal rotationRad As Double, _
        ByVal fillColor As Long, ByVal lineColor As Long, ByVal weight As Double, _
        ByVal suffix As String, Optional ByVal deferFormatting As Boolean = False) As Object
    Dim centerChartX As Double
    Dim centerChartY As Double
    Dim widthValue As Double
    Dim heightValue As Double
    centerChartX = ModelToChartX(centerX)
    centerChartY = ModelToChartY(centerY)
    widthValue = Abs(ModelToChartX(centerX + widthModel / 2#) - ModelToChartX(centerX - widthModel / 2#))
    heightValue = Abs(ModelToChartY(centerY + heightModel / 2#) - ModelToChartY(centerY - heightModel / 2#))
    If widthValue <= 0# Or heightValue <= 0# Then Exit Function

    Dim rectShape As Object
    Set rectShape = mChart.Shapes.AddShape(MSO_SHAPE_RECTANGLE, centerChartX - widthValue / 2#, _
        centerChartY - heightValue / 2#, widthValue, heightValue)
    rectShape.Name = SHAPE_PREFIX & suffix & CStr(mChart.Shapes.Count)
    If Not deferFormatting Then
        rectShape.Fill.ForeColor.RGB = fillColor
        rectShape.Line.ForeColor.RGB = lineColor
        rectShape.Line.Weight = weight
    End If
    If Abs(rotationRad) > 0.000000000001 Then
        rectShape.Rotation = -rotationRad * 180# / GEOM_PI
        SetShapeCenter rectShape, centerChartX, centerChartY
    End If
    Set DrawModelRectangleCentered = rectShape
End Function

' Рисует овал/круг в координатах модели как Shape.
' Нужен для кругового контура и для стержней в поэлементном Shape-режиме.
Private Function DrawModelOval(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal suffix As String, ByVal filled As Boolean, _
        Optional ByVal fillColor As Long = -1, Optional ByVal deferFormatting As Boolean = False) As Object
    Dim leftPos As Double, topPos As Double, widthValue As Double, heightValue As Double
    leftPos = ModelToChartX(MinDouble(x1, x2))
    topPos = ModelToChartY(MaxDouble(y1, y2))
    widthValue = Abs(ModelToChartX(x2) - ModelToChartX(x1))
    heightValue = Abs(ModelToChartY(y2) - ModelToChartY(y1))
    If widthValue <= 0# Or heightValue <= 0# Then Exit Function

    Dim ovalShape As Object
    Set ovalShape = mChart.Shapes.AddShape(MSO_SHAPE_OVAL, leftPos, topPos, widthValue, heightValue)
    ovalShape.Name = SHAPE_PREFIX & suffix & CStr(mChart.Shapes.Count)
    If Not deferFormatting Then
        ovalShape.Line.ForeColor.RGB = colorValue
        ovalShape.Line.Weight = weight
        If filled Then
            If fillColor = -1 Then fillColor = colorValue
            ovalShape.Fill.ForeColor.RGB = fillColor
        Else
            ovalShape.Fill.Visible = False
        End If
    End If
    Set DrawModelOval = ovalShape
End Function

' Рисует одиночный служебный маркер как Shape поверх сетки.
' Положение считается централизованно через ModelToChartX/Y.
Private Sub DrawPointMarker(ByVal xCoord As Double, ByVal yCoord As Double, ByVal sizeValue As Double, _
        ByVal colorValue As Long, ByVal shapeType As Long, ByVal markerName As String)
    Dim leftPos As Double
    Dim topPos As Double
    leftPos = ModelToChartX(xCoord) - sizeValue / 2#
    topPos = ModelToChartY(yCoord) - sizeValue / 2#

    Dim pointShape As Object
    Set pointShape = mChart.Shapes.AddShape(shapeType, leftPos, topPos, sizeValue, sizeValue)
    pointShape.Name = SHAPE_PREFIX & markerName
    pointShape.Fill.ForeColor.RGB = colorValue
    pointShape.Line.ForeColor.RGB = RGB(0, 0, 0)
    BringShapeToFront pointShape
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
        Optional ByVal fontSize As Double = 8#, Optional ByVal useNativeVerticalText As Boolean = False, _
        Optional ByVal isItalic As Boolean = False) As Object
    Dim centerX As Double
    Dim centerY As Double
    centerX = leftPos + widthValue / 2#
    centerY = topPos + heightValue / 2#

    Dim orientation As Long
    orientation = MSO_TEXT_ORIENTATION_HORIZONTAL
    If useNativeVerticalText And Abs(rotation) = 90# Then
        If rotation < 0# Then orientation = MSO_TEXT_ORIENTATION_DOWNWARD Else orientation = MSO_TEXT_ORIENTATION_UPWARD ' Знак вертикального угла задаёт направление чтения, не меняя рамку подписи.
    End If

    Dim textShape As Object
    Set textShape = mChart.Shapes.AddTextbox(orientation, leftPos, topPos, widthValue, heightValue)
    textShape.Name = SHAPE_PREFIX & "Text" & CStr(mChart.Shapes.Count)
    textShape.TextFrame.Characters().Text = textValue
    If fontSize <= 0# Then fontSize = 8#
    textShape.TextFrame.Characters().Font.Size = fontSize
    textShape.TextFrame.Characters().Font.Color = colorValue
    textShape.TextFrame.Characters().Font.Bold = isBold
    textShape.TextFrame.Characters().Font.Italic = isItalic
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
    If rotation <> 0# And orientation = MSO_TEXT_ORIENTATION_HORIZONTAL Then
        textShape.Rotation = rotation
        textShape.Left = centerX - textShape.Width / 2# ' После поворота текста Excel меняет его внешнюю рамку. Возвращаем центр надписи в вычисленную точку, сохраняя заданный TextGap.
        textShape.Top = centerY - textShape.Height / 2#
    End If
    Set AddChartText = textShape
End Function

' Переводит X-координату модели в координату Shape внутри ChartObject.
Private Function ModelToChartX(ByVal xCoord As Double) As Double
    EnsurePlotAreaMetrics
    ModelToChartX = mPlotLeft + (xCoord - mMinX) / (mMaxX - mMinX) * mPlotWidth
End Function

' Переводит Y-координату модели в координату Shape внутри ChartObject.
' Направление Y в Excel идет вниз, поэтому формула инвертирует вертикаль.
Private Function ModelToChartY(ByVal yCoord As Double) As Double
    EnsurePlotAreaMetrics
    ModelToChartY = mPlotTop + mPlotHeight - _
        (yCoord - mMinY) / (mMaxY - mMinY) * mPlotHeight
End Function

' Возвращает единый масштаб модели в points для текущей области построения.
Private Function ChartScale() As Double
    Dim sx As Double, sy As Double
    Dim plotWidth As Double, plotHeight As Double
    If mPlotMetricsCaptured Then
        plotWidth = mPlotWidth
        plotHeight = mPlotHeight
    Else
        plotWidth = PlotCoordinateWidth()
        plotHeight = PlotCoordinateHeight()
    End If
    sx = plotWidth / (mMaxX - mMinX)
    sy = plotHeight / (mMaxY - mMinY)
    ChartScale = MinDouble(sx, sy)
End Function

' Сохраняет координатную область PlotArea для текущего прохода отрисовки.
' Excel может слегка пересчитать PlotArea после добавления Shapes и подписей,
' поэтому все Shape-элементы, оси, нейтральная линия и точка нагрузки должны
' использовать один и тот же снимок размеров.
Private Sub CapturePlotAreaMetrics()
    If mElementShapeLayerEnabled Then
        mPlotLeft = PLOT_FRAME_LEFT ' В Shape-режиме все расчетные элементы и служебные линии живут в нашей собственной рамке Chart. Это защищает схему от небольших автоматических смещений Excel PlotArea между первым и повторным построением одного и того же графика.
        mPlotTop = PLOT_FRAME_TOP
        mPlotWidth = StablePlotFrameWidth()
        mPlotHeight = StablePlotFrameHeight()
    Else
        mPlotLeft = mChart.PlotArea.InsideLeft ' При групповой отрисовке маркерами XY-диаграммы привязываем контуры и подписи к фактической области построения этих рядов.
        mPlotTop = mChart.PlotArea.InsideTop
        mPlotWidth = mChart.PlotArea.InsideWidth
        mPlotHeight = mChart.PlotArea.InsideHeight
    End If

    If mPlotWidth <= 0# Or mPlotHeight <= 0# Then
        mPlotLeft = mChart.PlotArea.Left
        mPlotTop = mChart.PlotArea.Top
        mPlotWidth = mChart.PlotArea.Width
        mPlotHeight = mChart.PlotArea.Height
    End If

    If mPlotWidth <= 0# Or mPlotHeight <= 0# Then
        mPlotLeft = PLOT_FRAME_LEFT
        mPlotTop = PLOT_FRAME_TOP
        mPlotWidth = StablePlotFrameWidth()
        mPlotHeight = StablePlotFrameHeight()
    End If

    mPlotMetricsCaptured = True
End Sub

' Гарантирует, что model-to-chart преобразование не читает PlotArea повторно
' в середине построения схемы.
Private Sub EnsurePlotAreaMetrics()
    If Not mPlotMetricsCaptured Then CapturePlotAreaMetrics
End Sub

' Возвращает ширину рабочей рамки Shape-схемы. Значение намеренно зависит
' только от ChartObject, а не от живого Excel PlotArea.
Private Function StablePlotFrameWidth() As Double
    StablePlotFrameWidth = mChartObject.Width - PLOT_FRAME_WIDTH_MARGIN
    If StablePlotFrameWidth <= 0# Then StablePlotFrameWidth = mChartObject.Width * 0.72
End Function

' Возвращает высоту рабочей рамки Shape-схемы.
Private Function StablePlotFrameHeight() As Double
    StablePlotFrameHeight = mChartObject.Height - PLOT_FRAME_HEIGHT_MARGIN
    If StablePlotFrameHeight <= 0# Then StablePlotFrameHeight = mChartObject.Height * 0.72
End Function

' Выбирает ширину области координат: заданную рамку для отдельных фигур
' или фактическую PlotArea при групповой отрисовке маркерами.
Private Function PlotCoordinateWidth() As Double
    If mElementShapeLayerEnabled Then
        PlotCoordinateWidth = StablePlotFrameWidth()
    Else
        PlotCoordinateWidth = mChart.PlotArea.InsideWidth
    End If
End Function

' Выбирает высоту координатной области с той же логикой, что и ширина.
Private Function PlotCoordinateHeight() As Double
    If mElementShapeLayerEnabled Then
        PlotCoordinateHeight = StablePlotFrameHeight()
    Else
        PlotCoordinateHeight = mChart.PlotArea.InsideHeight
    End If
End Function

' Удаляет внутри графика только Shape-объекты с префиксом этого plotter-а
' перед перерисовкой; пользовательские фигуры с другими именами сохраняет.
Private Sub ClearGeneratedShapes()
    Dim i As Long, count As Long, indexes() As Variant, name As String
    count = mChart.Shapes.Count
    If count = 0 Then Exit Sub
    ReDim indexes(1 To count): count = 0
    For i = 1 To mChart.Shapes.Count
        name = mChart.Shapes.Item(i).Name
        If Left$(name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then
            count = count + 1: indexes(count) = i
        End If
    Next i
    If count = 0 Then Exit Sub
    ReDim Preserve indexes(1 To count)
    mChart.Shapes.Range(indexes).Delete
End Sub


' Возвращает ширину бетонного элемента для схемы. Если размеры не сохранены,
' использует эквивалентный квадрат той же площади.
Private Function ConcreteElementWidth(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Double
    ConcreteElementWidth = data.Width(index)
    If ConcreteElementWidth <= 0# Then ConcreteElementWidth = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
End Function

' Возвращает высоту бетонного элемента для схемы. Если размеры не сохранены,
' использует эквивалентный квадрат той же площади.
Private Function ConcreteElementHeight(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Double
    ConcreteElementHeight = data.Height(index)
    If ConcreteElementHeight <= 0# Then ConcreteElementHeight = EquivalentLengthFromArea(data.Area(index), data.AreaUnit, data.LengthUnit)
End Function

' Возвращает угол элемента для схемы. Для известного прямоугольника берёт
' его собственный Rotation, для квадрата только по площади —
' средний угол импортированной сетки.
Private Function ConcreteElementVisualRotation(ByVal data As CSectionPlotDataReader, ByVal index As Long, _
        ByVal fallbackSquareRotation As Double) As Double
    ConcreteElementVisualRotation = data.Rotation(index)
    If IsEquivalentAreaFallbackElement(data, index) And Abs(ConcreteElementVisualRotation) <= 0.000000000001 Then
        ConcreteElementVisualRotation = fallbackSquareRotation
    End If
End Function

' Находит среднюю ориентацию бетонных элементов с известными Width/Height.
' Угол прямоугольника усредняется через 2*angle, потому что поворот на 180
' градусов описывает ту же самую ячейку.
Private Function AverageKnownConcreteElementRotation(ByVal data As CSectionPlotDataReader) As Double
    Dim sumCos As Double
    Dim sumSin As Double
    Dim count As Long
    Dim i As Long
    For i = 1 To data.Count
        If StrComp(data.MaterialType(i), "Concrete", vbTextCompare) = 0 And _
                Not IsEquivalentAreaFallbackElement(data, i) And _
                data.Width(i) > 0# And data.Height(i) > 0# Then
            sumCos = sumCos + Cos(2# * data.Rotation(i))
            sumSin = sumSin + Sin(2# * data.Rotation(i))
            count = count + 1
        End If
    Next i

    If count > 0 And (Abs(sumCos) > 0.000000000001 Or Abs(sumSin) > 0.000000000001) Then
        AverageKnownConcreteElementRotation = 0.5 * Atn2Local(sumSin, sumCos)
    End If
End Function

' Проверяет, представлен ли элемент эквивалентным квадратом только по площади.
' Принимает записи Equivalent square и EquivalentSquare;
' сохранённые положительные Width/Height означают, что размеры уже известны.
Private Function IsEquivalentAreaFallbackElement(ByVal data As CSectionPlotDataReader, ByVal index As Long) As Boolean
    Dim statusText As String
    statusText = Replace$(Trim$(data.ShapeType(index)), " ", vbNullString)
    IsEquivalentAreaFallbackElement = (StrComp(statusText, "EquivalentSquare", vbTextCompare) = 0) And _
        (data.Width(index) <= 0# Or data.Height(index) <= 0#)
End Function

' Возвращает угол atan2(y, x) в радианах для ориентации элементов
' и граничных дуг, включая все квадранты и осевые направления.
Private Function Atn2Local(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= 0.000000000001 Then
        If y > 0# Then
            Atn2Local = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2Local = -GEOM_PI / 2#
        Else
            Atn2Local = 0#
        End If
    Else
        Atn2Local = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then
                Atn2Local = Atn2Local + GEOM_PI
            Else
                Atn2Local = Atn2Local - GEOM_PI
            End If
        End If
    End If
End Function

' Возвращает размер маркера для групповой отрисовки элементов рядами диаграммы.
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

' Восстанавливает сторону эквивалентного квадрата по площади элемента.
Private Function EquivalentLengthFromArea(ByVal areaValue As Double, ByVal areaUnit As String, ByVal lengthUnit As String) As Double
    If areaValue <= 0# Then
        EquivalentLengthFromArea = 0#
    Else
        EquivalentLengthFromArea = Sqr(mUnits.OutputAreaToInternal(areaValue, areaUnit)) / mUnits.OutputLengthToInternal(1#, lengthUnit)
    End If
End Function

' Возвращает русскую приставку заголовка схемы без зависимости от кодировки
' исходного VBA-файла.
Private Function PlotTitlePrefix() As String
    PlotTitlePrefix = ChrW$(&H421) & ChrW$(&H445) & ChrW$(&H435) & ChrW$(&H43C) & ChrW$(&H430) & " " & _
        ChrW$(&H441) & ChrW$(&H435) & ChrW$(&H447) & ChrW$(&H435) & ChrW$(&H43D) & ChrW$(&H438) & ChrW$(&H44F) & " " & _
        ChrW$(&H43F) & ChrW$(&H440) & ChrW$(&H438) & " " & ChrW$(&H437) & ChrW$(&H430) & ChrW$(&H433) & _
        ChrW$(&H440) & ChrW$(&H443) & ChrW$(&H436) & ChrW$(&H435) & ChrW$(&H43D) & ChrW$(&H438) & ChrW$(&H438) & " "
End Function

' Формирует заголовок схемы для выбранного LC или режима просмотра геометрии.
Private Function PlotTitleText(ByVal data As CSectionPlotDataReader) As String
    If data.GeometryOnly Then
        If Len(Trim$(data.GeometryOnlyTitle)) > 0 Then
            PlotTitleText = data.GeometryOnlyTitle
        Else
            PlotTitleText = "Геометрия без НДС"
        End If
        Exit Function
    End If
    PlotTitleText = PlotTitlePrefix() & data.LoadCase
    If Len(Trim$(data.LoadCaseComment)) > 0 Then PlotTitleText = PlotTitleText & " (" & data.LoadCaseComment & ")"
End Function

' Форматирует значение Stress/Strain с профильной точностью выбранной величины.
Private Function FormatPlotValue(ByVal value As Double, ByVal precision As Long) As String
    If precision < 0 Then precision = 0
    If precision > 10 Then precision = 10

    Dim pattern As String
    If precision <= 0 Then
        pattern = "0"
    Else
        pattern = "0." & String$(precision, "0")
    End If
    FormatPlotValue = Replace$(Format$(value, pattern), ",", ".")
End Function

' Округляет неотрицательный размер маркера или номер цветовой ступени.
' Int до CLng сохраняет целые нечетные уровни и округляет ровно половину вверх.
Private Function RoundToLong(ByVal value As Double) As Long
    RoundToLong = CLng(Int(value + 0.5))
End Function

' Ограничивает значение заданным диапазоном.
Private Function ClampDouble(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    ClampDouble = value
    If ClampDouble < minValue Then ClampDouble = minValue
    If ClampDouble > maxValue Then ClampDouble = maxValue
End Function

' Утилитарный максимум без WorksheetFunction в коде схемы.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Утилитарный минимум без WorksheetFunction в коде схемы.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Рисует рабочую аналитическую область в фиксированной рамке 200 на 200 пунктов.
' Переданный Chart принадлежит тесту; Config, Results и расчет не затрагиваются.
Public Function TestDrawConcreteRegionShape(ByVal chart As Object, ByVal region As CConcreteRegion, _
        ByVal centerX As Double, ByVal centerY As Double) As Object
    Set mChart = chart
    Dim query As CSectionGeometryQuery: Set query = New CSectionGeometryQuery
    query.ProjectionBounds region, 1#, 0#, mMinX, mMaxX
    query.ProjectionBounds region, 0#, 1#, mMinY, mMaxY
    mPlotLeft = 0#: mPlotTop = 0#: mPlotWidth = 200#: mPlotHeight = 200#: mPlotMetricsCaptured = True
    Dim shape As Object: Set shape = DrawConcreteRegionShape(region, centerX, centerY)
    shape.Fill.ForeColor.RGB = RGB(210, 20, 40): shape.Line.Visible = False
    Set TestDrawConcreteRegionShape = shape
End Function
