VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CPlotAnnotationLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Компоновка аннотаций схемы
' ==========================================================================
' Класс получает уже готовый semantic-набор аннотаций: размерные линии и
' групповые подписи арматуры. Он не решает, какие размеры нужны конкретному
' сечению, и не определяет состав арматуры по координатам.
'
' Его задача ниже по pipeline: перевести модельные координаты сечения в
' координаты Chart, применить общие визуальные отступы, масштаб текста,
' стрелки. Выбор зоны/арматуры и расчетные данные остаются у верхних слоев.

Private Const ITEM_LINE As String = "LINE" ' Готовый layout-элемент линии для Chart.Shapes.
Private Const ITEM_TEXT As String = "TEXT" ' Готовый layout-элемент текстовой подписи для Chart.Shapes.

Private mMinX As Double ' Левая граница в единицах модели, переданных в Initialize; reader передает мм.
Private mMaxX As Double ' Правая граница в тех же единицах модели.
Private mMinY As Double ' Нижняя граница в тех же единицах модели.
Private mMaxY As Double ' Верхняя граница в тех же единицах модели.
Private mPlotLeft As Double ' Левый край внутренней PlotArea внутри Chart.
Private mPlotTop As Double ' Верхний край внутренней PlotArea внутри Chart.
Private mPlotWidth As Double ' Фактическая ширина PlotArea, по которой строится transform.
Private mPlotHeight As Double ' Фактическая высота PlotArea, по которой строится transform.
Private mScaleX As Double ' Количество chart-points на одну единицу модели по X.
Private mScaleY As Double ' Количество chart-points на одну единицу модели по Y.
Private mLengthUnitFactorToMm As Double ' Пересчет единицы переданной модели в мм через CUnitSystem.
Private mVisualScale As Double ' Общий масштаб оформления относительно базового размера Chart.
Private mDimensionFontSize As Double ' Итоговый размер текста размеров в points.
Private mRebarFontSize As Double ' Итоговый размер текста групповых подписей арматуры в points.
Private mDimensionBaseGap As Double ' Отступ размерной линии от своей anchor-линии в points.
Private mRebarBaseGap As Double ' Отступ подписи арматуры от линии осей арматуры в points.
Private mDimensionLineWeight As Double ' Толщина основной размерной линии.
Private mExtensionLineWeight As Double ' Толщина выносных линий размера.
Private mRebarLineWeight As Double ' Толщина линии-выноски под групповой подписью арматуры.
Private mDimensionTextHeightBase As Double ' Высота текста размеров в единицах mDimensionTextUnits.
Private mRebarTextHeightBase As Double ' Высота текста арматуры в единицах mRebarTextUnits.
Private mDimensionTextUnits As String ' Единицы TextHeight/TextGap для размеров: mm или pt.
Private mRebarTextUnits As String ' Единицы TextHeight/TextGap для подписей арматуры: mm или pt.
Private mDimensionBaseGapBase As Double ' Базовый отступ размеров в модельных мм.
Private mRebarBaseGapBase As Double ' Базовый отступ подписей арматуры в модельных мм.
Private mDimensionLineWeightBase As Double ' Базовая толщина размерной линии до visual-scale.
Private mExtensionLineWeightBase As Double ' Базовая толщина выносной линии до visual-scale.
Private mRebarLineWeightBase As Double ' Базовая толщина линии арматурной подписи до visual-scale.
Private mDimensionTextGapBase As Double ' Зазор размер-текст в единицах mDimensionTextUnits.
Private mRebarTextGapBase As Double ' Зазор линия-текст в единицах mRebarTextUnits.
Private mDimensionPlacement As String ' С какой стороны от размерной линии ставить текст: Outside или Inside.
Private mRebarPlacement As String ' Куда смещать подпись арматуры от линии стержней: Outside или Inside.
Private mRebarLineEnabled As Boolean ' Рисовать ли короткую линию под групповой подписью арматуры.
Private mDimensionColor As Long ' Цвет размерных линий и текста.
Private mDimensionExtensionColor As Long ' Цвет выносных линий размеров.
Private mRebarColor As Long ' Цвет подписей арматуры и их линий.
Private mDimensionsEnabled As Boolean ' Активна ли группа размеров; выключенные параметры не читаются.
Private mRebarLabelsEnabled As Boolean ' Активна ли группа подписей арматуры.

Private mCount As Long ' Число готовых layout-элементов, которые затем создаст CSectionPlotter.
Private mKind() As String ' Тип layout-элемента: линия или текст.
Private mX1() As Double ' Для линии: X начала; для текста: left.
Private mY1() As Double ' Для линии: Y начала; для текста: top.
Private mX2() As Double ' Для линии: X конца; для текста: width.
Private mY2() As Double ' Для линии: Y конца; для текста: height.
Private mText() As String ' Текст подписи; для линий пусто.
Private mRotation() As Double ' Поворот текстового блока в градусах.
Private mColor() As Long ' Цвет элемента Chart.Shape.
Private mWeight() As Double ' Толщина линии; для текста не используется.
Private mBold() As Boolean ' Нужно ли делать текст полужирным.
Private mArrowheads() As Boolean ' Нужно ли добавить стрелки с обеих сторон линии.
Private mFontSize() As Double ' Размер шрифта конкретной текстовой подписи.

' ------------------------------
' Настройки оформления
' ------------------------------
' Задает базовое оформление для самостоятельного layout без Config.
' При ConfigureFromSettings эти значения не подменяют обязательный ввод;
' настоящий черный RGB=0 сохраняется и не служит признаком отсутствия цвета.
Private Sub Class_Initialize()
    mDimensionsEnabled = True: mRebarLabelsEnabled = True
    mDimensionTextHeightBase = 13#: mRebarTextHeightBase = 13#
    mDimensionTextUnits = "mm": mRebarTextUnits = "mm"
    mDimensionBaseGapBase = 100#: mRebarBaseGapBase = 60#
    mDimensionLineWeightBase = 1.75: mExtensionLineWeightBase = 0.85: mRebarLineWeightBase = 1.35
    mDimensionTextGapBase = 9#: mRebarTextGapBase = 9#
    mDimensionPlacement = "Outside": mRebarPlacement = "Outside": mRebarLineEnabled = True
    mDimensionColor = RGB(20, 30, 90): mDimensionExtensionColor = RGB(140, 140, 140): mRebarColor = RGB(20, 30, 90)
End Sub

' Читает только визуальные параметры схемы. Смысл аннотаций уже должен быть
' сформирован выше. Обязательные активные значения проверяются без defaults;
' выключенная группа не читает свои параметры, в том числе ошибочные формулы.
Public Sub ConfigureFromSettings(ByVal settings As CSystemSettingsReader)
    If settings Is Nothing Then Err.Raise vbObjectError + 4810, "CPlotAnnotationLayout", "Настройки оформления схемы не переданы."
    mDimensionsEnabled = settings.GetRequiredBoolean("Plot.Dimensions.Enabled")
    mRebarLabelsEnabled = settings.GetRequiredBoolean("Plot.RebarLabels.Enabled")
    If mDimensionsEnabled Then
        mDimensionTextHeightBase = AnnotationNumber(settings, "Plot.Dimensions.TextHeight", False)
        mDimensionTextUnits = settings.GetRequiredChoice("Plot.Dimensions.TextUnits", Array("mm", "pt"))
        mDimensionBaseGapBase = AnnotationNumber(settings, "Plot.Dimensions.Offset", True)
        mDimensionLineWeightBase = AnnotationNumber(settings, "Plot.Dimensions.LineWeight", False)
        mExtensionLineWeightBase = AnnotationNumber(settings, "Plot.Dimensions.ExtensionLineWeight", False)
        mDimensionTextGapBase = AnnotationNumber(settings, "Plot.Dimensions.TextGap", True)
        mDimensionPlacement = settings.GetRequiredChoice("Plot.Dimensions.Placement", Array("Outside", "Inside"))
        mDimensionColor = settings.GetRequiredRgbColor("Plot.Dimensions.Color")
        mDimensionExtensionColor = settings.GetRequiredRgbColor("Plot.Dimensions.ExtensionLineColor")
    End If
    If mRebarLabelsEnabled Then
        mRebarTextHeightBase = AnnotationNumber(settings, "Plot.RebarLabels.TextHeight", False)
        mRebarTextUnits = settings.GetRequiredChoice("Plot.RebarLabels.TextUnits", Array("mm", "pt"))
        mRebarBaseGapBase = AnnotationNumber(settings, "Plot.RebarLabels.Offset", True)
        mRebarLineEnabled = settings.GetRequiredBoolean("Plot.RebarLabels.LineEnabled")
        If mRebarLineEnabled Then mRebarLineWeightBase = AnnotationNumber(settings, "Plot.RebarLabels.LineWeight", False)
        mRebarTextGapBase = AnnotationNumber(settings, "Plot.RebarLabels.TextGap", True)
        mRebarPlacement = settings.GetRequiredChoice("Plot.RebarLabels.Placement", Array("Outside", "Inside"))
        mRebarColor = settings.GetRequiredRgbColor("Plot.RebarLabels.Color")
    End If
End Sub

' Проверяет знак обязательной визуальной величины до расчетов Chart.
' Нулевой Offset/TextGap допустим; высота текста и толщина должны быть >0.
Private Function AnnotationNumber(ByVal settings As CSystemSettingsReader, ByVal key As String, ByVal allowZero As Boolean) As Double
    Dim value As Double
    value = settings.GetRequiredDouble(key)
    If value < 0# Or (value = 0# And Not allowZero) Then
        Dim correction As String
        If allowZero Then correction = "Введите число не меньше нуля." Else correction = "Введите положительное число."
        Err.Raise vbObjectError + 4811, "CPlotAnnotationLayout", settings.InputErrorMessage(key, "Недопустимый размер оформления.", correction)
    End If
    AnnotationNumber = value
End Function

' ------------------------------
' Геометрия Chart
' ------------------------------
' Фиксирует границы модели и фактическую PlotArea. После этого offset всегда
' переводится из модельных мм, а TextGap/TextHeight - из выбранных единиц
' mm или pt в реальные points Chart.
Public Sub Initialize(ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, _
        ByVal plotLeft As Double, ByVal plotTop As Double, ByVal plotWidth As Double, ByVal plotHeight As Double, _
        Optional ByVal lengthUnit As String = "mm")
    mMinX = minX
    mMaxX = maxX
    mMinY = minY
    mMaxY = maxY
    mPlotLeft = plotLeft
    mPlotTop = plotTop
    mPlotWidth = plotWidth
    mPlotHeight = plotHeight
    mCount = 0
    If mMaxX = mMinX Then mMaxX = mMinX + 1#
    If mMaxY = mMinY Then mMaxY = mMinY + 1#
    mScaleX = mPlotWidth / (mMaxX - mMinX)
    mScaleY = mPlotHeight / (mMaxY - mMinY)
    mLengthUnitFactorToMm = LengthFactorToMm(lengthUnit)
    InitializeVisualMetrics plotWidth, plotHeight
End Sub

' Рассчитывает итоговые размеры оформления. Текст и зазоры масштабируются от
' реального размера Chart, а не от текущих размеров сечения.
Private Sub InitializeVisualMetrics(ByVal plotWidth As Double, ByVal plotHeight As Double)
    Dim sx As Double, sy As Double
    sx = plotWidth / 640#
    sy = plotHeight / 430#
    mVisualScale = ClampDouble(MinDouble(sx, sy), 0.65, 1.8)

    mDimensionFontSize = ClampDouble(TextMetricToChartPoints(mDimensionTextHeightBase, mDimensionTextUnits), 5#, 28#)
    mRebarFontSize = ClampDouble(TextMetricToChartPoints(mRebarTextHeightBase, mRebarTextUnits), 5#, 28#)
    mDimensionBaseGap = ClampDouble(ModelMmToChartPoints(mDimensionBaseGapBase), 4#, 180#)
    mRebarBaseGap = ClampDouble(ModelMmToChartPoints(mRebarBaseGapBase), 2#, 120#)
    mDimensionLineWeight = mDimensionLineWeightBase * mVisualScale
    mExtensionLineWeight = mExtensionLineWeightBase * mVisualScale
    mRebarLineWeight = mRebarLineWeightBase * mVisualScale
End Sub

' ------------------------------
' Размерная аннотация
' ------------------------------
' На вход приходят две точки размера и наружная нормаль в координатах модели.
' Layout сам ставит выносные линии, основную линию со стрелками и текст.
Public Sub AddDimension(ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal textValue As String)
    If Not mDimensionsEnabled Then Exit Sub
    Dim dirX As Double, dirY As Double
    dirX = endX - startX
    dirY = endY - startY
    NormalizeModelVector dirX, dirY, 1#, 0#
    NormalizeModelVector normalX, normalY, -dirY, dirX

    Dim normalPxX As Double, normalPxY As Double
    PixelUnitVector normalX, normalY, normalPxX, normalPxY

    Dim baseOffset As Double
    baseOffset = mDimensionBaseGap

    Dim sx As Double, sy As Double, ex As Double, ey As Double
    sx = ModelToChartX(startX) + normalPxX * baseOffset
    sy = ModelToChartY(startY) + normalPxY * baseOffset
    ex = ModelToChartX(endX) + normalPxX * baseOffset
    ey = ModelToChartY(endY) + normalPxY * baseOffset

    Dim rotation As Double
    rotation = AnnotationRotation(dirX, dirY)

    Dim textNormalPxX As Double, textNormalPxY As Double
    textNormalPxX = normalPxX
    textNormalPxY = normalPxY
    ApplyPlacement textNormalPxX, textNormalPxY, mDimensionPlacement

    Dim textLeft As Double, textTop As Double, textWidth As Double, textHeight As Double
    DimensionTextBox sx, sy, ex, ey, textNormalPxX, textNormalPxY, rotation, textLeft, textTop, textWidth, textHeight

    AddLine ModelToChartX(startX), ModelToChartY(startY), sx, sy, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine ModelToChartX(endX), ModelToChartY(endY), ex, ey, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine sx, sy, ex, ey, mDimensionColor, mDimensionLineWeight, True
    AddText textLeft, textTop, textWidth, textHeight, textValue, rotation, mDimensionColor, True, mDimensionFontSize
End Sub

' ------------------------------
' Групповая подпись арматуры
' ------------------------------
' AnchorStart/AnchorEnd описывают реальную линию раскладки стержней. Поэтому
' подпись центрируется относительно группы арматуры, а не относительно Bounds.
Public Sub AddRebarLabel(ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal textValue As String)
    If Not mRebarLabelsEnabled Then Exit Sub
    Dim dirX As Double, dirY As Double
    dirX = endX - startX
    dirY = endY - startY
    NormalizeModelVector dirX, dirY, 1#, 0#
    NormalizeModelVector normalX, normalY, -dirY, dirX

    Dim normalPxX As Double, normalPxY As Double
    PixelUnitVector normalX, normalY, normalPxX, normalPxY
    ApplyPlacement normalPxX, normalPxY, mRebarPlacement

    Dim dirPxX As Double, dirPxY As Double
    PixelUnitVector dirX, dirY, dirPxX, dirPxY

    Dim anchorX As Double, anchorY As Double
    anchorX = 0.5 * (startX + endX)
    anchorY = 0.5 * (startY + endY)

    Dim labelOffset As Double
    labelOffset = mRebarBaseGap

    Dim anchorPxX As Double, anchorPxY As Double
    anchorPxX = ModelToChartX(anchorX) + normalPxX * labelOffset
    anchorPxY = ModelToChartY(anchorY) + normalPxY * labelOffset

    Dim rotation As Double
    rotation = AnnotationRotation(dirX, dirY)

    Dim textLeft As Double, textTop As Double, textWidth As Double, textHeight As Double
    RebarTextBox anchorPxX, anchorPxY, normalPxX, normalPxY, rotation, textLeft, textTop, textWidth, textHeight

    Dim halfLength As Double
    halfLength = MaxDouble(26# * mVisualScale, 0.22 * Sqr((ModelToChartX(endX) - ModelToChartX(startX)) * _
        (ModelToChartX(endX) - ModelToChartX(startX)) + (ModelToChartY(endY) - ModelToChartY(startY)) * _
        (ModelToChartY(endY) - ModelToChartY(startY))))
    If mRebarLineEnabled Then
        AddLine anchorPxX - dirPxX * halfLength, anchorPxY - dirPxY * halfLength, _
            anchorPxX + dirPxX * halfLength, anchorPxY + dirPxY * halfLength, mRebarColor, mRebarLineWeight, False
    End If
    AddText textLeft, textTop, textWidth, textHeight, textValue, rotation, mRebarColor, True, mRebarFontSize
End Sub

' Определяет bbox текста размера вокруг середины размерной линии. TextGap
' задает отступ центра подписи от размерной линии по выбранной нормали. Для
' вертикальных размеров bbox сразу узкий и высокий: Excel затем создает
' вертикальный текстовый Shape без поворота широкого горизонтального блока.
Private Sub DimensionTextBox(ByVal sx As Double, ByVal sy As Double, ByVal ex As Double, ByVal ey As Double, _
        ByVal normalPxX As Double, ByVal normalPxY As Double, ByVal rotation As Double, _
        ByRef leftPos As Double, ByRef topPos As Double, ByRef widthValue As Double, ByRef heightValue As Double)
    If rotation = 0# Then
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mDimensionFontSize * 1.65)
    Else
        widthValue = MaxDouble(18# * mVisualScale, mDimensionFontSize * 1.65)
        heightValue = 150# * mVisualScale
    End If

    Dim textGap As Double
    textGap = ClampDouble(TextMetricToChartPoints(mDimensionTextGapBase, mDimensionTextUnits), 0#, 80#)

    Dim midX As Double, midY As Double
    midX = 0.5 * (sx + ex) + normalPxX * textGap
    midY = 0.5 * (sy + ey) + normalPxY * textGap

    leftPos = midX - widthValue / 2#
    topPos = midY - heightValue / 2#
End Sub

' Определяет bbox групповой подписи арматуры. Положение считается от середины
' anchor-линии, а не от крайних фактических стержней. Вертикальные подписи
' получают узкий высокий bbox по той же причине, что и размерные подписи.
Private Sub RebarTextBox(ByVal anchorPxX As Double, ByVal anchorPxY As Double, ByVal normalPxX As Double, _
        ByVal normalPxY As Double, ByVal rotation As Double, ByRef leftPos As Double, ByRef topPos As Double, _
        ByRef widthValue As Double, ByRef heightValue As Double)
    If rotation = 0# Then
        widthValue = 150# * mVisualScale
        heightValue = MaxDouble(18# * mVisualScale, mRebarFontSize * 1.65)
    Else
        widthValue = MaxDouble(18# * mVisualScale, mRebarFontSize * 1.65)
        heightValue = 150# * mVisualScale
    End If

    Dim textGap As Double
    textGap = ClampDouble(TextMetricToChartPoints(mRebarTextGapBase, mRebarTextUnits), 0#, 80#)

    Dim centerX As Double, centerY As Double
    centerX = anchorPxX + normalPxX * textGap
    centerY = anchorPxY + normalPxY * textGap

    leftPos = centerX - widthValue / 2#
    topPos = centerY - heightValue / 2#
End Sub

' Для текущей схемы достаточно горизонтального текста и поворота на 90 градусов.
' Наклонные грани сохраняют направление anchor-линии, но текст не вращается на
' произвольный угол, чтобы он оставался читаемым в Excel.
Private Function AnnotationRotation(ByVal dirX As Double, ByVal dirY As Double) As Double
    If Abs(dirY) > Abs(dirX) Then AnnotationRotation = 90# Else AnnotationRotation = 0#
End Function

' При Inside разворачивает наружную нормаль внутрь сечения. Так одна и та же
' semantic-аннотация может рисоваться по обе стороны своей anchor-линии.
Private Sub ApplyPlacement(ByRef normalPxX As Double, ByRef normalPxY As Double, ByVal placement As String)
    If StrComp(placement, "Inside", vbTextCompare) = 0 Then
        normalPxX = -normalPxX
        normalPxY = -normalPxY
    End If
End Sub

' Нормализует модельный вектор. Если annotation пришла с нулевой нормалью или
' нулевым направлением, используется fallback, чтобы layout не делил на ноль.
Private Sub NormalizeModelVector(ByRef x As Double, ByRef y As Double, ByVal fallbackX As Double, ByVal fallbackY As Double)
    Dim lengthValue As Double
    lengthValue = Sqr(x * x + y * y)
    If lengthValue <= 0# Then
        x = fallbackX
        y = fallbackY
        lengthValue = Sqr(x * x + y * y)
    End If
    If lengthValue <= 0# Then
        x = 1#
        y = 0#
    Else
        x = x / lengthValue
        y = y / lengthValue
    End If
End Sub

' Переводит направление из координат модели в единичный вектор Chart. По Y
' знак меняется, потому что в Chart координата вниз увеличивается.
Private Sub PixelUnitVector(ByVal modelX As Double, ByVal modelY As Double, ByRef pixelX As Double, ByRef pixelY As Double)
    pixelX = modelX * mScaleX
    pixelY = -modelY * mScaleY
    Dim lengthValue As Double
    lengthValue = Sqr(pixelX * pixelX + pixelY * pixelY)
    If lengthValue <= 0# Then
        pixelX = 1#
        pixelY = 0#
    Else
        pixelX = pixelX / lengthValue
        pixelY = pixelY / lengthValue
    End If
End Sub

' X модели -> X внутри Chart.PlotArea.
Private Function ModelToChartX(ByVal xCoord As Double) As Double
    ModelToChartX = mPlotLeft + (xCoord - mMinX) / (mMaxX - mMinX) * mPlotWidth
End Function

' Y модели -> Y внутри Chart.PlotArea с переворотом вертикальной оси Excel.
Private Function ModelToChartY(ByVal yCoord As Double) As Double
    ModelToChartY = mPlotTop + mPlotHeight - (yCoord - mMinY) / (mMaxY - mMinY) * mPlotHeight
End Function

' Берем меньший масштаб из X/Y, чтобы модельные миллиметры оформления не
' растягивались при прямоугольной области построения.
Private Function ModelScale() As Double
    ModelScale = MinDouble(mScaleX, mScaleY)
End Function

' Переводит модельную миллиметровую величину оформления в экранные points
' текущей PlotArea. TextUnits=pt идет через TextMetricToChartPoints ниже.
Private Function ModelMmToChartPoints(ByVal valueMm As Double) As Double
    If mLengthUnitFactorToMm <= 0# Then mLengthUnitFactorToMm = 1#
    ModelMmToChartPoints = valueMm / mLengthUnitFactorToMm * ModelScale()
End Function

' Переводит TextHeight/TextGap в points Chart. В режиме pt значение уже
' является экранной величиной Excel, поэтому не зависит от размера сечения.
Private Function TextMetricToChartPoints(ByVal value As Double, ByVal unitName As String) As Double
    If StrComp(unitName, "pt", vbTextCompare) = 0 Then
        TextMetricToChartPoints = value
    Else
        TextMetricToChartPoints = ModelMmToChartPoints(value)
    End If
End Function

' Показывает, сколько мм содержится в одной output-единице длины Results.
' Это нужно, если Results был записан не в мм, а например в метрах.
Private Function LengthFactorToMm(ByVal unitName As String) As Double
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    LengthFactorToMm = units.OutputLengthToInternal(1#, unitName)
End Function

' Добавляет линию в список layout-элементов; фактический Shape создаст plotter.
Private Sub AddLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal hasArrowheads As Boolean)
    AppendItem ITEM_LINE, x1, y1, x2, y2, vbNullString, 0#, colorValue, weight, False, hasArrowheads
End Sub

' Добавляет текстовую подпись в список layout-элементов; фактический Shape
' создаст plotter внутри нужного Chart.
Private Sub AddText(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal textValue As String, ByVal rotation As Double, _
        ByVal colorValue As Long, ByVal isBold As Boolean, ByVal fontSize As Double)
    AppendItem ITEM_TEXT, leftPos, topPos, widthValue, heightValue, textValue, rotation, colorValue, 0#, isBold, False, fontSize
End Sub

' Складывает готовый layout-элемент в параллельные массивы. Это простой VBA-
' формат без классов на каждый Shape, чтобы обновление схемы оставалось легким.
Private Sub AppendItem(ByVal kind As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal textValue As String, ByVal rotation As Double, ByVal colorValue As Long, ByVal weight As Double, _
        ByVal isBold As Boolean, ByVal hasArrowheads As Boolean, Optional ByVal fontSize As Double = 0#)
    mCount = mCount + 1
    ReDim Preserve mKind(1 To mCount)
    ReDim Preserve mX1(1 To mCount)
    ReDim Preserve mY1(1 To mCount)
    ReDim Preserve mX2(1 To mCount)
    ReDim Preserve mY2(1 To mCount)
    ReDim Preserve mText(1 To mCount)
    ReDim Preserve mRotation(1 To mCount)
    ReDim Preserve mColor(1 To mCount)
    ReDim Preserve mWeight(1 To mCount)
    ReDim Preserve mBold(1 To mCount)
    ReDim Preserve mArrowheads(1 To mCount)
    ReDim Preserve mFontSize(1 To mCount)

    mKind(mCount) = kind
    mX1(mCount) = x1
    mY1(mCount) = y1
    mX2(mCount) = x2
    mY2(mCount) = y2
    mText(mCount) = textValue
    mRotation(mCount) = rotation
    mColor(mCount) = colorValue
    mWeight(mCount) = weight
    mBold(mCount) = isBold
    mArrowheads(mCount) = hasArrowheads
    mFontSize(mCount) = fontSize
End Sub

' Ограничивает визуальные величины диапазоном, в котором они остаются читаемыми.
Private Function ClampDouble(ByVal value As Double, ByVal minValue As Double, ByVal maxValue As Double) As Double
    If value < minValue Then
        ClampDouble = minValue
    ElseIf value > maxValue Then
        ClampDouble = maxValue
    Else
        ClampDouble = value
    End If
End Function

' Небольшая локальная функция, чтобы не тянуть WorksheetFunction в горячий слой.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Небольшая локальная функция, чтобы не тянуть WorksheetFunction в горячий слой.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' ------------------------------
' Готовые элементы для Plotter
' ------------------------------
' Ниже только чтение уже рассчитанного layout. Plotter проходит по этим
' массивам и создает Chart.Shapes с нужными координатами, цветами и стрелками.
Public Property Get Count() As Long: Count = mCount: End Property
Public Property Get ItemKind(ByVal index As Long) As String: ItemKind = mKind(index): End Property
Public Property Get X1(ByVal index As Long) As Double: X1 = mX1(index): End Property
Public Property Get Y1(ByVal index As Long) As Double: Y1 = mY1(index): End Property
Public Property Get X2(ByVal index As Long) As Double: X2 = mX2(index): End Property
Public Property Get Y2(ByVal index As Long) As Double: Y2 = mY2(index): End Property
Public Property Get Text(ByVal index As Long) As String: Text = mText(index): End Property
Public Property Get Rotation(ByVal index As Long) As Double: Rotation = mRotation(index): End Property
Public Property Get Color(ByVal index As Long) As Long: Color = mColor(index): End Property
Public Property Get Weight(ByVal index As Long) As Double: Weight = mWeight(index): End Property
Public Property Get Bold(ByVal index As Long) As Boolean: Bold = mBold(index): End Property
Public Property Get Arrowheads(ByVal index As Long) As Boolean: Arrowheads = mArrowheads(index): End Property
Public Property Get FontSize(ByVal index As Long) As Double: FontSize = mFontSize(index): End Property
Public Property Get VisualScale() As Double: VisualScale = mVisualScale: End Property
Public Property Get DimensionsEnabled() As Boolean: DimensionsEnabled = mDimensionsEnabled: End Property
Public Property Get RebarLabelsEnabled() As Boolean: RebarLabelsEnabled = mRebarLabelsEnabled: End Property
