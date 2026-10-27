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
' Размещение размеров и подписей арматуры на схеме
' ==========================================================================
' Получает готовые аннотации в координатах сечения, переводит их
' в координаты диаграммы Excel и применяет отступы, размеры текста и стрелок.
' Состав арматуры и перечень нужных размеров определены до этого этапа;
' данный класс отвечает за их размещение и оформление. Координаты внутри Chart,
' размеры шрифта и толщины линий заданы в пунктах Excel (pt), а не в пикселях.

Private Const ITEM_LINE As String = "LINE" ' Обозначение графического элемента «линия» для последующего создания Chart.Shape.
Private Const ITEM_TEXT As String = "TEXT" ' Обозначение графического элемента «текст» для последующего создания Chart.Shape.

Private mMinX As Double ' Левая граница области в переданных единицах модели; CSectionPlotDataReader использует мм.
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
Private mOffsetScale As Double ' Общий коэффициент отступов, переданный плоттером по габариту сечения.
Private mDimensionFontSize As Double ' Итоговый размер текста размеров в points.
Private mRebarFontSize As Double ' Итоговый размер текста групповых подписей арматуры в points.
Private mDimensionBaseGap As Double ' Отступ размерной линии от своей anchor-линии в points.
Private mRebarBaseGap As Double ' Отступ подписи арматуры от бетонной границы в points.
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

Private mCapacity As Long ' Закрытый резерв графических элементов; Count ограничивает доступную заполненную часть.
Private mCount As Long ' Число подготовленных линий и текстов, которые затем создаст CSectionPlotter.
Private mKind() As String ' Вид подготовленного графического элемента: линия или текст.
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
' Задаёт начальное оформление для отдельного программного вызова без Config.
' При ConfigureFromSettings эти значения не заменяют обязательный ввод.
' Чёрный цвет RGB = 0 является полноценным цветом, а не отсутствующим значением.
Private Sub Class_Initialize()
    mDimensionsEnabled = True: mRebarLabelsEnabled = True
    mDimensionTextHeightBase = 13#: mRebarTextHeightBase = 13#
    mDimensionTextUnits = "mm": mRebarTextUnits = "mm"
    mDimensionBaseGapBase = 100#: mRebarBaseGapBase = 30#
    mDimensionLineWeightBase = 1.75: mExtensionLineWeightBase = 0.85: mRebarLineWeightBase = 1.35
    mDimensionTextGapBase = 9#: mRebarTextGapBase = 9#
    mDimensionPlacement = "Outside": mRebarPlacement = "Outside": mRebarLineEnabled = True
    mDimensionColor = RGB(20, 30, 90): mDimensionExtensionColor = RGB(140, 140, 140): mRebarColor = RGB(20, 30, 90)
End Sub

' Читает только визуальные параметры схемы. Смысл аннотаций уже должен быть
' сформирован выше. Обязательные активные значения проверяются без подстановок;
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
        Optional ByVal lengthUnit As String = "mm", Optional ByVal offsetScale As Double = 1#)
    If offsetScale < 1# Then Err.Raise vbObjectError + 4811, "CPlotAnnotationLayout", "Коэффициент отступов должен быть не меньше 1."
    mOffsetScale = offsetScale
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
    mDimensionBaseGap = ClampDouble(ModelMmToChartPoints(mDimensionBaseGapBase * mOffsetScale), 4#, 180#)
    mRebarBaseGap = ClampDouble(ModelMmToChartPoints(mRebarBaseGapBase * mOffsetScale), 2#, 120#)
    mDimensionLineWeight = mDimensionLineWeightBase * mVisualScale
    mExtensionLineWeight = mExtensionLineWeightBase * mVisualScale
    mRebarLineWeight = mRebarLineWeightBase * mVisualScale
End Sub

' ------------------------------
' Размещение размера
' ------------------------------
' По двум точкам размера и направлению выноски в координатах сечения
' вычисляет положение выносных линий, размерной линии, стрелок и текста.
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
    DimensionTextBox sx, sy, ex, ey, textNormalPxX, textNormalPxY, rotation, textValue, textLeft, textTop, textWidth, textHeight

    AddLine ModelToChartX(startX), ModelToChartY(startY), sx, sy, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine ModelToChartX(endX), ModelToChartY(endY), ex, ey, mDimensionExtensionColor, mExtensionLineWeight, False
    AddLine sx, sy, ex, ey, mDimensionColor, mDimensionLineWeight, True
    AddText textLeft, textTop, textWidth, textHeight, textValue, rotation, mDimensionColor, True, mDimensionFontSize
End Sub

' Текстовая модельная привязка использует общую типографику размеров.
' Она показывается вместе с размерами, без фиктивной размерной линии.
Public Sub AddPointLabel(ByVal x As Double, ByVal y As Double, ByVal textValue As String)
    If Not mDimensionsEnabled Then Exit Sub
    Dim widthValue As Double, heightValue As Double
    widthValue = MaxDouble(90# * mVisualScale, Len(textValue) * mDimensionFontSize * 0.62)
    heightValue = 1.8 * mDimensionFontSize
    AddText ModelToChartX(x) - widthValue / 2#, ModelToChartY(y) - heightValue / 2#, _
        widthValue, heightValue, textValue, 0#, mDimensionColor, True, mDimensionFontSize
End Sub

' ------------------------------
' Групповая подпись арматуры
' ------------------------------
' Подпись центрируется по группе стержней. Опорная точка сначала переносится
' по нормали к бетону; Outside/Inside и Offset отсчитываются от этой границы.
Public Sub AddRebarLabel(ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double, _
        ByVal normalX As Double, ByVal normalY As Double, ByVal textValue As String, Optional ByVal boundaryDistance As Double = 0#)
    If Not mRebarLabelsEnabled Then Exit Sub
    Dim dirX As Double, dirY As Double
    dirX = endX - startX
    dirY = endY - startY
    If Abs(dirX) + Abs(dirY) <= GEOM_TOLERANCE And Abs(normalX) > GEOM_TOLERANCE And Abs(normalY) > GEOM_TOLERANCE Then ' Для одиночного стержня направление линии подписи определяем по заданной нормали. Обычный выбор горизонтальной или вертикальной подписи сохраняется.
        dirX = normalY: dirY = -normalX
    End If
    NormalizeModelVector dirX, dirY, 1#, 0#
    NormalizeModelVector normalX, normalY, -dirY, dirX

    Dim normalPxX As Double, normalPxY As Double
    PixelUnitVector normalX, normalY, normalPxX, normalPxY
    ApplyPlacement normalPxX, normalPxY, mRebarPlacement

    Dim dirPxX As Double, dirPxY As Double
    PixelUnitVector dirX, dirY, dirPxX, dirPxY

    Dim anchorX As Double, anchorY As Double
    anchorX = 0.5 * (startX + endX) + normalX * boundaryDistance
    anchorY = 0.5 * (startY + endY) + normalY * boundaryDistance

    Dim labelOffset As Double
    labelOffset = mRebarBaseGap

    Dim anchorPxX As Double, anchorPxY As Double
    anchorPxX = ModelToChartX(anchorX) + normalPxX * labelOffset
    anchorPxY = ModelToChartY(anchorY) + normalPxY * labelOffset

    Dim rotation As Double
    rotation = AnnotationRotation(dirX, dirY)
    If rotation = 90# And normalPxX > 0# Then rotation = -90# ' Правая вертикальная выноска читается сверху вниз, левая — снизу вверх. Сторону определяет фактическая нормаль после применения Inside/Outside.

    Dim textLeft As Double, textTop As Double, textWidth As Double, textHeight As Double
    RebarTextBox anchorPxX, anchorPxY, normalPxX, normalPxY, rotation, textValue, textLeft, textTop, textWidth, textHeight

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

' Размещает текст размера у середины его линии. Общая функция учитывает
' поворот рамки и размещает горизонтальный/вертикальный текст относительно размерной линии.
Private Sub DimensionTextBox(ByVal sx As Double, ByVal sy As Double, ByVal ex As Double, ByVal ey As Double, _
        ByVal normalPxX As Double, ByVal normalPxY As Double, ByVal rotation As Double, ByVal textValue As String, _
        ByRef leftPos As Double, ByRef topPos As Double, ByRef widthValue As Double, ByRef heightValue As Double)
    PositionTextBox 0.5 * (sx + ex), 0.5 * (sy + ey), normalPxX, normalPxY, rotation, textValue, _
        mDimensionFontSize, mDimensionTextGapBase, mDimensionTextUnits, leftPos, topPos, widthValue, heightValue
End Sub

' Размещает групповую подпись у её опорной линии. Положение зависит от
' нормали и TextGap, а рамка наклонного текста поворачивается вместе с линией.
Private Sub RebarTextBox(ByVal anchorPxX As Double, ByVal anchorPxY As Double, ByVal normalPxX As Double, _
        ByVal normalPxY As Double, ByVal rotation As Double, ByVal textValue As String, ByRef leftPos As Double, ByRef topPos As Double, _
        ByRef widthValue As Double, ByRef heightValue As Double)
    PositionTextBox anchorPxX, anchorPxY, normalPxX, normalPxY, rotation, textValue, _
        mRebarFontSize, mRebarTextGapBase, mRebarTextUnits, leftPos, topPos, widthValue, heightValue
End Sub

' Возвращает читаемый угол по фактическому направлению линии на экране.
' Разные масштабы X/Y и переворот Y учитываются до общей математики углов.
Private Function AnnotationRotation(ByVal dirX As Double, ByVal dirY As Double) As Double
    Dim px As Double, py As Double
    PixelUnitVector dirX, dirY, px, py
    AnnotationRotation = modGeometryPrimitives.ReadableLineAngle(px, py) * 180# / GEOM_PI
End Function

' Наклонный текст хранится в обычной горизонтальной рамке и затем поворачивается.
' Центр текста отнесён от линии на TextGap, как у горизонтальных/вертикальных подписей.
' При растянутой рамке графика учитывается
' угол между нормалью модели и перпендикуляром к экранной линии. Для 0/90 градусов
' размеры блока и зазор вычисляются непосредственно, включая вертикальный текст.
Private Sub PositionTextBox(ByVal anchorX As Double, ByVal anchorY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal rotation As Double, ByVal textValue As String, ByVal fontSize As Double, ByVal gapBase As Double, ByVal textUnits As String, _
        ByRef leftPos As Double, ByRef topPos As Double, ByRef widthValue As Double, ByRef heightValue As Double)
    widthValue = 150# * mVisualScale
    heightValue = MaxDouble(18# * mVisualScale, fontSize * 1.65)
    Dim gap As Double, angle As Double, normalProjection As Double, swap As Double
    gap = ClampDouble(TextMetricToChartPoints(gapBase, textUnits), 0#, 80#)
    If Abs(rotation) = 90# Then
        swap = widthValue: widthValue = heightValue: heightValue = swap
    ElseIf rotation <> 0# Then
        widthValue = MaxDouble(fontSize * 2#, Len(textValue) * fontSize * 0.62) ' Наклонной подписи не нужна широкая пустая рамка: длина следует тексту.
        angle = rotation * GEOM_PI / 180#
        normalProjection = Abs(-Sin(angle) * normalX + Cos(angle) * normalY) ' Проекция экранной нормали на перпендикуляр к повёрнутой строке текста задаёт фактический зазор.
        gap = gap / MaxDouble(normalProjection, 0.000001) ' Для корректной нормали проекция ненулевая; защитный минимум не допускает деления на ноль.
    End If
    leftPos = anchorX + normalX * gap - widthValue / 2#
    topPos = anchorY + normalY * gap - heightValue / 2#
End Sub

' При Inside направляет вынос аннотации внутрь сечения.
' Положение линии размера или подписи при этом задаётся теми же исходными точками.
Private Sub ApplyPlacement(ByRef normalPxX As Double, ByRef normalPxY As Double, ByVal placement As String)
    If StrComp(placement, "Inside", vbTextCompare) = 0 Then
        normalPxX = -normalPxX
        normalPxY = -normalPxY
    End If
End Sub

' Приводит направление к единичной длине. Если вектор нулевой,
' использует переданное запасное направление без деления на ноль.
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

' Возвращает число мм в одной выходной единице длины результата,
' например 1000 для метров.
Private Function LengthFactorToMm(ByVal unitName As String) As Double
    Dim units As CUnitSystem
    Set units = New CUnitSystem
    LengthFactorToMm = units.OutputLengthToInternal(1#, unitName)
End Function

' Добавляет подготовленную линию в общий список; сам Shape создаётся при отрисовке.
Private Sub AddLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal colorValue As Long, ByVal weight As Double, ByVal hasArrowheads As Boolean)
    AppendItem ITEM_LINE, x1, y1, x2, y2, vbNullString, 0#, colorValue, weight, False, hasArrowheads
End Sub

' Добавляет подготовленный текст в общий список; сам Shape
' создаётся при отрисовке внутри нужной диаграммы.
Private Sub AddText(ByVal leftPos As Double, ByVal topPos As Double, ByVal widthValue As Double, _
        ByVal heightValue As Double, ByVal textValue As String, ByVal rotation As Double, _
        ByVal colorValue As Long, ByVal isBold As Boolean, ByVal fontSize As Double)
    AppendItem ITEM_TEXT, leftPos, topPos, widthValue, heightValue, textValue, rotation, colorValue, 0#, isBold, False, fontSize
End Sub

' Сохраняет вид, координаты и оформление элемента в общих массивах.
' Отдельный объект на каждую линию или надпись не создаётся,
' чтобы обновление схемы не замедлялось.
Private Sub AppendItem(ByVal kind As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
        ByVal textValue As String, ByVal rotation As Double, ByVal colorValue As Long, ByVal weight As Double, _
        ByVal isBold As Boolean, ByVal hasArrowheads As Boolean, Optional ByVal fontSize As Double = 0#)
    EnsureItemCapacity mCount + 1
    mCount = mCount + 1

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

' Растит закрытые массивы геометрически, чтобы подготовка N подписей не
' копировала весь накопленный набор на каждой итерации. Повторный Initialize
' оставляет резерв и обнуляет только Count; заполненная часть перезаписывается.
Private Sub EnsureItemCapacity(ByVal required As Long)
    If required <= mCapacity Then Exit Sub
    Dim capacity As Long: capacity = required
    If mCapacity = 0 Then
        If capacity < 64 Then capacity = 64
    ElseIf mCapacity <= 1073741823 Then
        If capacity < mCapacity * 2 Then capacity = mCapacity * 2
    End If
    ReDim Preserve mKind(1 To capacity)
    ReDim Preserve mX1(1 To capacity)
    ReDim Preserve mY1(1 To capacity)
    ReDim Preserve mX2(1 To capacity)
    ReDim Preserve mY2(1 To capacity)
    ReDim Preserve mText(1 To capacity)
    ReDim Preserve mRotation(1 To capacity)
    ReDim Preserve mColor(1 To capacity)
    ReDim Preserve mWeight(1 To capacity)
    ReDim Preserve mBold(1 To capacity)
    ReDim Preserve mArrowheads(1 To capacity)
    ReDim Preserve mFontSize(1 To capacity)
    mCapacity = capacity
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

' Возвращает большее из двух чисел при расчёте размеров оформления.
Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

' Возвращает меньшее из двух чисел при выборе масштаба и габарита текста.
Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function

' Возвращает число подготовленных линий и текстов. CSectionPlotter
' создаёт по этим данным Chart.Shapes с заданным оформлением.
Public Property Get Count() As Long: Count = mCount: End Property
' Возвращает вид подготовленного графического элемента: линия или текст.
Public Property Get ItemKind(ByVal index As Long) As String: ItemKind = mKind(index): End Property
' Для линии: X начала; для текста: left.
Public Property Get X1(ByVal index As Long) As Double: X1 = mX1(index): End Property
' Для линии: Y начала; для текста: top.
Public Property Get Y1(ByVal index As Long) As Double: Y1 = mY1(index): End Property
' Для линии: X конца; для текста: width.
Public Property Get X2(ByVal index As Long) As Double: X2 = mX2(index): End Property
' Для линии: Y конца; для текста: height.
Public Property Get Y2(ByVal index As Long) As Double: Y2 = mY2(index): End Property
' Текст подписи; для линий пусто.
Public Property Get Text(ByVal index As Long) As String: Text = mText(index): End Property
' Поворот текстового блока в градусах.
Public Property Get Rotation(ByVal index As Long) As Double: Rotation = mRotation(index): End Property
' Цвет элемента Chart.Shape.
Public Property Get Color(ByVal index As Long) As Long: Color = mColor(index): End Property
' Толщина линии; для текста не используется.
Public Property Get Weight(ByVal index As Long) As Double: Weight = mWeight(index): End Property
' Нужно ли делать текст полужирным.
Public Property Get Bold(ByVal index As Long) As Boolean: Bold = mBold(index): End Property
' Нужно ли добавить стрелки с обеих сторон линии.
Public Property Get Arrowheads(ByVal index As Long) As Boolean: Arrowheads = mArrowheads(index): End Property
' Размер шрифта конкретной текстовой подписи.
Public Property Get FontSize(ByVal index As Long) As Double: FontSize = mFontSize(index): End Property
' Общий масштаб оформления относительно базового размера Chart.
Public Property Get VisualScale() As Double: VisualScale = mVisualScale: End Property
' Активна ли группа размеров; выключенные параметры не читаются.
Public Property Get DimensionsEnabled() As Boolean: DimensionsEnabled = mDimensionsEnabled: End Property
' Активна ли группа подписей арматуры.
Public Property Get RebarLabelsEnabled() As Boolean: RebarLabelsEnabled = mRebarLabelsEnabled: End Property
