VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Унифицированная расчетная модель сечения
' ==========================================================================
' CSectionModel - это единственный формат геометрии, который видит расчетное
' ядро. Неважно, пришла сетка из встроенного генератора или из AutoCAD import:
' ниже по цепочке solver получает один и тот же набор бетонных элементов,
' стержней арматуры и semantic-аннотаций для вывода.
'
' Важно: расчетные имена элементов C1/R1 создаются здесь. Генераторы
' геометрии и импорт из AutoCAD не должны придумывать собственный формат имен
' для расчетного ядра.

Private mSourceType As String ' Диагностическая метка источника модели: Generated, AutoCADImport и т.п.

Private mConcreteID() As String ' Расчетные имена бетонных элементов: C1, C2, ...
Private mConcreteSourceName() As String ' Необязательное исходное имя области для трассировки.
Private mConcreteSourceHandle() As String ' Необязательный AutoCAD handle для поиска исходного Region.
Private mConcreteX() As Double ' X центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteY() As Double ' Y центра тяжести бетонного элемента во внутренней системе координат, мм.
Private mConcreteArea() As Double ' Площадь бетонного элемента, мм2.
Private mConcreteMaterialID() As Long ' Номер/тип материала бетона для будущего расширения на разные материалы.
Private mConcreteShapeType() As String ' Исходная геометрическая подсказка источника: Rectangle для известных ячеек, Region для AutoCAD.
Private mConcreteWidth() As Double ' Реальная ширина элемента, только если источник достоверно знает прямоугольную оболочку; иначе 0.
Private mConcreteHeight() As Double ' Реальная высота элемента, только если источник достоверно знает прямоугольную оболочку; иначе 0.
Private mConcreteRotation() As Double ' Угол достоверной оболочки или угол простой грани почти изотропного Region, рад.
Private mConcreteLocalIx() As Double ' Реальный собственный Ix элемента вокруг его центра, мм4; не заменяется оболочкой.
Private mConcreteLocalIy() As Double ' Реальный собственный Iy элемента вокруг его центра, мм4; не заменяется оболочкой.
Private mConcreteLocalIxy() As Double ' Реальный собственный центробежный момент элемента, мм4; не заменяется оболочкой.
Private mConcreteComment() As String ' Короткая диагностическая заметка для Results.
Private mConcreteCount As Long ' Фактическое число бетонных элементов в модели.

Private mRebarID() As String ' Расчетные имена стержней: R1, R2, ...
Private mRebarSourceName() As String ' Необязательное исходное имя арматурной области.
Private mRebarSourceHandle() As String ' Необязательный AutoCAD handle арматурного Region.
Private mRebarX() As Double ' X центра стержня, мм.
Private mRebarY() As Double ' Y центра стержня, мм.
Private mRebarArea() As Double ' Площадь стержня, мм2.
Private mRebarDiameter() As Double ' Диаметр стержня, мм.
Private mRebarSteelClass() As String ' Расчетная подпись арматуры; сейчас хранит профиль стержней для трассировки.
Private mRebarMaterialID() As Long ' Номер/тип материала арматуры для будущего расширения.
Private mRebarComment() As String ' Диагностическая заметка о грани, ряде или источнике стержня.
Private mRebarCount As Long ' Фактическое число стержней в модели.

Private mAnnotations As CSectionAnnotations ' Semantic-аннотации, сформированные генератором сечения.

Private Const LOCAL_INERTIA_SQUARE_REL_TOLERANCE As Double = 0.000001 ' Допуск, при котором главные локальные I считаются изотропными.
Private Const LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE As Double = 0.01 ' Допуск проверки, что реальные A/I совместимы с прямоугольником.

' Очищает модель перед новым построением сечения.
' После Clear объект снова готов принимать бетонные элементы, арматуру и
' аннотации без следов предыдущего расчета.
Public Sub Clear()
    mSourceType = vbNullString

    Erase mConcreteID
    Erase mConcreteSourceName
    Erase mConcreteSourceHandle
    Erase mConcreteX
    Erase mConcreteY
    Erase mConcreteArea
    Erase mConcreteMaterialID
    Erase mConcreteShapeType
    Erase mConcreteWidth
    Erase mConcreteHeight
    Erase mConcreteRotation
    Erase mConcreteLocalIx
    Erase mConcreteLocalIy
    Erase mConcreteLocalIxy
    Erase mConcreteComment
    mConcreteCount = 0

    Erase mRebarID
    Erase mRebarSourceName
    Erase mRebarSourceHandle
    Erase mRebarX
    Erase mRebarY
    Erase mRebarArea
    Erase mRebarDiameter
    Erase mRebarSteelClass
    Erase mRebarMaterialID
    Erase mRebarComment
    mRebarCount = 0

    Set mAnnotations = New CSectionAnnotations
End Sub

Public Property Let SourceType(ByVal value As String)
    mSourceType = Trim$(value)
End Property

Public Property Get SourceType() As String
    SourceType = mSourceType
End Property

' Добавляет одно бетонное волокно или импортированную бетонную область.
' Координаты и площадь уже должны быть во внутренних единицах проекта. Метод
' только проверяет минимальную физическую корректность и присваивает элементу
' стабильное расчетное имя Cn.
Public Sub AddConcreteElement(ByVal xCoord As Double, ByVal yCoord As Double, ByVal area As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal shapeType As String = "Rectangle", _
        Optional ByVal width As Double = 0#, Optional ByVal height As Double = 0#, _
        Optional ByVal rotation As Double = 0#, Optional ByVal comment As String = vbNullString, _
        Optional ByVal localIx As Double = 0#, Optional ByVal localIy As Double = 0#, _
        Optional ByVal localIxy As Double = 0#)
    If area <= 0# Then Err.Raise vbObjectError + 5200, "CSectionModel", "Площадь бетонного элемента должна быть положительной."

    ' Массивы хранятся параллельно, чтобы горячие расчетные циклы работали без
    ' коллекций и COM-объектов на каждом элементе.
    mConcreteCount = mConcreteCount + 1
    ReDim Preserve mConcreteID(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceName(1 To mConcreteCount)
    ReDim Preserve mConcreteSourceHandle(1 To mConcreteCount)
    ReDim Preserve mConcreteX(1 To mConcreteCount)
    ReDim Preserve mConcreteY(1 To mConcreteCount)
    ReDim Preserve mConcreteArea(1 To mConcreteCount)
    ReDim Preserve mConcreteMaterialID(1 To mConcreteCount)
    ReDim Preserve mConcreteShapeType(1 To mConcreteCount)
    ReDim Preserve mConcreteWidth(1 To mConcreteCount)
    ReDim Preserve mConcreteHeight(1 To mConcreteCount)
    ReDim Preserve mConcreteRotation(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIx(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIy(1 To mConcreteCount)
    ReDim Preserve mConcreteLocalIxy(1 To mConcreteCount)
    ReDim Preserve mConcreteComment(1 To mConcreteCount)

    mConcreteID(mConcreteCount) = "C" & CStr(mConcreteCount)
    mConcreteSourceName(mConcreteCount) = sourceName
    mConcreteSourceHandle(mConcreteCount) = sourceHandle
    mConcreteX(mConcreteCount) = xCoord
    mConcreteY(mConcreteCount) = yCoord
    mConcreteArea(mConcreteCount) = area
    mConcreteMaterialID(mConcreteCount) = materialID
    mConcreteShapeType(mConcreteCount) = shapeType
    mConcreteWidth(mConcreteCount) = width
    mConcreteHeight(mConcreteCount) = height
    mConcreteRotation(mConcreteCount) = rotation
    mConcreteLocalIx(mConcreteCount) = localIx
    mConcreteLocalIy(mConcreteCount) = localIy
    mConcreteLocalIxy(mConcreteCount) = localIxy
    If Not IsPositiveDefiniteInertia(mConcreteLocalIx(mConcreteCount), _
            mConcreteLocalIy(mConcreteCount), mConcreteLocalIxy(mConcreteCount)) Then
        If HasExplicitRectangleGeometry(mConcreteCount) Then
            ' Генератор сетки может передать реальные размеры прямоугольной ячейки.
            ' В этом случае CSectionModel сам формирует расчетные локальные I.
            ' AutoCAD importer такие размеры не угадывает и сюда их не передает.
            RectangleLocalInertiaComponents area, width, height, rotation, _
                mConcreteLocalIx(mConcreteCount), _
                mConcreteLocalIy(mConcreteCount), _
                mConcreteLocalIxy(mConcreteCount)
        End If
    End If
    mConcreteComment(mConcreteCount) = comment
End Sub

' Назначает последним fallback-квадратам общий угол ориентации.
' Метод оставлен для совместимости со старыми snapshot-данными: новые AutoCAD
' Region получают угол простой грани сразу при импорте, а интерпретацию формы
' выполняет ConcreteBoundaryRectangle.
Public Sub ApplyAverageRotationToEquivalentAreaFallbacks()
    Dim fallbackRotation As Double
    fallbackRotation = AverageKnownConcreteElementRotation()

    Dim i As Long
    For i = 1 To mConcreteCount
        If IsConcreteEquivalentAreaFallback(i) And Abs(mConcreteRotation(i)) <= GEOM_TOLERANCE Then
            mConcreteRotation(i) = fallbackRotation
        End If
    Next i
End Sub

' Находит средний угол бетонных элементов с известной прямоугольной оболочкой.
' Угол прямоугольника периодичен через 180 градусов, поэтому усреднение идет
' через двойной угол: так ориентации 0 и 180 градусов не гасят друг друга.
Public Function AverageKnownConcreteElementRotation() As Double
    Dim sumCos As Double
    Dim sumSin As Double
    Dim count As Long
    Dim i As Long
    For i = 1 To mConcreteCount
        If HasExplicitRectangleGeometry(i) Then
            sumCos = sumCos + Cos(2# * mConcreteRotation(i))
            sumSin = sumSin + Sin(2# * mConcreteRotation(i))
            count = count + 1
        End If
    Next i

    If count > 0 And (Abs(sumCos) > GEOM_TOLERANCE Or Abs(sumSin) > GEOM_TOLERANCE) Then
        AverageKnownConcreteElementRotation = 0.5 * Atn2Local(sumSin, sumCos)
    End If
End Function

' Проверяет, что бетонный элемент является последним fallback-ом по площади.
' Для новых данных это определяется через интерпретацию оболочки, а не через
' исходную строковую подсказку importer-а.
Public Function IsConcreteEquivalentAreaFallback(ByVal index As Long) As Boolean
    CheckConcreteIndex index
    IsConcreteEquivalentAreaFallback = _
        (StrComp(ConcreteBoundaryShapeType(index), "Equivalent square", vbTextCompare) = 0)
End Function

' Возвращает прямоугольную оболочку бетонного элемента для геометрических задач.
' Это единственная точка, где реальные A/I интерпретируются как Width/Height/
' Rotation. Расчетные инерции от этой оболочки обратно не пересчитываются.
Public Function ConcreteBoundaryRectangle(ByVal index As Long, ByRef widthValue As Double, _
        ByRef heightValue As Double, ByRef rotationValue As Double) As Boolean
    Dim statusText As String
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteBoundaryRectangle = (widthValue > GEOM_TOLERANCE And heightValue > GEOM_TOLERANCE)
End Function

' Возвращает статус интерпретации геометрической оболочки бетонного элемента.
' Возможные значения: Rectangle, Equivalent rectangle, Equivalent square.
Public Function ConcreteBoundaryShapeType(ByVal index As Long) As String
    Dim statusText As String
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    InterpretConcreteBoundary index, statusText, widthValue, heightValue, rotationValue
    ConcreteBoundaryShapeType = statusText
End Function

' Возвращает половину проекции интерпретированной оболочки элемента на ось.
' Метод нужен габаритам, a_s, ядровым расстояниям, схеме и AutoCAD export.
Public Function ConcreteBoundaryHalfProjection(ByVal index As Long, ByVal axisX As Double, _
        ByVal axisY As Double) As Double
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    If ConcreteBoundaryRectangle(index, widthValue, heightValue, rotationValue) Then
        ConcreteBoundaryHalfProjection = RectangleHalfProjection(widthValue, heightValue, rotationValue, axisX, axisY)
    End If
End Function

' Возвращает реальные локальные инерции элемента. Если источник вообще не дал
' пригодные I, применяется последний расчетный fallback: квадрат той же площади.
Public Sub ConcreteLocalInertiaComponents(ByVal index As Long, ByRef localIx As Double, _
        ByRef localIy As Double, ByRef localIxy As Double)
    CheckConcreteIndex index

    localIx = mConcreteLocalIx(index)
    localIy = mConcreteLocalIy(index)
    localIxy = mConcreteLocalIxy(index)
    If IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then Exit Sub

    EquivalentSquareLocalInertiaComponents mConcreteArea(index), localIx, localIy, localIxy
End Sub

' Проверяет, что локальная матрица инерций почти изотропна и не задает
' однозначного направления главных осей. AutoCAD importer использует это,
' чтобы обращаться к геометрии Region только за углом простой грани.
Public Function LocalInertiaIsNearlySquare(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double, Optional ByVal relativeTolerance As Double = LOCAL_INERTIA_SQUARE_REL_TOLERANCE) As Boolean
    If Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then Exit Function

    Dim i1 As Double
    Dim i2 As Double
    Dim angle1 As Double
    PrincipalLocalInertia localIx, localIy, localIxy, i1, i2, angle1
    If i1 <= GEOM_TOLERANCE Then Exit Function
    LocalInertiaIsNearlySquare = (Abs(i1 - i2) / i1 <= relativeTolerance)
End Function

' Интерпретирует оболочку бетонного элемента по единому каскаду:
' 1) явно известный прямоугольник источника;
' 2) прямоугольник, восстановленный по реальным A/I;
' 3) эквивалентный квадрат по площади.
Private Sub InterpretConcreteBoundary(ByVal index As Long, ByRef statusText As String, _
        ByRef widthValue As Double, ByRef heightValue As Double, ByRef rotationValue As Double)
    CheckConcreteIndex index

    If HasExplicitRectangleGeometry(index) Then
        statusText = "Rectangle"
        widthValue = mConcreteWidth(index)
        heightValue = mConcreteHeight(index)
        rotationValue = NormalizeAnglePi(mConcreteRotation(index))
        Exit Sub
    End If

    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy

    If LocalInertiaIsNearlySquare(localIx, localIy, localIxy) Then
        statusText = "Equivalent square"
        widthValue = Sqr(mConcreteArea(index))
        heightValue = widthValue
        rotationValue = NormalizeAnglePi(mConcreteRotation(index))
        Exit Sub
    End If

    If TryInterpretAnisotropicBoundary(mConcreteArea(index), localIx, localIy, localIxy, _
            statusText, widthValue, heightValue, rotationValue) Then Exit Sub

    statusText = "Equivalent square"
    widthValue = Sqr(mConcreteArea(index))
    heightValue = widthValue
    rotationValue = NormalizeAnglePi(mConcreteRotation(index))
End Sub

' Восстанавливает оболочку неизотропного Region. Если реальные A/I совпадают
' с настоящим прямоугольником, статус Rectangle; иначе строится area-preserving
' эквивалентный прямоугольник с тем же отношением главных инерций.
Private Function TryInterpretAnisotropicBoundary(ByVal areaValue As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByRef statusText As String, ByRef widthValue As Double, _
        ByRef heightValue As Double, ByRef rotationValue As Double) As Boolean
    If areaValue <= GEOM_TOLERANCE Then Exit Function

    Dim i1 As Double
    Dim i2 As Double
    Dim angle1 As Double
    PrincipalLocalInertia localIx, localIy, localIxy, i1, i2, angle1
    If i1 <= GEOM_TOLERANCE Or i2 <= GEOM_TOLERANCE Then Exit Function

    Dim rectangleWidth As Double
    Dim rectangleHeight As Double
    rectangleWidth = Sqr(12# * i2 / areaValue)
    rectangleHeight = Sqr(12# * i1 / areaValue)

    If rectangleWidth > GEOM_TOLERANCE And rectangleHeight > GEOM_TOLERANCE Then
        If Abs(rectangleWidth * rectangleHeight - areaValue) / areaValue <= _
                LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE Then
            statusText = "Rectangle"
            widthValue = rectangleWidth
            heightValue = rectangleHeight
            rotationValue = NormalizeAnglePi(angle1)
            TryInterpretAnisotropicBoundary = True
            Exit Function
        End If
    End If

    statusText = "Equivalent rectangle"
    heightValue = Sqr(areaValue) * (i1 / i2) ^ 0.25
    widthValue = Sqr(areaValue) * (i2 / i1) ^ 0.25
    rotationValue = NormalizeAnglePi(angle1)
    TryInterpretAnisotropicBoundary = (widthValue > GEOM_TOLERANCE And heightValue > GEOM_TOLERANCE)
End Function

' Возвращает главные локальные инерции и направление оси I1.
' Угол считается от глобальной +X против часовой стрелки; ось физически
' беззнаковая, поэтому +180 градусов означает то же направление.
Public Sub PrincipalLocalInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double, ByRef i1 As Double, ByRef i2 As Double, _
        ByRef angle1 As Double)
    Dim averageI As Double
    Dim radiusI As Double
    averageI = (localIx + localIy) / 2#
    radiusI = Sqr(((localIx - localIy) / 2#) ^ 2 + localIxy ^ 2)
    i1 = averageI + radiusI
    i2 = averageI - radiusI

    If LocalInertiaIsNumericallyIsotropic(localIx, localIy, localIxy) Then
        angle1 = 0#
    Else
        angle1 = NormalizeAnglePi(0.5 * Atn2Local(-2# * localIxy, localIx - localIy))
    End If
End Sub

' Явно известная прямоугольная оболочка допустима только как геометрический
' факт источника. Она не отменяет реальные Ix/Iy/Ixy, если они уже переданы.
Private Function HasExplicitRectangleGeometry(ByVal index As Long) As Boolean
    HasExplicitRectangleGeometry = _
        StrComp(Trim$(mConcreteShapeType(index)), "Rectangle", vbTextCompare) = 0 And _
        mConcreteWidth(index) > GEOM_TOLERANCE And _
        mConcreteHeight(index) > GEOM_TOLERANCE And _
        mConcreteArea(index) > GEOM_TOLERANCE And _
        Abs(mConcreteWidth(index) * mConcreteHeight(index) - mConcreteArea(index)) / mConcreteArea(index) <= _
            LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE
End Function

' Считает Ix/Iy/Ixy прямоугольной ячейки вокруг собственного центра.
' Используется только когда прямоугольная форма известна до расчета A/I,
' например у встроенного генератора сетки.
Private Sub RectangleLocalInertiaComponents(ByVal areaValue As Double, _
        ByVal widthValue As Double, ByVal heightValue As Double, ByVal angleValue As Double, _
        ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    Dim ixLocal As Double
    Dim iyLocal As Double
    Dim c As Double
    Dim s As Double
    ixLocal = areaValue * heightValue * heightValue / 12#
    iyLocal = areaValue * widthValue * widthValue / 12#
    c = Cos(angleValue)
    s = Sin(angleValue)

    localIx = ixLocal * c * c + iyLocal * s * s
    localIy = ixLocal * s * s + iyLocal * c * c
    localIxy = (iyLocal - ixLocal) * s * c
End Sub

' Собственные инерции последнего fallback-квадрата той же площади.
Private Sub EquivalentSquareLocalInertiaComponents(ByVal areaValue As Double, _
        ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    localIx = areaValue * areaValue / 12#
    localIy = localIx
    localIxy = 0#
End Sub

' Возвращает половину проекции прямоугольной оболочки на произвольную ось.
Private Function RectangleHalfProjection(ByVal widthValue As Double, ByVal heightValue As Double, _
        ByVal angleValue As Double, ByVal axisX As Double, ByVal axisY As Double) As Double
    Dim ux As Double
    Dim uy As Double
    Dim vx As Double
    Dim vy As Double
    ux = Cos(angleValue)
    uy = Sin(angleValue)
    vx = -uy
    vy = ux
    RectangleHalfProjection = 0.5 * (widthValue * Abs(axisX * ux + axisY * uy) + _
        heightValue * Abs(axisX * vx + axisY * vy))
End Function

' Проверяет положительную определенность локальной матрицы инерций.
Private Function IsPositiveDefiniteInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double) As Boolean
    If localIx <= GEOM_TOLERANCE Or localIy <= GEOM_TOLERANCE Then Exit Function

    Dim scaleValue As Double
    scaleValue = GeomMax(Abs(localIx * localIy), 1#)
    IsPositiveDefiniteInertia = ((localIx * localIy - localIxy * localIxy) / scaleValue > GEOM_TOLERANCE)
End Function

Private Function LocalInertiaIsNumericallyIsotropic(ByVal localIx As Double, _
        ByVal localIy As Double, ByVal localIxy As Double) As Boolean
    Dim scaleValue As Double
    scaleValue = GeomMax(GeomMax(Abs(localIx), Abs(localIy)), 1#)
    LocalInertiaIsNumericallyIsotropic = _
        Abs(localIx - localIy) / scaleValue <= LOCAL_INERTIA_SQUARE_REL_TOLERANCE And _
        Abs(localIxy) / scaleValue <= LOCAL_INERTIA_SQUARE_REL_TOLERANCE
End Function

' Нормализует угол беззнаковой оси в диапазон примерно [-pi/2; pi/2].
Private Function NormalizeAnglePi(ByVal angleValue As Double) As Double
    Do While angleValue > GEOM_PI / 2#
        angleValue = angleValue - GEOM_PI
    Loop
    Do While angleValue <= -GEOM_PI / 2#
        angleValue = angleValue + GEOM_PI
    Loop
    NormalizeAnglePi = angleValue
End Function

' Добавляет один стержень арматуры.
' Если площадь не передана явно, она считается по диаметру. Это удобно для
' встроенных раскладчиков, а AutoCAD import может передать площадь Region
' напрямую, если она известна надежнее диаметра.
Public Sub AddRebarElement(ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal materialID As Long = 1, Optional ByVal sourceName As String = vbNullString, _
        Optional ByVal sourceHandle As String = vbNullString, Optional ByVal comment As String = vbNullString)
    If diameter <= 0# Then Err.Raise vbObjectError + 5210, "CSectionModel", "Диаметр арматуры должен быть положительным."

    Dim actualArea As Double ' Площадь, которая реально попадет в расчетную модель.
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 5211, "CSectionModel", "Площадь арматурного элемента должна быть положительной."

    ' Нумерация R1/R2/... создается здесь, чтобы у всех источников геометрии
    ' был одинаковый формат расчетных имен.
    mRebarCount = mRebarCount + 1
    ReDim Preserve mRebarID(1 To mRebarCount)
    ReDim Preserve mRebarSourceName(1 To mRebarCount)
    ReDim Preserve mRebarSourceHandle(1 To mRebarCount)
    ReDim Preserve mRebarX(1 To mRebarCount)
    ReDim Preserve mRebarY(1 To mRebarCount)
    ReDim Preserve mRebarArea(1 To mRebarCount)
    ReDim Preserve mRebarDiameter(1 To mRebarCount)
    ReDim Preserve mRebarSteelClass(1 To mRebarCount)
    ReDim Preserve mRebarMaterialID(1 To mRebarCount)
    ReDim Preserve mRebarComment(1 To mRebarCount)

    mRebarID(mRebarCount) = "R" & CStr(mRebarCount)
    mRebarSourceName(mRebarCount) = sourceName
    mRebarSourceHandle(mRebarCount) = sourceHandle
    mRebarX(mRebarCount) = xCoord
    mRebarY(mRebarCount) = yCoord
    mRebarArea(mRebarCount) = actualArea
    mRebarDiameter(mRebarCount) = diameter
    mRebarSteelClass(mRebarCount) = steelClass
    mRebarMaterialID(mRebarCount) = materialID
    mRebarComment(mRebarCount) = comment
End Sub

Public Property Get ConcreteCount() As Long
    ConcreteCount = mConcreteCount
End Property

Public Property Get RebarCount() As Long
    RebarCount = mRebarCount
End Property

Public Property Get ConcreteID(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteID = mConcreteID(index)
End Property

Public Property Get ConcreteSourceName(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceName = mConcreteSourceName(index)
End Property

Public Property Get ConcreteSourceHandle(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteSourceHandle = mConcreteSourceHandle(index)
End Property

Public Property Get ConcreteX(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteX = mConcreteX(index)
End Property

Public Property Get ConcreteY(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteY = mConcreteY(index)
End Property

Public Property Get ConcreteArea(ByVal index As Long) As Double
    CheckConcreteIndex index
    ConcreteArea = mConcreteArea(index)
End Property

Public Property Get ConcreteMaterialID(ByVal index As Long) As Long
    CheckConcreteIndex index
    ConcreteMaterialID = mConcreteMaterialID(index)
End Property

Public Property Get ConcreteShapeType(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteShapeType = ConcreteBoundaryShapeType(index)
End Property

Public Property Get ConcreteWidth(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    ConcreteBoundaryRectangle index, widthValue, heightValue, rotationValue
    ConcreteWidth = widthValue
End Property

Public Property Get ConcreteHeight(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    ConcreteBoundaryRectangle index, widthValue, heightValue, rotationValue
    ConcreteHeight = heightValue
End Property

Public Property Get ConcreteRotation(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim widthValue As Double
    Dim heightValue As Double
    Dim rotationValue As Double
    ConcreteBoundaryRectangle index, widthValue, heightValue, rotationValue
    ConcreteRotation = rotationValue
End Property

Public Property Get ConcreteLocalIx(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIx = localIx
End Property

Public Property Get ConcreteLocalIy(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIy = localIy
End Property

Public Property Get ConcreteLocalIxy(ByVal index As Long) As Double
    CheckConcreteIndex index
    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ConcreteLocalInertiaComponents index, localIx, localIy, localIxy
    ConcreteLocalIxy = localIxy
End Property

Public Property Get ConcreteComment(ByVal index As Long) As String
    CheckConcreteIndex index
    ConcreteComment = mConcreteComment(index)
End Property

Public Property Get RebarID(ByVal index As Long) As String
    CheckRebarIndex index
    RebarID = mRebarID(index)
End Property

Public Property Get RebarSourceHandle(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceHandle = mRebarSourceHandle(index)
End Property

Public Property Get RebarX(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarX = mRebarX(index)
End Property

Public Property Get RebarY(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarY = mRebarY(index)
End Property

Public Property Get RebarArea(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarArea = mRebarArea(index)
End Property

Public Property Get RebarDiameter(ByVal index As Long) As Double
    CheckRebarIndex index
    RebarDiameter = mRebarDiameter(index)
End Property

Public Property Get RebarMaterialID(ByVal index As Long) As Long
    CheckRebarIndex index
    RebarMaterialID = mRebarMaterialID(index)
End Property

Public Property Get RebarComment(ByVal index As Long) As String
    CheckRebarIndex index
    RebarComment = mRebarComment(index)
End Property

Public Property Get Annotations() As CSectionAnnotations
    If mAnnotations Is Nothing Then Set mAnnotations = New CSectionAnnotations
    Set Annotations = mAnnotations
End Property

Public Property Get AnnotationCount() As Long
    If mAnnotations Is Nothing Then
        AnnotationCount = 0
    Else
        AnnotationCount = mAnnotations.Count
    End If
End Property

Private Sub CheckConcreteIndex(ByVal index As Long)
    If index < 1 Or index > mConcreteCount Then Err.Raise vbObjectError + 5220, "CSectionModel", "Индекс бетонного элемента вне допустимого диапазона."
End Sub

Private Sub CheckRebarIndex(ByVal index As Long)
    If index < 1 Or index > mRebarCount Then Err.Raise vbObjectError + 5221, "CSectionModel", "Индекс арматурного элемента вне допустимого диапазона."
End Sub

' Возвращает atan2(y, x) без зависимости от внешних модулей.
' Нужен только для усреднения ориентаций импортированных прямоугольных ячеек.
Private Function Atn2Local(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
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

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Эти трассировочные свойства нужны regression-тестам AutoCAD/import и старых
' сценариев, где проверяется перенос исходных имен и профиля арматуры внутрь
' CSectionModel. Обычные расчеты берут рабочие свойства стержня через ID,
' координаты, площадь, диаметр, материал и комментарий.

Public Property Get RebarSourceName(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSourceName = mRebarSourceName(index)
End Property

Public Property Get RebarSteelClass(ByVal index As Long) As String
    CheckRebarIndex index
    RebarSteelClass = mRebarSteelClass(index)
End Property
