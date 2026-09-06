VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoCADSectionModelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Импорт расчетной сетки из AutoCAD Region
' ==========================================================================
' Импортер читает активный чертеж AutoCAD и превращает Region на заданных
' слоях бетона и арматуры в CSectionModel. После этого расчетное ядро уже не
' знает, что геометрия пришла из AutoCAD: оно видит те же бетонные элементы
' и арматуру, что и при автоматической генерации.
'
' Класс не выполняет расчет и не пишет Results. Его ответственность - только
' аккуратно отфильтровать AutoCAD-сущности и заполнить модель сечения.

Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Техническая метка расчетной арматуры в snapshot геометрии.
Private Const RECTANGLE_BOUNDS_AREA_REL_TOL As Double = 0.02 ' Bounds считаются реальными сторонами Region, если площадь отличается не больше чем на 2%.
Private Const RECTANGLE_EDGE_REL_TOL As Double = 0.02 ' Допуск на равенство пар сторон при распознавании повернутого прямоугольника.
Private Const RECTANGLE_ORTHOGONAL_TOL As Double = 0.02 ' Допуск на ортогональность соседних сторон прямоугольного Region.

Public Function ImportFromActiveDocument(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки Config не переданы."

    Dim acad As Object
    On Error Resume Next
    Set acad = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    If acad Is Nothing Then Err.Raise vbObjectError + 4401, "CAutoCADSectionModelImporter", _
        "AutoCAD не открыт. Откройте AutoCAD с нужным чертежом и повторите импорт."

    Dim doc As Object
    On Error Resume Next
    Set doc = acad.ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then Err.Raise vbObjectError + 4402, "CAutoCADSectionModelImporter", _
        "AutoCAD открыт, но активный чертеж не найден. Откройте или создайте чертеж и повторите импорт."

    Set ImportFromActiveDocument = ImportFromModelSpace(doc.ModelSpace, _
        settings.GetString("AutoCAD.Import.ConcreteLayer", "Concrete"), _
        settings.GetString("AutoCAD.Import.RebarLayer", "Reinf"), _
        REBAR_STEEL_CLASS, _
        InputArea(settings.GetDouble("AutoCAD.Import.MinArea", 0.000001), units), units)
End Function

Public Function ImportFromModelSpace(ByVal modelSpace As Object, ByVal concreteLayer As String, _
        ByVal rebarLayer As String, ByVal steelClass As String, ByVal minArea As Double, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If modelSpace Is Nothing Then Err.Raise vbObjectError + 4403, "CAutoCADSectionModelImporter", "ModelSpace AutoCAD не передан."

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    Dim entity As Object
    For Each entity In modelSpace
        If IsRegionEntity(entity) Then
            If SameLayer(entity.Layer, concreteLayer) Then
                AddConcreteRegionEntity model, entity, minArea, units
            ElseIf SameLayer(entity.Layer, rebarLayer) Then
                AddRebarRegionEntity model, entity, steelClass, minArea, units
            End If
        End If
    Next entity

    model.ApplyAverageRotationToEquivalentAreaFallbacks
    ValidateImportedModel model
    Set ImportFromModelSpace = model
End Function

Private Sub AddConcreteRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, ByVal minArea As Double, ByVal units As CUnitSystem)
    Dim area As Double
    area = InputArea(CDbl(entity.Area), units)
    If area < minArea Then Exit Sub

    Dim centroid As Variant
    centroid = entity.Centroid

    Dim xCoord As Double
    Dim yCoord As Double
    xCoord = InputLength(CDbl(centroid(LBound(centroid))), units)
    yCoord = InputLength(CDbl(centroid(LBound(centroid) + 1)), units)

    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ReadCentralRegionInertia entity, area, xCoord, yCoord, units, localIx, localIy, localIxy

    Dim boundWidth As Double
    Dim boundHeight As Double
    Dim boundRotation As Double
    ReadRegionRectangleGeometry entity, units, area, boundWidth, boundHeight, boundRotation

    AddConcreteRegion model, area, xCoord, yCoord, localIx, localIy, localIxy, _
        boundWidth, boundHeight, boundRotation, entity.Layer, EntityHandle(entity)
End Sub

Private Sub AddRebarRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, _
        ByVal steelClass As String, ByVal minArea As Double, ByVal units As CUnitSystem)
    Dim area As Double
    area = InputArea(CDbl(entity.Area), units)
    If area < minArea Then Exit Sub

    Dim centroid As Variant
    centroid = entity.Centroid
    AddRebarRegion model, area, InputLength(CDbl(centroid(LBound(centroid))), units), _
        InputLength(CDbl(centroid(LBound(centroid) + 1)), units), steelClass, entity.Layer, EntityHandle(entity)
End Sub

Private Sub AddConcreteRegionArray(ByVal model As CSectionModel, ByVal regions As Variant, ByVal minArea As Double)
    If Not HasArrayRows(regions) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(regions, 1) To UBound(regions, 1)
        Dim area As Double
        area = CDbl(regions(rowIndex, 1))
        If area >= minArea Then
            AddConcreteRegion model, area, CDbl(regions(rowIndex, 2)), CDbl(regions(rowIndex, 3)), _
                OptionalRegionDouble(regions, rowIndex, 4), OptionalRegionDouble(regions, rowIndex, 5), _
                OptionalRegionDouble(regions, rowIndex, 6), OptionalRegionDouble(regions, rowIndex, 8), _
                OptionalRegionDouble(regions, rowIndex, 9), OptionalRegionDouble(regions, rowIndex, 10), _
                "Concrete", OptionalRegionString(regions, rowIndex, 7)
        End If
    Next rowIndex
End Sub

Private Sub AddRebarRegionArray(ByVal model As CSectionModel, ByVal regions As Variant, _
        ByVal steelClass As String, ByVal minArea As Double)
    If Not HasArrayRows(regions) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(regions, 1) To UBound(regions, 1)
        Dim area As Double
        area = CDbl(regions(rowIndex, 1))
        If area >= minArea Then
            AddRebarRegion model, area, CDbl(regions(rowIndex, 2)), CDbl(regions(rowIndex, 3)), _
                steelClass, "Reinf", OptionalRegionString(regions, rowIndex, 4)
        End If
    Next rowIndex
End Sub

Private Sub AddConcreteRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal realWidth As Double, ByVal realHeight As Double, ByVal realRotation As Double, _
        ByVal sourceName As String, ByVal sourceHandle As String)
    Dim width As Double
    Dim height As Double
    Dim rotation As Double
    Dim shapeType As String
    Dim commentText As String

    If IsReliableRectangleBounds(area, realWidth, realHeight) Then
        width = realWidth
        height = realHeight
        rotation = realRotation
        If Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
            RectangleLocalInertiaComponents area, width, height, rotation, localIx, localIy, localIxy
        End If
        shapeType = "Rectangle"
        commentText = "AutoCAD Region: реальные прямоугольные размеры этого бетонного элемента взяты из геометрии Region."
    ElseIf IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        ' Размеры элемента неизвестны, поэтому Width/Height в CSectionModel
        ' оставляем пустыми. Эквивалентный прямоугольник по A/I строится ниже
        ' только внутри CSectionPropertiesCalculator для h, r и границ.
        width = 0#
        height = 0#
        rotation = 0#
        shapeType = "Region"
        commentText = "AutoCAD Region: реальные размеры неизвестны; расчетные границы восстанавливаются по центральным инерциям."
    Else
        ApplyEquivalentSquareFallback area, width, height, rotation, localIx, localIy, localIxy, shapeType, commentText
    End If

    model.AddConcreteElement xCoord, yCoord, area, 1, sourceName, sourceHandle, _
        shapeType, width, height, rotation, commentText, _
        localIx, localIy, localIxy
End Sub

Private Sub AddRebarRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal steelClass As String, _
        ByVal sourceName As String, ByVal sourceHandle As String)
    Dim diameter As Double
    diameter = Sqr(4# * area / GEOM_PI)
    model.AddRebarElement xCoord, yCoord, diameter, area, steelClass, 1, sourceName, sourceHandle, _
        "AutoCAD Region: стержень арматуры; диаметр и радиус восстановлены по площади как эквивалентный круг."
End Sub

Private Sub ReadCentralRegionInertia(ByVal entity As Object, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal units As CUnitSystem, ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    localIx = 0#
    localIy = 0#
    localIxy = 0#

    On Error Resume Next
    Dim moments As Variant
    moments = entity.MomentOfInertia
    If Err.Number = 0 Then
        localIx = InputFourthPowerLength(CDbl(moments(LBound(moments))), units) - area * yCoord * yCoord
        localIy = InputFourthPowerLength(CDbl(moments(LBound(moments) + 1)), units) - area * xCoord * xCoord
    End If
    Err.Clear

    Dim product As Variant
    product = entity.ProductOfInertia
    If Err.Number = 0 Then
        If IsArray(product) Then
            localIxy = InputFourthPowerLength(CDbl(product(LBound(product))), units) - area * xCoord * yCoord
        Else
            localIxy = InputFourthPowerLength(CDbl(product), units) - area * xCoord * yCoord
        End If
    End If
    On Error GoTo 0

    If localIx <= 0# Or localIy <= 0# Then
        localIx = 0#
        localIy = 0#
        localIxy = 0#
    End If
End Sub

' Читает реальные размеры прямоугольного Region. Сначала используется быстрый
' BoundingBox для неповернутых ячеек. Если он не похож на реальную площадь,
' пробуем разложить копию Region на ребра и определить повернутый прямоугольник.
Private Sub ReadRegionRectangleGeometry(ByVal entity As Object, ByVal units As CUnitSystem, _
        ByVal area As Double, ByRef width As Double, ByRef height As Double, ByRef rotation As Double)
    rotation = 0#
    ReadRegionBounds entity, units, width, height
    If IsReliableRectangleBounds(area, width, height) Then Exit Sub

    If TryReadRotatedRectangleBounds(entity, units, area, width, height, rotation) Then Exit Sub
End Sub

' Читает осевой BoundingBox Region. Для повернутого прямоугольника это только
' внешняя рамка, поэтому результат отдельно проверяется по площади элемента.
Private Sub ReadRegionBounds(ByVal entity As Object, ByVal units As CUnitSystem, _
        ByRef width As Double, ByRef height As Double)
    width = 0#
    height = 0#

    On Error GoTo NoBounds
    Dim minPoint As Variant
    Dim maxPoint As Variant
    entity.GetBoundingBox minPoint, maxPoint
    If Not IsArray(minPoint) Or Not IsArray(maxPoint) Then GoTo NoBounds

    width = Abs(InputLength(CDbl(maxPoint(LBound(maxPoint))) - CDbl(minPoint(LBound(minPoint))), units))
    height = Abs(InputLength(CDbl(maxPoint(LBound(maxPoint) + 1)) - CDbl(minPoint(LBound(minPoint) + 1)), units))
    If width <= GEOM_TOLERANCE Or height <= GEOM_TOLERANCE Then
        width = 0#
        height = 0#
    End If
    Exit Sub

NoBounds:
    width = 0#
    height = 0#
End Sub

' Определяет повернутый прямоугольный Region по его ребрам. Метод работает с
' копией объекта AutoCAD и удаляет временные линии после чтения, чтобы импорт не
' менял исходную расчетную геометрию чертежа.
Private Function TryReadRotatedRectangleBounds(ByVal entity As Object, ByVal units As CUnitSystem, _
        ByVal area As Double, ByRef width As Double, ByRef height As Double, ByRef rotation As Double) As Boolean
    If area <= GEOM_TOLERANCE Then Exit Function

    Dim tempEntity As Object
    Dim pieces As Variant

    On Error GoTo Failed
    Set tempEntity = entity.Copy
    pieces = tempEntity.Explode
    DeleteAcadObjectQuietly tempEntity

    If Not IsArray(pieces) Then GoTo Cleanup

    Dim lengths(1 To 4) As Double
    Dim ux(1 To 4) As Double
    Dim uy(1 To 4) As Double
    Dim pointX(1 To 8) As Double
    Dim pointY(1 To 8) As Double
    Dim pointCount As Long
    Dim lineCount As Long

    Dim i As Long
    For i = LBound(pieces) To UBound(pieces)
        Dim piece As Object
        Set piece = pieces(i)

        If IsLineEntity(piece) Then
            lineCount = lineCount + 1
            If lineCount > 4 Then GoTo Cleanup
            ReadLineGeometry piece, units, lengths(lineCount), ux(lineCount), uy(lineCount), pointX, pointY, pointCount
        Else
            GoTo Cleanup
        End If
    Next i

    If lineCount <> 4 Or pointCount <> 4 Then GoTo Cleanup
    If Not ResolveRectangleFromEdges(lengths, ux, uy, area, width, height, rotation) Then GoTo Cleanup

    TryReadRotatedRectangleBounds = True

Cleanup:
    DeleteExplodedPiecesQuietly pieces
    Exit Function

Failed:
    On Error Resume Next
    DeleteAcadObjectQuietly tempEntity
    DeleteExplodedPiecesQuietly pieces
End Function

' Читает длину и направление линии AutoCAD во внутренних единицах проекта.
' Концы линии дополнительно собираются как уникальные вершины прямоугольника.
Private Sub ReadLineGeometry(ByVal lineEntity As Object, ByVal units As CUnitSystem, _
        ByRef lengthValue As Double, ByRef directionX As Double, ByRef directionY As Double, _
        ByRef pointX() As Double, ByRef pointY() As Double, ByRef pointCount As Long)
    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = lineEntity.StartPoint
    endPoint = lineEntity.EndPoint

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = InputLength(CDbl(startPoint(LBound(startPoint))), units)
    y1 = InputLength(CDbl(startPoint(LBound(startPoint) + 1)), units)
    x2 = InputLength(CDbl(endPoint(LBound(endPoint))), units)
    y2 = InputLength(CDbl(endPoint(LBound(endPoint) + 1)), units)

    Dim dx As Double
    Dim dy As Double
    dx = x2 - x1
    dy = y2 - y1
    lengthValue = Sqr(dx * dx + dy * dy)
    If lengthValue <= GEOM_TOLERANCE Then Err.Raise vbObjectError + 4430, "CAutoCADSectionModelImporter", _
        "В контуре AutoCAD Region найдено ребро нулевой длины."

    directionX = dx / lengthValue
    directionY = dy / lengthValue
    AddUniquePoint pointX, pointY, pointCount, x1, y1
    AddUniquePoint pointX, pointY, pointCount, x2, y2
End Sub

' По четырем ребрам восстанавливает стороны и поворот прямоугольника.
' Для квадратной ячейки поворот берется по первому ребру: при равных сторонах
' выбор соседнего ребра дает ту же геометрическую проекцию.
Private Function ResolveRectangleFromEdges(ByRef lengths() As Double, ByRef ux() As Double, _
        ByRef uy() As Double, ByVal area As Double, ByRef width As Double, _
        ByRef height As Double, ByRef rotation As Double) As Boolean
    Dim sideA As Double
    Dim sideB As Double
    Dim countA As Long
    Dim countB As Long
    Dim widthIndex As Long

    sideA = lengths(1)
    countA = 1

    Dim i As Long
    For i = 2 To 4
        If RelativeClose(lengths(i), sideA, RECTANGLE_EDGE_REL_TOL) Then
            countA = countA + 1
        ElseIf sideB <= GEOM_TOLERANCE Then
            sideB = lengths(i)
            countB = 1
        ElseIf RelativeClose(lengths(i), sideB, RECTANGLE_EDGE_REL_TOL) Then
            countB = countB + 1
        Else
            Exit Function
        End If
    Next i

    If sideB <= GEOM_TOLERANCE Then
        If countA <> 4 Then Exit Function
        width = sideA
        height = sideA
        widthIndex = 1
    Else
        If countA <> 2 Or countB <> 2 Then Exit Function
        If sideA >= sideB Then
            width = sideA
            height = sideB
        Else
            width = sideB
            height = sideA
        End If
        widthIndex = FindEdgeIndexByLength(lengths, width)
    End If

    If width <= GEOM_TOLERANCE Or height <= GEOM_TOLERANCE Then Exit Function
    If Abs(width * height - area) / area > RECTANGLE_BOUNDS_AREA_REL_TOL Then Exit Function
    If Not HasOrthogonalEdge(ux, uy, widthIndex) Then Exit Function

    rotation = Atn2Local(uy(widthIndex), ux(widthIndex))
    ResolveRectangleFromEdges = True
End Function

' Добавляет вершину, если такой точки еще нет среди концов ребер.
Private Sub AddUniquePoint(ByRef pointX() As Double, ByRef pointY() As Double, _
        ByRef pointCount As Long, ByVal xCoord As Double, ByVal yCoord As Double)
    Dim i As Long
    For i = 1 To pointCount
        If Abs(pointX(i) - xCoord) <= GEOM_TOLERANCE And _
                Abs(pointY(i) - yCoord) <= GEOM_TOLERANCE Then Exit Sub
    Next i

    pointCount = pointCount + 1
    If pointCount <= UBound(pointX) Then
        pointX(pointCount) = xCoord
        pointY(pointCount) = yCoord
    End If
End Sub

' Ищет ребро с нужной длиной, чтобы взять по нему направление стороны Width.
Private Function FindEdgeIndexByLength(ByRef lengths() As Double, ByVal targetLength As Double) As Long
    Dim i As Long
    For i = 1 To 4
        If RelativeClose(lengths(i), targetLength, RECTANGLE_EDGE_REL_TOL) Then
            FindEdgeIndexByLength = i
            Exit Function
        End If
    Next i
    FindEdgeIndexByLength = 1
End Function

' Проверяет, что среди четырех ребер есть направление, ортогональное выбранной
' стороне. Это не дает принять произвольный четырехугольник за прямоугольник.
Private Function HasOrthogonalEdge(ByRef ux() As Double, ByRef uy() As Double, _
        ByVal referenceIndex As Long) As Boolean
    Dim i As Long
    For i = 1 To 4
        If i <> referenceIndex Then
            If Abs(ux(referenceIndex) * ux(i) + uy(referenceIndex) * uy(i)) <= RECTANGLE_ORTHOGONAL_TOL Then
                HasOrthogonalEdge = True
                Exit Function
            End If
        End If
    Next i
End Function

' Проверяет относительное равенство длин с учетом абсолютного геометрического
' допуска для очень маленьких элементов.
Private Function RelativeClose(ByVal actual As Double, ByVal expected As Double, _
        ByVal relativeTolerance As Double) As Boolean
    Dim scaleValue As Double
    scaleValue = Abs(expected)
    If scaleValue < GEOM_TOLERANCE Then scaleValue = 1#
    RelativeClose = (Abs(actual - expected) <= GeomMax(GEOM_TOLERANCE, scaleValue * relativeTolerance))
End Function

Private Function IsLineEntity(ByVal entity As Object) As Boolean
    On Error Resume Next
    IsLineEntity = (StrComp(CStr(entity.ObjectName), "AcDbLine", vbTextCompare) = 0)
    On Error GoTo 0
End Function

' Считает собственные инерции прямоугольного бетонного элемента в глобальных
' осях. Это нужно, когда реальные Width/Height/Rotation найдены по контуру
' Region, но AutoCAD не отдал пригодные MomentOfInertia/ProductOfInertia.
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

' Удаляет временные объекты AutoCAD, созданные при разложении копии Region.
Private Sub DeleteExplodedPiecesQuietly(ByVal pieces As Variant)
    On Error Resume Next
    If IsArray(pieces) Then
        Dim i As Long
        For i = LBound(pieces) To UBound(pieces)
            DeleteAcadObjectQuietly pieces(i)
        Next i
    End If
    On Error GoTo 0
End Sub

Private Sub DeleteAcadObjectQuietly(ByVal entity As Object)
    On Error Resume Next
    If Not entity Is Nothing Then entity.Delete
    On Error GoTo 0
End Sub

' Проверяет, что AutoCAD Bounds можно считать реальными сторонами
' прямоугольного бетонного элемента, а не только внешней рамкой произвольной
' области. Именно такие размеры должны храниться в CSectionModel как реальные.
Private Function IsReliableRectangleBounds(ByVal area As Double, ByVal width As Double, _
        ByVal height As Double) As Boolean
    If area <= GEOM_TOLERANCE Or width <= GEOM_TOLERANCE Or height <= GEOM_TOLERANCE Then Exit Function

    Dim boundsArea As Double
    boundsArea = width * height
    If boundsArea <= GEOM_TOLERANCE Then Exit Function

    IsReliableRectangleBounds = (Abs(boundsArea - area) / area <= RECTANGLE_BOUNDS_AREA_REL_TOL)
End Function

' Матрица собственных инерций должна быть положительно определенной, иначе
' эквивалентный прямоугольник по A/I физически не восстанавливается надежно.
Private Function IsPositiveDefiniteInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double) As Boolean
    If localIx <= GEOM_TOLERANCE Or localIy <= GEOM_TOLERANCE Then Exit Function
    IsPositiveDefiniteInertia = (localIx * localIy - localIxy * localIxy > GEOM_TOLERANCE)
End Function

' Последний fallback для импортированного Region: размеры элемента в модели
' остаются пустыми. Для границ, инерций и визуализации принимается квадрат той
' же площади; после чтения всей сетки ему назначается средний угол
' распознанных прямоугольных элементов.
Private Sub ApplyEquivalentSquareFallback(ByVal area As Double, _
        ByRef width As Double, ByRef height As Double, ByRef rotation As Double, _
        ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double, _
        ByRef shapeType As String, ByRef commentText As String)
    width = 0#
    height = 0#
    rotation = 0#
    localIx = area * area / 12#
    localIy = localIx
    localIxy = 0#
    shapeType = "EquivalentSquare"
    commentText = "AutoCAD Region: реальные размеры и надежные Ix/Iy не найдены, используется квадрат той же площади."
End Sub

Private Function IsRegionEntity(ByVal entity As Object) As Boolean
    On Error Resume Next
    IsRegionEntity = (StrComp(CStr(entity.ObjectName), "AcDbRegion", vbTextCompare) = 0)
    On Error GoTo 0
End Function

Private Function SameLayer(ByVal actualLayer As String, ByVal expectedLayer As String) As Boolean
    SameLayer = (StrComp(Trim$(actualLayer), Trim$(expectedLayer), vbTextCompare) = 0)
End Function

Private Function EntityHandle(ByVal entity As Object) As String
    On Error Resume Next
    EntityHandle = CStr(entity.Handle)
    On Error GoTo 0
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateImportedModel(ByVal model As CSectionModel)
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4405, "CAutoCADSectionModelImporter", _
        "Не найдено бетонных AutoCAD Region на заданном слое импорта."
    If model.RebarCount <= 0 Then Err.Raise vbObjectError + 4406, "CAutoCADSectionModelImporter", _
        "Не найдено арматурных AutoCAD Region на заданном слое импорта."
End Sub

Private Function HasArrayRows(ByVal values As Variant) As Boolean
    On Error GoTo NoRows
    Dim firstRow As Long
    firstRow = LBound(values, 1)
    HasArrayRows = (UBound(values, 1) >= firstRow)
    Exit Function
NoRows:
    HasArrayRows = False
End Function

Private Function OptionalRegionDouble(ByVal values As Variant, ByVal rowIndex As Long, _
        ByVal colIndex As Long) As Double
    On Error GoTo MissingValue
    If colIndex > UBound(values, 2) Then GoTo MissingValue
    If Len(Trim$(CStr(values(rowIndex, colIndex)))) = 0 Then GoTo MissingValue
    OptionalRegionDouble = CDbl(values(rowIndex, colIndex))
    Exit Function
MissingValue:
    OptionalRegionDouble = 0#
End Function

Private Function OptionalRegionString(ByVal values As Variant, ByVal rowIndex As Long, _
        ByVal colIndex As Long) As String
    On Error GoTo MissingValue
    If colIndex > UBound(values, 2) Then GoTo MissingValue
    OptionalRegionString = CStr(values(rowIndex, colIndex))
    Exit Function
MissingValue:
    OptionalRegionString = vbNullString
End Function

Private Function InputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputLength = value Else InputLength = units.InputLengthToInternal(value)
End Function

Private Function InputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputArea = value Else InputArea = units.InputAreaToInternal(value)
End Function

Private Function InputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputFourthPowerLength = value Else InputFourthPowerLength = units.InputFourthPowerLengthToInternal(value)
End Function

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
' Тесты собирают AutoCAD-import модель из массивов, чтобы проверить фильтрацию,
' заполнение CSectionModel и расчетные поля без запуска AutoCAD COM.

Public Function BuildFromRegionArrays(ByVal concreteRegions As Variant, ByVal rebarRegions As Variant, _
        ByVal steelClass As String, Optional ByVal minArea As Double = 0.000001) As CSectionModel
    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    AddConcreteRegionArray model, concreteRegions, minArea
    AddRebarRegionArray model, rebarRegions, steelClass, minArea
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    ValidateImportedModel model

    Set BuildFromRegionArrays = model
End Function

