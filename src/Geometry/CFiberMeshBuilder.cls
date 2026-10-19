VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFiberMeshBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Построение бетонной расчётной сетки
' ==========================================================================
' Разбивает бетонную область на прямоугольные элементы заданного шага.
' Подготавливает их центры, размеры и площади для переноса в CSectionModel.
' Класс отвечает за геометрию сетки; нагрузки и материалы обрабатываются отдельно.

Private mFibers() As TFiber ' Единственный массив готовых бетонных элементов: координаты, размеры, площади и материал.
Private mReadOnly As Boolean ' Готовая сетка общей сборки только читается, без копирования массива.
Private mFiberCount As Long ' Число заполненных элементов; не размер зарезервированного массива.
Private mBuildSeconds As Double ' Время создания только сетки; не полное время расчёта сочетания.

' Разбивает бетонную область регулярной сеткой. Полностью внутренние ячейки
' добавляет целиком, а пограничные при необходимости делит на более мелкие
' по настройке Mesh.BoundarySubdivisions.
Public Sub BuildMesh(ByVal geometry As CGeometryRegion, ByVal stepX As Double, ByVal stepY As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal boundarySubdivisions As Long = 1, _
        Optional ByVal sourceReference As String = vbNullString)

    On Error GoTo FailedBuild
    If mReadOnly Then Err.Raise vbObjectError + 2111, "CFiberMeshBuilder", "Опубликованную сетку нельзя перестроить. Создайте новый построитель."
    Dim message As String
    If geometry Is Nothing Then Err.Raise vbObjectError + 2107, "CFiberMeshBuilder", "Геометрия бетонного сечения не передана."
    If Not geometry.IsValid(message) Then Err.Raise vbObjectError + 2100, "CFiberMeshBuilder", message
    If stepX <= 0# Then Err.Raise vbObjectError + 2101, "CFiberMeshBuilder", "Mesh.StepX должен быть положительным."
    If stepY <= 0# Then Err.Raise vbObjectError + 2102, "CFiberMeshBuilder", "Mesh.StepY должен быть положительным."
    If boundarySubdivisions < 1 Then Err.Raise vbObjectError + 2109, "CFiberMeshBuilder", _
        "Mesh.BoundarySubdivisions должен быть целым числом не меньше 1."

    Dim t0 As Double
    t0 = Timer

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    minX = geometry.MinX
    maxX = geometry.MaxX
    minY = geometry.MinY
    maxY = geometry.MaxY

    If maxX <= minX + GEOM_TOLERANCE Or maxY <= minY + GEOM_TOLERANCE Then
        Err.Raise vbObjectError + 2103, "CFiberMeshBuilder", _
            "Габариты бетонной геометрии вырождены: minX=" & CStr(minX) & _
            "; maxX=" & CStr(maxX) & "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY) & "."
    End If

    Dim nx As Long
    Dim ny As Long
    ' Проверяем размер до деления на сверхмалый шаг и преобразования в Long.
    ' Это ограничение представимости массива, а не изменение точности сетки.
    If stepX <= (maxX - minX) / 2147483646# Or stepY <= (maxY - minY) / 2147483646# Then
        Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
            "Шаг бетонной сетки слишком мал: число ячеек по одной оси превышает предел счётчика программы (2147483647). Увеличьте Mesh.StepX или Mesh.StepY."
    End If
    nx = CLng(Fix((maxX - minX) / stepX)) + 1
    ny = CLng(Fix((maxY - minY) / stepY)) + 1
    If nx < 1 Or ny < 1 Then Err.Raise vbObjectError + 2104, "CFiberMeshBuilder", "Габаритный прямоугольник сетки пуст."
    If CDbl(nx) * CDbl(ny) * CDbl(boundarySubdivisions) * CDbl(boundarySubdivisions) > 2147483646# Then
        Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
            "Верхняя оценка числа бетонных элементов с учетом Mesh.BoundarySubdivisions не помещается в Long. " & _
            "Увеличьте шаг сетки или уменьшите число подъячеек границы."
    End If

    ReDim mFibers(1 To nx * ny)
    mFiberCount = 0

    Dim ix As Long
    Dim iy As Long
    Dim x As Double
    Dim y As Double
    Dim cellMinX As Double
    Dim cellMinY As Double
    Dim allCornersInside As Boolean
    Dim anyCornerInside As Boolean
    Dim openingBounds() As Double, openingCount As Long, openingIndex As Long
    Dim contours As CSectionContours, query As CSectionGeometryQuery
    Dim outerEdges() As TRegionEdge, outerEdgeCount As Long
    ' Даже если четыре угла лежат в бетоне, внутри ячейки может быть отверстие.
    ' Один раз получаем габариты всех отверстий и уточняем такие ячейки
    ' проверкой принадлежности точек бетону через ContainsPoint.
    If boundarySubdivisions > 1 Then
        Set contours = New CSectionContours: geometry.BuildContours contours
        Set query = New CSectionGeometryQuery
        query.OpeningBounds contours, openingBounds, openingCount
        query.CopyMeshOuterEdges contours, outerEdges, outerEdgeCount
    End If

    For iy = 0 To ny - 1
        cellMinY = minY + iy * stepY
        y = cellMinY + stepY / 2#
        For ix = 0 To nx - 1
            cellMinX = minX + ix * stepX
            x = cellMinX + stepX / 2#
            ' Центр последней базовой ячейки может быть за габаритом, хотя
            ' часть ее подъячеек находится в бетоне. Отбор выполняет форма
            ' по центрам фактических элементов, не по центру базовой ячейки.
            If boundarySubdivisions = 1 Then
                If geometry.ContainsPoint(x, y) Then AddFiber x, y, stepX, stepY, materialID
            Else
                GetCornerStatus geometry, cellMinX, cellMinY, stepX, stepY, allCornersInside, anyCornerInside
                If allCornersInside Then
                    For openingIndex = 1 To openingCount
                        ' Касание габарита без пересечения площади не требует дробления ячейки.
                        ' Так малая погрешность координат дуги не вызывает лишнее уточнение
                        ' сетки.
                        If cellMinX < openingBounds(openingIndex, 2) - GEOM_TOLERANCE And cellMinX + stepX > openingBounds(openingIndex, 1) + GEOM_TOLERANCE And _
                                cellMinY < openingBounds(openingIndex, 4) - GEOM_TOLERANCE And cellMinY + stepY > openingBounds(openingIndex, 3) + GEOM_TOLERANCE Then
                            allCornersInside = False
                            Exit For
                        End If
                    Next openingIndex
                End If
                If allCornersInside Then
                    If query.MeshBoundaryCrossesCell(outerEdges, outerEdgeCount, cellMinX, cellMinY, cellMinX + stepX, cellMinY + stepY) Then allCornersInside = False
                End If
                If allCornersInside Then
                    AddFiber x, y, stepX, stepY, materialID
                Else
                    AddSubcellFibers geometry, cellMinX, cellMinY, stepX, stepY, boundarySubdivisions, materialID
                End If
            End If
        Next ix
    Next iy

    If mFiberCount <= 0 Then
        Err.Raise vbObjectError + 2105, "CFiberMeshBuilder", _
            "Сетка бетонных элементов пуста: minX=" & CStr(minX) & "; maxX=" & CStr(maxX) & _
            "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY) & _
            "; Mesh.StepX=" & CStr(stepX) & "; Mesh.StepY=" & CStr(stepY) & "."
    End If

    mBuildSeconds = Timer - t0
    Exit Sub
FailedBuild:
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    If Len(sourceReference) > 0 Then reason = reason & " Параметры бетонной сетки: " & sourceReference & "."
    Err.Raise number, source, reason
End Sub

' Дробит пограничную ячейку на более мелкие прямоугольники и добавляет только
' те центры подъячеек, которые попали внутрь исходной геометрии.
Private Sub AddSubcellFibers(ByVal geometry As CGeometryRegion, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal boundarySubdivisions As Long, ByVal materialID As Long)
    Dim subX As Double
    Dim subY As Double
    subX = stepX / CDbl(boundarySubdivisions)
    subY = stepY / CDbl(boundarySubdivisions)

    Dim ix As Long
    Dim iy As Long
    Dim x As Double
    Dim y As Double
    For iy = 0 To boundarySubdivisions - 1
        y = cellMinY + subY / 2# + iy * subY
        For ix = 0 To boundarySubdivisions - 1
            x = cellMinX + subX / 2# + ix * subX
            If geometry.ContainsPoint(x, y) Then AddFiber x, y, subX, subY, materialID
        Next ix
    Next iy
End Sub

' Проверяет четыре угла расчетной ячейки. Это быстрый способ понять, можно ли
' взять всю ячейку целиком или нужно уточнять границу подъячеечным разбиением.
Private Sub GetCornerStatus(ByVal geometry As CGeometryRegion, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal width As Double, ByVal height As Double, ByRef allInside As Boolean, ByRef anyInside As Boolean)
    Dim p1 As Boolean
    Dim p2 As Boolean
    Dim p3 As Boolean
    Dim p4 As Boolean
    p1 = geometry.ContainsPoint(cellMinX, cellMinY)
    p2 = geometry.ContainsPoint(cellMinX + width, cellMinY)
    p3 = geometry.ContainsPoint(cellMinX + width, cellMinY + height)
    p4 = geometry.ContainsPoint(cellMinX, cellMinY + height)
    allInside = p1 And p2 And p3 And p4
    anyInside = p1 Or p2 Or p3 Or p4
End Sub

' Добавляет прямоугольный бетонный элемент в подготовленную сетку.
' Его Width и Height сохраняются для определения границ, расстояний
' до края бетона и отрисовки фактических размеров элемента.
Private Sub AddFiber(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByVal materialID As Long)
    If mFiberCount >= 2147483646 Then Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
        "Число бетонных элементов достигло предела представимости массива. Увеличьте шаг сетки."
    mFiberCount = mFiberCount + 1
    If mFiberCount > UBound(mFibers) Then
        Dim newCapacity As Long
        If UBound(mFibers) > 1073741823 Then
            newCapacity = 2147483646
        Else
            newCapacity = UBound(mFibers) * 2
        End If
        ReDim Preserve mFibers(1 To newCapacity)
    End If
    mFibers(mFiberCount).X = x
    mFibers(mFiberCount).Y = y
    mFibers(mFiberCount).Width = width
    mFibers(mFiberCount).Height = height
    mFibers(mFiberCount).Area = width * height
    mFibers(mFiberCount).FillFactor = 1#
    mFibers(mFiberCount).MaterialID = materialID
End Sub

' Количество подготовленных бетонных ячеек сетки.
Public Property Get FiberCount() As Long
    FiberCount = mFiberCount
End Property

' Возвращает координату X центра бетонного элемента по индексу, мм.
Public Property Get FiberX(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberX = mFibers(index).X
End Property

' Возвращает координату Y центра бетонного элемента по индексу, мм.
Public Property Get FiberY(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberY = mFibers(index).Y
End Property

' Возвращает ширину прямоугольной ячейки бетонного элемента, мм.
Public Property Get FiberWidth(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberWidth = mFibers(index).Width
End Property

' Возвращает высоту прямоугольной ячейки бетонного элемента, мм.
Public Property Get FiberHeight(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberHeight = mFibers(index).Height
End Property

' Возвращает площадь прямоугольной ячейки бетонного элемента: ширина * высота, мм2.
Public Property Get FiberArea(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberArea = mFibers(index).Area
End Property

' Возвращает долю площади прямоугольной ячейки, занятую бетоном.
Public Property Get FiberFillFactor(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberFillFactor = mFibers(index).FillFactor
End Property

' Возвращает идентификатор материала бетонного элемента по индексу.
Public Property Get FiberMaterialID(ByVal index As Long) As Long
    CheckFiberIndex index
    FiberMaterialID = mFibers(index).MaterialID
End Property

' Проверяет, что номер бетонного элемента находится в пределах построенной сетки.
' При выходе за границы выдаёт понятную ошибку.
Private Sub CheckFiberIndex(ByVal index As Long)
    If index < 1 Or index > mFiberCount Then
        Err.Raise vbObjectError + 2106, "CFiberMeshBuilder", "Индекс бетонного элемента вне допустимого диапазона."
    End If
End Sub


' Завершает построение сетки и запрещает её изменение перед выдачей из CSectionBuilder.
' Готовый массив элементов остаётся тем же и не копируется.
Friend Sub Freeze()
    If mFiberCount <= 0 Then Err.Raise vbObjectError + 2111, "CFiberMeshBuilder", "Нет готовой сетки для публикации."
    mReadOnly = True
End Sub

' Показывает, защищена ли готовая сетка от дальнейших изменений.
Public Property Get IsReadOnly() As Boolean
    IsReadOnly = mReadOnly
End Property

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================

' ДЛЯ ТЕСТОВ: возвращает время построения сетки для проверки скорости.
' Рабочий расчёт не использует это время при выборе результата.
Public Property Get BuildSeconds() As Double
    BuildSeconds = mBuildSeconds
End Property
