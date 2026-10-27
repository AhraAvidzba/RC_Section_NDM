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
' Разбивает бетонную область ячейками заданного шага: адаптивными
' прямоугольниками либо точными пересечениями с контуром. Подготавливает
' центры, площади, инерции и необязательные границы для CSectionModel.
' Класс отвечает за геометрию сетки; нагрузки и материалы обрабатываются отдельно.

Private mFibers() As TFiber ' Единственный массив готовых бетонных элементов: координаты, размеры, площади и материал.
Private mFiberGeometry() As CConcreteRegion ' Фактические области Contour; для компактных прямоугольников Adaptive хранится Nothing.
Private mReadOnly As Boolean ' Готовая сетка общей сборки только читается, без копирования массива.
Private mFiberCount As Long ' Число заполненных элементов; не размер зарезервированного массива.
Private mBuildSeconds As Double ' Время создания только сетки; не полное время расчёта сочетания.

' Разбивает бетонную область регулярной сеткой. Полностью внутренние ячейки
' добавляет целиком. Adaptive уточняет пограничные ячейки последовательно
' на Mesh.BoundarySubdivisions уровней; Contour сохраняет точное пересечение.
' Каждый уровень делит пересечённую ячейку пополам по обеим осям;
' внутренние ячейки остаются крупными. При autoStepFactor > 0 одинаковый
' шаг по X/Y равен этой доле меньшего точного габарита области в WCS.
' minArea задаётся в мм2; каждая конечная часть меньшей площади пропускается.
' Равенство порогу сохраняет элемент, а ноль отключает фильтр.
Public Sub BuildMesh(ByVal geometry As CGeometryRegion, ByVal stepX As Double, ByVal stepY As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal boundarySubdivisions As Long = 1, _
        Optional ByVal sourceReference As String = vbNullString, Optional ByVal autoStepFactor As Double = 0#, _
        Optional ByVal cellGeometry As String = "Adaptive", Optional ByVal minArea As Double = 0#)

    On Error GoTo FailedBuild
    If mReadOnly Then Err.Raise vbObjectError + 2111, "CFiberMeshBuilder", "Опубликованную сетку нельзя перестроить. Создайте новый построитель."
    mFiberCount = 0: mBuildSeconds = 0#: Erase mFibers: Erase mFiberGeometry ' Попытка построения не публикует остатки другой сетки или её границы.
    If Len(cellGeometry) = 0 Then cellGeometry = "Adaptive" ' Незаполненное программное описание использует адаптивную сетку; пользовательский ввод проверяется до ядра.
    If cellGeometry <> "Adaptive" And cellGeometry <> "Contour" Then Err.Raise vbObjectError + 2109, "CFiberMeshBuilder", "Неизвестный способ геометрии ячеек Mesh.CellGeometry."
    If minArea < 0# Then Err.Raise vbObjectError + 2112, "CFiberMeshBuilder", "Mesh.MinElementArea не может быть отрицательной."
    Dim message As String
    If geometry Is Nothing Then Err.Raise vbObjectError + 2107, "CFiberMeshBuilder", "Геометрия бетонного сечения не передана."
    If Not geometry.IsValid(message) Then Err.Raise vbObjectError + 2100, "CFiberMeshBuilder", message
    If autoStepFactor < 0# Then Err.Raise vbObjectError + 2101, "CFiberMeshBuilder", "Коэффициент автоматического шага не может быть отрицательным."
    If cellGeometry = "Adaptive" And boundarySubdivisions < 1 Then Err.Raise vbObjectError + 2109, "CFiberMeshBuilder", _
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

    If autoStepFactor > 0# Then
        stepX = (maxX - minX) * autoStepFactor
        If maxY - minY < maxX - minX Then stepX = (maxY - minY) * autoStepFactor ' Размеры берутся по всей области, включая смещения и поворот; отверстия не становятся отдельным габаритом.
        stepY = stepX
    End If
    If stepX <= 0# Then Err.Raise vbObjectError + 2101, "CFiberMeshBuilder", "Mesh.StepX должен быть положительным."
    If stepY <= 0# Then Err.Raise vbObjectError + 2102, "CFiberMeshBuilder", "Mesh.StepY должен быть положительным."

    Dim nx As Long
    Dim ny As Long
    If stepX <= (maxX - minX) / 2147483646# Or stepY <= (maxY - minY) / 2147483646# Then ' Проверяем размер до деления на сверхмалый шаг и преобразования в Long. Это ограничение представимости массива, а не изменение точности сетки.
        Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
            "Шаг бетонной сетки слишком мал: число ячеек по одной оси превышает предел счётчика программы (2147483647). Увеличьте Mesh.StepX или Mesh.StepY."
    End If
    nx = CLng(Fix((maxX - minX) / stepX)) + 1
    ny = CLng(Fix((maxY - minY) / stepY)) + 1
    If nx < 1 Or ny < 1 Then Err.Raise vbObjectError + 2104, "CFiberMeshBuilder", "Габаритный прямоугольник сетки пуст."
    Dim upperCount As Double, level As Long
    upperCount = CDbl(nx) * CDbl(ny)
    If upperCount > 2147483646# Then Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", "Число базовых ячеек бетонной сетки не помещается в Long. Увеличьте шаг сетки."
    Dim subdivisionLevels As Long
    If cellGeometry = "Adaptive" Then subdivisionLevels = boundarySubdivisions
    For level = 1 To subdivisionLevels
        If upperCount > 2147483646# / 4# Then
            Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
                "Верхняя оценка числа бетонных элементов с учетом уровней Mesh.BoundarySubdivisions не помещается в Long. " & _
                "Увеличьте шаг сетки или уменьшите число уровней уточнения границы."
        End If
        upperCount = upperCount * 4# ' Проверяем до умножения, не вычисляя переполняющуюся степень 4^N.
    Next level

    ReDim mFibers(1 To nx * ny)
    ReDim mFiberGeometry(1 To nx * ny)
    Dim ix As Long, iy As Long, cellMinX As Double, cellMinY As Double, query As CSectionGeometryQuery
    If cellGeometry = "Contour" Then Set query = New CSectionGeometryQuery ' Один запрос обслуживает все точные ячейки; у Adaptive объект не создаётся.
    For iy = 0 To ny - 1
        cellMinY = minY + iy * stepY
        For ix = 0 To nx - 1
            cellMinX = minX + ix * stepX
            If cellGeometry = "Contour" Then
                AddContourCell geometry, cellMinX, cellMinY, stepX, stepY, materialID, minArea, query
            Else
                AddAdaptiveCell geometry, cellMinX, cellMinY, stepX, stepY, boundarySubdivisions, materialID, minArea
            End If
        Next ix
    Next iy

    If mFiberCount <= 0 Then
        Err.Raise vbObjectError + 2105, "CFiberMeshBuilder", _
            "Сетка бетонных элементов пуста: minX=" & CStr(minX) & "; maxX=" & CStr(maxX) & _
            "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY) & _
            "; Mesh.StepX=" & CStr(stepX) & "; Mesh.StepY=" & CStr(stepY) & _
            "; Mesh.MinElementArea=" & CStr(minArea) & " мм2."
    End If

    mBuildSeconds = Timer - t0
    If mBuildSeconds < 0# Then mBuildSeconds = mBuildSeconds + 86400# ' Timer обнуляется в полночь; длительность успешной сборки остаётся положительной.
    Exit Sub
FailedBuild:
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    If Not mReadOnly Then
        mFiberCount = 0: mBuildSeconds = 0#: Erase mFibers: Erase mFiberGeometry ' Не публикуем частичную сетку; отказ изменения Frozen-сетки её не очищает.
    End If
    If Len(sourceReference) > 0 Then reason = reason & " Параметры бетонной сетки: " & sourceReference & "."
    Err.Raise number, source, reason
End Sub


' Уточняет только ячейку, через открытую внутренность которой проходит точная
' граница бетона или отверстия. Без пересечения принадлежность центра определяет
' всю ячейку; на последнем уровне центр остаётся квадратурной аппроксимацией.
Private Sub AddAdaptiveCell(ByVal geometry As CGeometryRegion, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal remainingLevels As Long, ByVal materialID As Long, ByVal minArea As Double)
    If stepX * stepY < minArea Then Exit Sub ' Потомки только уменьшаются: если этот прямоугольник меньше порога, ни одна конечная часть не будет принята.
    Dim x As Double, y As Double
    x = cellMinX + stepX / 2#: y = cellMinY + stepY / 2#
    If remainingLevels > 0 Then
        If geometry.MeshBoundaryCrossesCell(cellMinX, cellMinY, cellMinX + stepX, cellMinY + stepY) Then
            Dim halfX As Double, halfY As Double, ix As Long, iy As Long
            halfX = stepX / 2#: halfY = stepY / 2#
            For iy = 0 To 1
                For ix = 0 To 1
                    AddAdaptiveCell geometry, cellMinX + ix * halfX, cellMinY + iy * halfY, halfX, halfY, remainingLevels - 1, materialID, minArea
                Next ix
            Next iy
            Exit Sub
        End If
    End If
    If geometry.ContainsPoint(x, y) Then AddFiber x, y, stepX, stepY, materialID
End Sub

' Пересекает только пограничную базовую ячейку с точной бетонной областью.
' Несвязные части становятся отдельными элементами; принадлежность центра
' не используется как замена пересечению. Внутренний прямоугольник тоже имеет точную область.
Private Sub AddContourCell(ByVal geometry As CGeometryRegion, ByVal minX As Double, ByVal minY As Double, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal materialID As Long, ByVal minArea As Double, ByVal query As CSectionGeometryQuery)
    Dim pieces As Collection, region As CConcreteRegion
    If geometry.MeshBoundaryCrossesCell(minX, minY, minX + stepX, minY + stepY) Then
        Set pieces = geometry.ClipCell(minX, minY, minX + stepX, minY + stepY)
        For Each region In pieces
            AddContourRegion query, region, stepX, stepY, materialID, minArea
        Next region
    Else
        If geometry.ContainsPoint(minX + stepX / 2#, minY + stepY / 2#) Then
            Set region = query.RectangleRegion(minX + stepX / 2#, minY + stepY / 2#, stepX, stepY)
            AddContourRegion query, region, stepX, stepY, materialID, minArea
        End If
    End If
End Sub

' Публикует одну связную точную часть только после проверки её собственной
' площади. Площадь ограничивающей ячейки не заменяет площадь части с отверстиями.
Private Sub AddContourRegion(ByVal query As CSectionGeometryQuery, ByVal region As CConcreteRegion, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal materialID As Long, ByVal minArea As Double)
    Dim properties As TGeometryProperties
    query.RegionProperties region, properties
    If properties.Area < minArea Then Exit Sub
    AddFiber properties.CentroidX, properties.CentroidY, stepX, stepY, materialID
    mFibers(mFiberCount).Area = properties.Area
    mFibers(mFiberCount).LocalIx = properties.Ixc: mFibers(mFiberCount).LocalIy = properties.Iyc: mFibers(mFiberCount).LocalIxy = properties.Ixyc
    Set mFiberGeometry(mFiberCount) = region
End Sub

' Добавляет ячейку с прямоугольными A/I и размерами опорного прямоугольника.
' Для Adaptive это её фактическая форма; AddContourRegion заменяет A/c/I
' точными свойствами и сохраняет самостоятельную границу части Contour.
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
        ReDim Preserve mFiberGeometry(1 To newCapacity)
    End If
    mFibers(mFiberCount).X = x
    mFibers(mFiberCount).Y = y
    mFibers(mFiberCount).Width = width
    mFibers(mFiberCount).Height = height
    mFibers(mFiberCount).Area = width * height
    mFibers(mFiberCount).FillFactor = 1#
    mFibers(mFiberCount).MaterialID = materialID
    mFibers(mFiberCount).LocalIx = mFibers(mFiberCount).Area * height * height / 12#
    mFibers(mFiberCount).LocalIy = mFibers(mFiberCount).Area * width * width / 12#
    mFibers(mFiberCount).LocalIxy = 0#
End Sub

' Возвращает независимую копию области одного точного элемента; Nothing
' обозначает компактный прямоугольник. Массив сетки при чтении не копируется.
Public Function FiberGeometry(ByVal index As Long) As CConcreteRegion
    CheckFiberIndex index
    If Not mFiberGeometry(index) Is Nothing Then Set FiberGeometry = mFiberGeometry(index).Clone
End Function

' Возвращает собственные инерции уже подготовленной ячейки без повторного интегрирования границы.
Public Sub FiberLocalInertia(ByVal index As Long, ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    CheckFiberIndex index
    localIx = mFibers(index).LocalIx: localIy = mFibers(index).LocalIy: localIxy = mFibers(index).LocalIxy
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

' Возвращает ширину опорного прямоугольника ячейки, мм; точная форма читается через FiberGeometry.
Public Property Get FiberWidth(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberWidth = mFibers(index).Width
End Property

' Возвращает высоту опорного прямоугольника ячейки, мм; у Adaptive он совпадает с границей элемента.
Public Property Get FiberHeight(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberHeight = mFibers(index).Height
End Property

' Возвращает фактическую площадь ячейки; у компактного прямоугольника это Width * Height, мм2.
Public Property Get FiberArea(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberArea = mFibers(index).Area
End Property

' Возвращает коэффициент заполнения расчётной площади: готовая ячейка уже содержит только бетон и имеет коэффициент 1.
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает время построения сетки для проверки скорости.
' Рабочий расчёт не использует это время при выборе результата.
Public Property Get TestBuildSeconds() As Double
    TestBuildSeconds = mBuildSeconds
End Property
