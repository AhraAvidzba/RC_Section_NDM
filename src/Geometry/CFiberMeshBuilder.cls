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
' добавляет целиком, а пограничные уточняет последовательно на заданное
' число уровней Mesh.BoundarySubdivisions.
' Каждый уровень делит пересечённую ячейку пополам по обеим осям;
' внутренние ячейки остаются крупными. При autoStepFactor > 0 одинаковый
' шаг по X/Y равен этой доле меньшего точного габарита области в WCS.
Public Sub BuildMesh(ByVal geometry As CGeometryRegion, ByVal stepX As Double, ByVal stepY As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal boundarySubdivisions As Long = 1, _
        Optional ByVal sourceReference As String = vbNullString, Optional ByVal autoStepFactor As Double = 0#)

    On Error GoTo FailedBuild
    If mReadOnly Then Err.Raise vbObjectError + 2111, "CFiberMeshBuilder", "Опубликованную сетку нельзя перестроить. Создайте новый построитель."
    mFiberCount = 0: mBuildSeconds = 0#: Erase mFibers ' Ошибочная новая попытка не должна оставлять доступной сетку прежнего запуска.
    Dim message As String
    If geometry Is Nothing Then Err.Raise vbObjectError + 2107, "CFiberMeshBuilder", "Геометрия бетонного сечения не передана."
    If Not geometry.IsValid(message) Then Err.Raise vbObjectError + 2100, "CFiberMeshBuilder", message
    If autoStepFactor < 0# Then Err.Raise vbObjectError + 2101, "CFiberMeshBuilder", "Коэффициент автоматического шага не может быть отрицательным."
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
    For level = 1 To boundarySubdivisions
        If upperCount > 2147483646# / 4# Then
            Err.Raise vbObjectError + 2108, "CFiberMeshBuilder", _
                "Верхняя оценка числа бетонных элементов с учетом уровней Mesh.BoundarySubdivisions не помещается в Long. " & _
                "Увеличьте шаг сетки или уменьшите число уровней уточнения границы."
        End If
        upperCount = upperCount * 4# ' Проверяем до умножения, не вычисляя переполняющуюся степень 4^N.
    Next level

    ReDim mFibers(1 To nx * ny)
    Dim ix As Long, iy As Long, cellMinX As Double, cellMinY As Double
    For iy = 0 To ny - 1
        cellMinY = minY + iy * stepY
        For ix = 0 To nx - 1
            cellMinX = minX + ix * stepX
            AddAdaptiveCell geometry, cellMinX, cellMinY, stepX, stepY, boundarySubdivisions, materialID
        Next ix
    Next iy

    If mFiberCount <= 0 Then
        Err.Raise vbObjectError + 2105, "CFiberMeshBuilder", _
            "Сетка бетонных элементов пуста: minX=" & CStr(minX) & "; maxX=" & CStr(maxX) & _
            "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY) & _
            "; Mesh.StepX=" & CStr(stepX) & "; Mesh.StepY=" & CStr(stepY) & "."
    End If

    mBuildSeconds = Timer - t0
    If mBuildSeconds < 0# Then mBuildSeconds = mBuildSeconds + 86400# ' Timer обнуляется в полночь; длительность успешной сборки остаётся положительной.
    Exit Sub
FailedBuild:
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    If Not mReadOnly Then
        mFiberCount = 0: mBuildSeconds = 0#: Erase mFibers ' Не публикуем частично собранные элементы; отказ изменения Frozen-сетки её не очищает.
    End If
    If Len(sourceReference) > 0 Then reason = reason & " Параметры бетонной сетки: " & sourceReference & "."
    Err.Raise number, source, reason
End Sub


' Уточняет только ячейку, через открытую внутренность которой проходит точная
' граница бетона или отверстия. Без пересечения принадлежность центра определяет
' всю ячейку; на последнем уровне центр остаётся квадратурной аппроксимацией.
Private Sub AddAdaptiveCell(ByVal geometry As CGeometryRegion, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal remainingLevels As Long, ByVal materialID As Long)
    Dim x As Double, y As Double
    x = cellMinX + stepX / 2#: y = cellMinY + stepY / 2#
    If remainingLevels > 0 Then
        If geometry.MeshBoundaryCrossesCell(cellMinX, cellMinY, cellMinX + stepX, cellMinY + stepY) Then
            Dim halfX As Double, halfY As Double, ix As Long, iy As Long
            halfX = stepX / 2#: halfY = stepY / 2#
            For iy = 0 To 1
                For ix = 0 To 1
                    AddAdaptiveCell geometry, cellMinX + ix * halfX, cellMinY + iy * halfY, halfX, halfY, remainingLevels - 1, materialID
                Next ix
            Next iy
            Exit Sub
        End If
    End If
    If geometry.ContainsPoint(x, y) Then AddFiber x, y, stepX, stepY, materialID
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

' -------------------- ДЛЯ ТЕСТОВ --------------------
' возвращает время построения сетки для проверки скорости.
' Рабочий расчёт не использует это время при выборе результата.
Public Property Get TestBuildSeconds() As Double
    TestBuildSeconds = mBuildSeconds
End Property
