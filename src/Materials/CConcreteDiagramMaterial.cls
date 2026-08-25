VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConcreteDiagramMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Кусочно-линейная диаграмма бетона
' ==========================================================================
' Класс является легким расчетным объектом для solver-а: он получает уже
' готовые точки диаграммы "деформация - напряжение" во внутренних единицах
' и быстро возвращает напряжение, касательный и секущий модуль.
'
' Здесь нет выбора класса бетона, ГПС, TwoLine/ThreeLine или режима работы
' растянутого бетона. Эти решения принимает CMaterialModelProvider через
' builders, а сюда попадает только итоговая диаграмма для конкретного
' расчетного назначения: Strength, Mcrc или CrackedNDS.
'
' Если растянутая ветвь не включена в набор точек, крайняя точка диаграммы
' остается (0; 0), поэтому положительная деформация автоматически дает
' нулевое напряжение без отдельной настройки в самом материале.

Private mStrains() As Double ' Узлы диаграммы по деформации, отсортированные от сжатия к растяжению.
Private mStresses() As Double ' Напряжения в узлах диаграммы во внутренних единицах, МПа.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова GetStress до задания диаграммы.

' ------------------------------
' Инициализация
' ------------------------------
' Загружает точки, подготовленные builder-ом материала. Точки уже должны быть
' во внутренних единицах проекта: деформация безразмерная, напряжение в МПа.
Public Sub InitializeFromPointSet(ByVal points As CMaterialPointSet)
    If points Is Nothing Then Err.Raise vbObjectError + 3106, "CConcreteDiagramMaterial", "Точки диаграммы бетона не переданы."
    If points.Count < 2 Then Err.Raise vbObjectError + 3106, "CConcreteDiagramMaterial", "Диаграмма бетона должна содержать не менее двух точек."

    ReDim mStrains(1 To points.Count)
    ReDim mStresses(1 To points.Count)

    Dim i As Long
    For i = 1 To points.Count
        mStrains(i) = points.Strain(i)
        mStresses(i) = points.Stress(i)
    Next i

    SortPointsByStrain
    ValidatePointOrder
    mPointCount = points.Count
    mInitialized = True
End Sub

' Тестовый и служебный конструктор двухлинейной сжатой ветви.
' Нулевая точка добавляется автоматически, чтобы диаграмма замыкалась в начале.
Public Sub InitializeFromTwoCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet

    If eps1 >= 0# Then Err.Raise vbObjectError + 3100, "CConcreteDiagramMaterial", "Первая деформация сжатия бетона должна быть отрицательной."
    If eps2 >= eps1 Then Err.Raise vbObjectError + 3101, "CConcreteDiagramMaterial", "Вторая деформация сжатия бетона должна быть более отрицательной, чем первая."
    If stress1 >= 0# Or stress2 >= 0# Then Err.Raise vbObjectError + 3102, "CConcreteDiagramMaterial", "Напряжения сжатия бетона должны быть отрицательными."

    points.AddPoint eps2, stress2
    points.AddPoint eps1, stress1
    points.AddPoint 0#, 0#
    InitializeFromPointSet points
End Sub

' Короткий старый алиас для тестов, которые создают простой материал вручную.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    InitializeFromTwoCompressionPoints eps1, stress1, eps2, stress2
End Sub

' Тестовый и служебный конструктор трехточечной сжатой ветви.
Public Sub InitializeFromThreeCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet

    If eps1 >= 0# Then Err.Raise vbObjectError + 3100, "CConcreteDiagramMaterial", "Первая деформация сжатия бетона должна быть отрицательной."
    If eps2 >= eps1 Then Err.Raise vbObjectError + 3101, "CConcreteDiagramMaterial", "Вторая деформация сжатия бетона должна быть более отрицательной, чем первая."
    If eps3 >= eps2 Then Err.Raise vbObjectError + 3104, "CConcreteDiagramMaterial", "Третья деформация сжатия бетона должна быть более отрицательной, чем вторая."
    If stress1 >= 0# Or stress2 >= 0# Or stress3 >= 0# Then Err.Raise vbObjectError + 3102, "CConcreteDiagramMaterial", "Напряжения сжатия бетона должны быть отрицательными."

    points.AddPoint eps3, stress3
    points.AddPoint eps2, stress2
    points.AddPoint eps1, stress1
    points.AddPoint 0#, 0#
    InitializeFromPointSet points
End Sub

' Возвращает наиболее сжатую точку диаграммы как ориентир предельной деформации.
Public Property Get UltimateCompressionStrain() As Double
    EnsureInitialized
    UltimateCompressionStrain = mStrains(1)
End Property

' Возвращает крайнюю растянутую точку диаграммы. Если растяжение отключено,
' последней точкой остается ноль и растянутый бетон не ограничивает расчет.
Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    If mStrains(mPointCount) > 0# Then UltimateTensionStrain = mStrains(mPointCount)
End Property

' Количество точек используется в тестах и диагностике построителя диаграмм.
Public Property Get PointCount() As Long
    PointCount = mPointCount
End Property

' Возвращает деформацию узла диаграммы по индексу.
Public Function PointStrain(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3108, "CConcreteDiagramMaterial", "Индекс точки диаграммы бетона вне диапазона."
    PointStrain = mStrains(index)
End Function

' Возвращает напряжение узла диаграммы по индексу.
Public Function PointStress(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3109, "CConcreteDiagramMaterial", "Индекс точки диаграммы бетона вне диапазона."
    PointStress = mStresses(index)
End Function

' ------------------------------
' Расчетный контракт материала
' ------------------------------
' Возвращает напряжение бетона по текущей деформации. За пределами заданной
' диаграммы напряжение удерживается на крайнем значении.
Public Function GetStress(ByVal strain As Double) As Double
    EnsureInitialized

    If strain <= mStrains(1) Then
        GetStress = mStresses(1)
    ElseIf strain >= mStrains(mPointCount) Then
        GetStress = mStresses(mPointCount)
    Else
        GetStress = InterpolateStress(strain)
    End If
End Function

' Возвращает касательный модуль dσ/dε для Newton-матрицы.
' На плато за крайними точками модуль равен нулю.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If strain < mStrains(1) Or strain > mStrains(mPointCount) Then Exit Function
    GetTangentModulus = SegmentTangent(strain)
End Function

' Возвращает секущий модуль σ/ε.
' Около нуля используется касательный модуль, чтобы избежать деления на слишком
' маленькую деформацию.
Public Function GetSecantModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If Abs(strain) <= 0.000000000001 Then
        GetSecantModulus = GetTangentModulus(strain)
    Else
        GetSecantModulus = GetStress(strain) / strain
    End If
End Function

' Выполняет линейную интерполяцию напряжения между соседними точками диаграммы.
Private Function InterpolateStress(ByVal strain As Double) As Double
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            InterpolateStress = mStresses(i) + (mStresses(i + 1) - mStresses(i)) * _
                (strain - mStrains(i)) / (mStrains(i + 1) - mStrains(i))
            Exit Function
        End If
    Next i
End Function

' Возвращает наклон текущего отрезка диаграммы.
Private Function SegmentTangent(ByVal strain As Double) As Double
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            SegmentTangent = (mStresses(i + 1) - mStresses(i)) / (mStrains(i + 1) - mStrains(i))
            Exit Function
        End If
    Next i
End Function

' Сортирует точки по возрастанию деформации.
Private Sub SortPointsByStrain()
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(mStrains) - 1
        For j = i + 1 To UBound(mStrains)
            If mStrains(j) < mStrains(i) Then SwapPoints i, j
        Next j
    Next i
End Sub

' Меняет местами две точки диаграммы при сортировке.
Private Sub SwapPoints(ByVal firstIndex As Long, ByVal secondIndex As Long)
    Dim tmp As Double
    tmp = mStrains(firstIndex)
    mStrains(firstIndex) = mStrains(secondIndex)
    mStrains(secondIndex) = tmp
    tmp = mStresses(firstIndex)
    mStresses(firstIndex) = mStresses(secondIndex)
    mStresses(secondIndex) = tmp
End Sub

' Проверяет, что после сортировки нет повторяющихся деформаций.
Private Sub ValidatePointOrder()
    Dim i As Long
    For i = 1 To UBound(mStrains) - 1
        If mStrains(i + 1) <= mStrains(i) Then
            Err.Raise vbObjectError + 3107, "CConcreteDiagramMaterial", "Деформации в диаграмме бетона не должны повторяться."
        End If
    Next i
End Sub

' Проверяет, что материал уже получил диаграмму.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3103, "CConcreteDiagramMaterial", "Материал бетона не инициализирован."
End Sub
