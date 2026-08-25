VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSteelDiagramMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Кусочно-линейная диаграмма арматуры
' ==========================================================================
' Класс является расчетным материалом для solver-а. Он не знает класс
' арматуры, ГПС, профиль стержня и нормативные формулы построения диаграммы:
' все это заранее решает CMaterialModelProvider.
'
' На вход сюда попадает только готовый набор точек во внутренних единицах.
' Solver дальше вызывает единый материалный контракт GetStress,
' GetTangentModulus и GetSecantModulus для каждого стержня.

Private mStrains() As Double ' Узлы диаграммы по деформации, отсортированные от сжатия к растяжению.
Private mStresses() As Double ' Напряжения в узлах диаграммы во внутренних единицах, МПа.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова материала до задания диаграммы.

' ------------------------------
' Инициализация
' ------------------------------
' Загружает точки, построенные CSteelDiagramBuilder по выбранному расчетному
' назначению и типу диаграммы.
Public Sub InitializeFromPointSet(ByVal points As CMaterialPointSet)
    If points Is Nothing Then Err.Raise vbObjectError + 3125, "CSteelDiagramMaterial", "Точки диаграммы арматуры не переданы."
    If points.Count < 2 Then Err.Raise vbObjectError + 3125, "CSteelDiagramMaterial", "Диаграмма арматуры должна содержать не менее двух точек."

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

' Тестовый конструктор симметричной двухлинейной диаграммы арматуры.
' По заданной растянутой ветви автоматически создается такая же сжатая ветвь.
Public Sub InitializeFromTwoTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet

    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "Первая деформация растяжения арматуры должна быть положительной."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "Вторая деформация растяжения арматуры должна быть больше первой."
    If stress1 <= 0# Or stress2 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Напряжения растяжения арматуры должны быть положительными."

    points.AddPoint -eps2, -stress2
    points.AddPoint -eps1, -stress1
    points.AddPoint 0#, 0#
    points.AddPoint eps1, stress1
    points.AddPoint eps2, stress2
    InitializeFromPointSet points
End Sub

' Короткий старый алиас для тестов, которые создают простой материал вручную.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, ByVal eps2 As Double)
    InitializeFromTwoTensionPoints eps1, stress1, eps2, stress1
End Sub

' Тестовый конструктор симметричной трехточечной диаграммы арматуры.
Public Sub InitializeFromThreeTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    Dim points As CMaterialPointSet
    Set points = New CMaterialPointSet

    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "Первая деформация растяжения арматуры должна быть положительной."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "Вторая деформация растяжения арматуры должна быть больше первой."
    If eps3 <= eps2 Then Err.Raise vbObjectError + 3124, "CSteelDiagramMaterial", "Третья деформация растяжения арматуры должна быть больше второй."
    If stress1 <= 0# Or stress2 <= 0# Or stress3 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Напряжения растяжения арматуры должны быть положительными."

    points.AddPoint -eps3, -stress3
    points.AddPoint -eps2, -stress2
    points.AddPoint -eps1, -stress1
    points.AddPoint 0#, 0#
    points.AddPoint eps1, stress1
    points.AddPoint eps2, stress2
    points.AddPoint eps3, stress3
    InitializeFromPointSet points
End Sub

' Возвращает крайнюю растянутую точку диаграммы как ориентир предельной деформации.
Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    UltimateTensionStrain = mStrains(mPointCount)
End Property

' Возвращает крайнюю сжатую точку диаграммы как ориентир предельной деформации.
Public Property Get UltimateCompressionStrain() As Double
    EnsureInitialized
    UltimateCompressionStrain = mStrains(1)
End Property

' Количество точек используется тестами построителя диаграмм.
Public Property Get PointCount() As Long
    PointCount = mPointCount
End Property

' Возвращает деформацию узла диаграммы по индексу.
Public Function PointStrain(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3127, "CSteelDiagramMaterial", "Индекс точки диаграммы арматуры вне диапазона."
    PointStrain = mStrains(index)
End Function

' Возвращает напряжение узла диаграммы по индексу.
Public Function PointStress(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3128, "CSteelDiagramMaterial", "Индекс точки диаграммы арматуры вне диапазона."
    PointStress = mStresses(index)
End Function

' ------------------------------
' Расчетный контракт материала
' ------------------------------
' Возвращает напряжение стержня по его деформации.
' За пределами заданной диаграммы напряжение удерживается на крайнем значении.
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

' Возвращает касательный модуль участка диаграммы.
' За крайними точками модуль равен нулю, потому что напряжение не меняется.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If strain <= mStrains(1) Or strain >= mStrains(mPointCount) Then Exit Function
    GetTangentModulus = SegmentTangent(strain)
End Function

' Возвращает секущий модуль σ/ε; около нуля использует касательный модуль.
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

' Сортирует пользовательские точки по возрастанию деформации.
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

' Проверяет, что в диаграмме нет повторяющихся деформаций.
Private Sub ValidatePointOrder()
    Dim i As Long
    For i = 1 To UBound(mStrains) - 1
        If mStrains(i + 1) <= mStrains(i) Then
            Err.Raise vbObjectError + 3126, "CSteelDiagramMaterial", "Деформации в диаграмме арматуры не должны повторяться."
        End If
    Next i
End Sub

' Проверяет, что материал уже получил диаграмму.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3123, "CSteelDiagramMaterial", "Материал арматуры не инициализирован."
End Sub
