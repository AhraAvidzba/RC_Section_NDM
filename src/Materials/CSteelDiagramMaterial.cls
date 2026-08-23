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
' Материал арматуры интерполирует напряжения по пользовательским точкам
' диаграммы. Он не знает о диаметрах и координатах стержней: solver передает
' сюда только деформацию конкретного стержня.

Private mStrains() As Double ' Узлы диаграммы по деформации, отсортированные от сжатия к растяжению.
Private mStresses() As Double ' Напряжения в узлах диаграммы во внутренних единицах, МПа.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова материала до задания диаграммы.

' Инициализирует диаграмму арматуры по точкам Config.
' Напряжения в reader-е еще записаны в пользовательской INPUT-единице, поэтому
' здесь они переводятся в МПа через CUnitSystem.
Public Sub InitializeFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    Dim count As Long
    count = settings.SteelDiagramPointCount
    If count < 2 Then Err.Raise vbObjectError + 3125, "CSteelDiagramMaterial", "Диаграмма арматуры должна содержать не менее двух точек."

    ReDim mStrains(1 To count)
    ReDim mStresses(1 To count)

    Dim i As Long
    ' Копируем точки в локальные массивы, чтобы solver мог быстро получать
    ' напряжение стержня без обращения к Excel-reader-у.
    For i = 1 To count
        mStrains(i) = settings.SteelDiagramStrain(i)
        mStresses(i) = settings.SteelDiagramStress(i)
        If Not units Is Nothing Then mStresses(i) = units.InputStressToInternal(mStresses(i))
    Next i

    SortPointsByStrain
    ValidatePointOrder
    mPointCount = count
    mInitialized = True
End Sub

' Тестовый конструктор симметричной двухлинейной диаграммы арматуры.
' По заданной растянутой ветви автоматически создается такая же сжатая ветвь.
Public Sub InitializeFromTwoTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "Первая деформация растяжения арматуры должна быть положительной."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "Вторая деформация растяжения арматуры должна быть больше первой."
    If stress1 <= 0# Or stress2 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Напряжения растяжения арматуры должны быть положительными."

    ReDim mStrains(1 To 5)
    ReDim mStresses(1 To 5)
    mStrains(1) = -eps2: mStresses(1) = -stress2
    mStrains(2) = -eps1: mStresses(2) = -stress1
    mStrains(3) = 0#: mStresses(3) = 0#
    mStrains(4) = eps1: mStresses(4) = stress1
    mStrains(5) = eps2: mStresses(5) = stress2
    mPointCount = 5
    mInitialized = True
End Sub

' Короткий старый алиас для двухточечной растянутой ветви.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, ByVal eps2 As Double)
    InitializeFromTwoTensionPoints eps1, stress1, eps2, stress1
End Sub

' Тестовый конструктор симметричной трехточечной диаграммы арматуры.
Public Sub InitializeFromThreeTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    If eps1 <= 0# Then Err.Raise vbObjectError + 3120, "CSteelDiagramMaterial", "Первая деформация растяжения арматуры должна быть положительной."
    If eps2 <= eps1 Then Err.Raise vbObjectError + 3121, "CSteelDiagramMaterial", "Вторая деформация растяжения арматуры должна быть больше первой."
    If eps3 <= eps2 Then Err.Raise vbObjectError + 3124, "CSteelDiagramMaterial", "Третья деформация растяжения арматуры должна быть больше второй."
    If stress1 <= 0# Or stress2 <= 0# Or stress3 <= 0# Then Err.Raise vbObjectError + 3122, "CSteelDiagramMaterial", "Напряжения растяжения арматуры должны быть положительными."

    ReDim mStrains(1 To 7)
    ReDim mStresses(1 To 7)
    mStrains(1) = -eps3: mStresses(1) = -stress3
    mStrains(2) = -eps2: mStresses(2) = -stress2
    mStrains(3) = -eps1: mStresses(3) = -stress1
    mStrains(4) = 0#: mStresses(4) = 0#
    mStrains(5) = eps1: mStresses(5) = stress1
    mStrains(6) = eps2: mStresses(6) = stress2
    mStrains(7) = eps3: mStresses(7) = stress3
    mPointCount = 7
    mInitialized = True
End Sub

' Возвращает крайнюю растянутую точку диаграммы как ориентир предельной деформации.
Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    UltimateTensionStrain = mStrains(mPointCount)
End Property

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
' За крайними точками модуль равен 0, потому что напряжение не изменяется.
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
' Это позволяет вводить таблицу удобно, а расчету всегда работать с правильным
' порядком интерполяции.
Private Sub SortPointsByStrain()
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(mStrains) - 1
        For j = i + 1 To UBound(mStrains)
            If mStrains(j) < mStrains(i) Then
                SwapPoints i, j
            End If
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

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidatePointOrder()
    Dim i As Long
    For i = 1 To UBound(mStrains) - 1
        If mStrains(i + 1) <= mStrains(i) Then
            Err.Raise vbObjectError + 3126, "CSteelDiagramMaterial", "Деформации в диаграмме арматуры не должны повторяться."
        End If
    Next i
End Sub

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3123, "CSteelDiagramMaterial", "Материал не инициализирован."
End Sub

