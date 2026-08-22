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
' Материал получает массив точек strain-stress и возвращает напряжение,
' касательный и секущий модуль по текущей деформации. Переключателей
' двухлинейная/трехлинейная диаграмма нет: расчет работает только с точками.
'
' TensionMode решает, участвует ли растянутый бетон в усилиях или его
' напряжение обнуляется при растяжении.

Private mStrains() As Double ' Узлы диаграммы по деформации, отсортированные от сжатия к растяжению.
Private mStresses() As Double ' Напряжения в узлах диаграммы во внутренних единицах, МПа.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова GetStress до задания диаграммы.
Private mTensionMode As String ' Ignore выключает растянутый бетон, UseDiagram оставляет растянутую ветвь.
Private mTensionElasticModulus As Double ' Резервная линейная жесткость растяжения для fallback-ветви.
Private mTensionStressLimit As Double ' Ограничение растянутого напряжения для fallback-ветви.

' По умолчанию растянутый бетон не участвует во внутренних усилиях.
Private Sub Class_Initialize()
    mTensionMode = "Ignore"
End Sub

' Инициализирует диаграмму по точкам Config и применяет настройки материала.
Public Sub InitializeFromSettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    InitializeFromSettingsPoints settings, units
    ApplySettings settings, units
End Sub

' Читает точки бетонной диаграммы из reader-а.
' Reader хранит напряжения в пользовательской INPUT-единице, поэтому здесь
' выполняется перевод в МПа через CUnitSystem.
Private Sub InitializeFromSettingsPoints(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem)
    Dim count As Long
    count = settings.ConcreteDiagramPointCount
    If count < 2 Then Err.Raise vbObjectError + 3106, "CConcreteDiagramMaterial", "Диаграмма бетона должна содержать не менее двух точек."

    ReDim mStrains(1 To count)
    ReDim mStresses(1 To count)

    Dim i As Long
    ' Копируем точки в локальные массивы материала, чтобы solver не зависел от
    ' Excel-reader-а и мог быстро интерполировать напряжение в hot loop.
    For i = 1 To count
        mStrains(i) = settings.ConcreteDiagramStrain(i)
        mStresses(i) = settings.ConcreteDiagramStress(i)
        If Not units Is Nothing Then mStresses(i) = units.InputStressToInternal(mStresses(i))
    Next i

    SortPointsByStrain
    ValidatePointOrder
    mPointCount = count
    mInitialized = True
End Sub

' Тестовый и служебный конструктор двухлинейной сжатой ветви.
' Нулевая точка добавляется автоматически, чтобы диаграмма замыкалась в начале.
Public Sub InitializeFromTwoCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    If eps1 >= 0# Then Err.Raise vbObjectError + 3100, "CConcreteDiagramMaterial", "first compression strain must be negative."
    If eps2 >= eps1 Then Err.Raise vbObjectError + 3101, "CConcreteDiagramMaterial", "second compression strain must be more negative than first."
    If stress1 >= 0# Or stress2 >= 0# Then Err.Raise vbObjectError + 3102, "CConcreteDiagramMaterial", "Concrete compression stresses must be negative."

    ReDim mStrains(1 To 3)
    ReDim mStresses(1 To 3)
    mStrains(1) = eps2: mStresses(1) = stress2
    mStrains(2) = eps1: mStresses(2) = stress1
    mStrains(3) = 0#: mStresses(3) = 0#
    mPointCount = 3
    mInitialized = True
End Sub

' Короткий старый алиас для InitializeFromTwoCompressionPoints.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    InitializeFromTwoCompressionPoints eps1, stress1, eps2, stress2
End Sub

' Тестовый и служебный конструктор трехточечной сжатой ветви.
Public Sub InitializeFromThreeCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    If eps1 >= 0# Then Err.Raise vbObjectError + 3100, "CConcreteDiagramMaterial", "first compression strain must be negative."
    If eps2 >= eps1 Then Err.Raise vbObjectError + 3101, "CConcreteDiagramMaterial", "second compression strain must be more negative than first."
    If eps3 >= eps2 Then Err.Raise vbObjectError + 3104, "CConcreteDiagramMaterial", "third compression strain must be more negative than second."
    If stress1 >= 0# Or stress2 >= 0# Or stress3 >= 0# Then Err.Raise vbObjectError + 3102, "CConcreteDiagramMaterial", "Concrete compression stresses must be negative."

    ReDim mStrains(1 To 4)
    ReDim mStresses(1 To 4)
    mStrains(1) = eps3: mStresses(1) = stress3
    mStrains(2) = eps2: mStresses(2) = stress2
    mStrains(3) = eps1: mStresses(3) = stress1
    mStrains(4) = 0#: mStresses(4) = 0#
    mPointCount = 4
    mInitialized = True
End Sub

' Возвращает наиболее сжатую точку диаграммы как ориентир предельной деформации.
Public Property Get UltimateCompressionStrain() As Double
    EnsureInitialized
    UltimateCompressionStrain = mStrains(1)
End Property

' Возвращает напряжение бетона по текущей деформации.
' Если растянутый бетон выключен, положительная деформация сразу дает 0:
' элемент остается геометрически в модели, но не дает внутреннего усилия.
Public Function GetStress(ByVal strain As Double) As Double
    EnsureInitialized

    If strain > 0# And StrComp(mTensionMode, "UseDiagram", vbTextCompare) <> 0 Then
        GetStress = 0#
    ElseIf strain <= mStrains(1) Then
        GetStress = mStresses(1)
    ElseIf strain > mStrains(mPointCount) Then
        If strain > 0# And StrComp(mTensionMode, "UseDiagram", vbTextCompare) = 0 And mStrains(mPointCount) <= 0# Then
            GetStress = GetTensionStressFallback(strain)
        Else
            GetStress = mStresses(mPointCount)
        End If
    Else
        GetStress = InterpolateStress(strain)
    End If
End Function

' Возвращает касательный модуль dσ/dε для Newton-матрицы.
' На плато за крайними точками модуль равен 0, потому что напряжение больше
' не меняется при росте деформации.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    EnsureInitialized

    If strain > 0# And StrComp(mTensionMode, "UseDiagram", vbTextCompare) <> 0 Then
        GetTangentModulus = 0#
    ElseIf strain <= mStrains(1) Then
        GetTangentModulus = 0#
    ElseIf strain > mStrains(mPointCount) Then
        If strain > 0# And StrComp(mTensionMode, "UseDiagram", vbTextCompare) = 0 And mStrains(mPointCount) <= 0# Then
            GetTangentModulus = GetTensionTangentFallback(strain)
        Else
            GetTangentModulus = 0#
        End If
    Else
        GetTangentModulus = SegmentTangent(strain)
    End If
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

' Применяет настройки растянутого бетона.
' Основная ветвь берется из точек диаграммы; Eb/Rbt используются только для
' fallback-режима, когда растяжение включено, но точек растянутой ветви нет.
Public Sub ApplySettings(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing)
    If settings Is Nothing Then Exit Sub
    TensionMode = settings.GetString("Concrete.TensionMode", mTensionMode)
    mTensionElasticModulus = settings.GetDouble("Concrete.Eb", mTensionElasticModulus)
    mTensionStressLimit = settings.GetDouble("Concrete.Rbt.SLS", mTensionStressLimit)
    If Not units Is Nothing Then
        mTensionElasticModulus = units.InputStressToInternal(mTensionElasticModulus)
        mTensionStressLimit = units.InputStressToInternal(mTensionStressLimit)
    End If
End Sub

Public Property Let TensionMode(ByVal value As String)
    If StrComp(Trim$(value), "UseDiagram", vbTextCompare) = 0 Then
        mTensionMode = "UseDiagram"
    Else
        mTensionMode = "Ignore"
    End If
End Property

Public Property Get TensionMode() As String
    TensionMode = mTensionMode
End Property

Public Property Let TensionElasticModulus(ByVal value As Double)
    mTensionElasticModulus = value
End Property

Public Property Get TensionElasticModulus() As Double
    TensionElasticModulus = mTensionElasticModulus
End Property

Public Property Let TensionStressLimit(ByVal value As Double)
    mTensionStressLimit = value
End Property

Public Property Get TensionStressLimit() As Double
    TensionStressLimit = mTensionStressLimit
End Property

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

Private Function SegmentTangent(ByVal strain As Double) As Double
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            SegmentTangent = (mStresses(i + 1) - mStresses(i)) / (mStrains(i + 1) - mStrains(i))
            Exit Function
        End If
    Next i
End Function

Private Function GetTensionStressFallback(ByVal strain As Double) As Double
    If StrComp(mTensionMode, "UseDiagram", vbTextCompare) <> 0 Then Exit Function
    If strain <= 0# Or mTensionElasticModulus <= 0# Or mTensionStressLimit <= 0# Then Exit Function
    GetTensionStressFallback = mTensionElasticModulus * strain
    If GetTensionStressFallback > mTensionStressLimit Then GetTensionStressFallback = mTensionStressLimit
End Function

Private Function GetTensionTangentFallback(ByVal strain As Double) As Double
    If StrComp(mTensionMode, "UseDiagram", vbTextCompare) <> 0 Then Exit Function
    If strain <= 0# Or mTensionElasticModulus <= 0# Or mTensionStressLimit <= 0# Then Exit Function
    If mTensionElasticModulus * strain < mTensionStressLimit Then
        GetTensionTangentFallback = mTensionElasticModulus
    End If
End Function

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
            Err.Raise vbObjectError + 3107, "CConcreteDiagramMaterial", "Деформации в диаграмме бетона не должны повторяться."
        End If
    Next i
End Sub

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3103, "CConcreteDiagramMaterial", "Материал не инициализирован."
End Sub

