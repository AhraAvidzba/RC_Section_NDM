VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMaterialDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Универсальная кусочно-линейная диаграмма материала
' ==========================================================================
' Класс хранит только готовые точки диаграммы "деформация - напряжение" во
' внутренних единицах проекта: деформация безразмерная, напряжение в МПа.
' Он не знает, для какого материала построена диаграмма, по какой группе
' предельных состояний она нужна и какой расчетный режим ее запросил.
'
' Все нормативные решения остаются выше, в CMaterialModelProvider: выбор I/II
' ГПС, выбор TwoLine/ThreeLine, сопротивлений R и режима растянутого бетона.
' Solver получает уже готовый объект и вызывает общий материалный контракт:
' StressAtStrain/GetStress, GetTangentModulus и GetSecantModulus.

Private mStrains() As Double ' Узлы диаграммы по деформации, отсортированные от сжатия к растяжению.
Private mStresses() As Double ' Напряжения в узлах диаграммы, МПа во внутренней системе.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова расчетных методов до загрузки точек.

' ------------------------------
' Инициализация
' ------------------------------
' Загружает готовые массивы точек диаграммы. Provider передает сюда уже
' рассчитанные и знаково оформленные координаты: сжатие в проекте имеет
' отрицательную деформацию и отрицательное напряжение, растяжение - положительное.
Public Sub InitializeFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long)
    If pointCount < 2 Then Err.Raise vbObjectError + 3100, "CMaterialDiagram", "Диаграмма материала должна содержать не менее двух точек."

    ReDim mStrains(1 To pointCount)
    ReDim mStresses(1 To pointCount)

    Dim i As Long
    For i = 1 To pointCount
        mStrains(i) = strains(i)
        mStresses(i) = stresses(i)
    Next i

    SortPointsByStrain
    ValidatePointOrder
    mPointCount = pointCount
    mInitialized = True
End Sub

' Создает простую сжатую ветвь с нулем. Метод оставлен как удобный тестовый
' конструктор для старых unit-тестов и служебных расчетов; production-код
' должен получать диаграммы через CMaterialModelProvider.
Public Sub InitializeFromTwoCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    Dim strains(1 To 3) As Double
    Dim stresses(1 To 3) As Double

    If eps1 >= 0# Or eps2 >= 0# Then Err.Raise vbObjectError + 3101, "CMaterialDiagram", "Сжатые деформации должны быть отрицательными."
    If stress1 >= 0# Or stress2 >= 0# Then Err.Raise vbObjectError + 3102, "CMaterialDiagram", "Сжатые напряжения должны быть отрицательными."

    strains(1) = eps2: stresses(1) = stress2
    strains(2) = eps1: stresses(2) = stress1
    strains(3) = 0#: stresses(3) = 0#
    InitializeFromArrays strains, stresses, 3
End Sub

' Создает трехточечную сжатую ветвь с нулем. Используется тестами для проверки
' интерполяции на нескольких отрезках без подключения Excel и provider-а.
Public Sub InitializeFromThreeCompressionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    Dim strains(1 To 4) As Double
    Dim stresses(1 To 4) As Double

    If eps1 >= 0# Or eps2 >= 0# Or eps3 >= 0# Then Err.Raise vbObjectError + 3103, "CMaterialDiagram", "Сжатые деформации должны быть отрицательными."
    If stress1 >= 0# Or stress2 >= 0# Or stress3 >= 0# Then Err.Raise vbObjectError + 3104, "CMaterialDiagram", "Сжатые напряжения должны быть отрицательными."

    strains(1) = eps3: stresses(1) = stress3
    strains(2) = eps2: stresses(2) = stress2
    strains(3) = eps1: stresses(3) = stress1
    strains(4) = 0#: stresses(4) = 0#
    InitializeFromArrays strains, stresses, 4
End Sub

' Создает симметричную растянутую/сжатую диаграмму с двумя точками на ветвь.
' Это тестовый конструктор для арматуры: реальный расчет строит такую диаграмму
' через provider, чтобы все нормативные решения оставались в одном месте.
Public Sub InitializeFromTwoTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double)
    Dim strains(1 To 5) As Double
    Dim stresses(1 To 5) As Double

    If eps1 <= 0# Or eps2 <= eps1 Then Err.Raise vbObjectError + 3105, "CMaterialDiagram", "Растянутые деформации должны возрастать и быть положительными."
    If stress1 <= 0# Or stress2 <= 0# Then Err.Raise vbObjectError + 3106, "CMaterialDiagram", "Растянутые напряжения должны быть положительными."

    strains(1) = -eps2: stresses(1) = -stress2
    strains(2) = -eps1: stresses(2) = -stress1
    strains(3) = 0#: stresses(3) = 0#
    strains(4) = eps1: stresses(4) = stress1
    strains(5) = eps2: stresses(5) = stress2
    InitializeFromArrays strains, stresses, 5
End Sub

' Создает симметричную диаграмму с тремя точками на растянутую и сжатую ветвь.
Public Sub InitializeFromThreeTensionPoints(ByVal eps1 As Double, ByVal stress1 As Double, _
        ByVal eps2 As Double, ByVal stress2 As Double, ByVal eps3 As Double, ByVal stress3 As Double)
    Dim strains(1 To 7) As Double
    Dim stresses(1 To 7) As Double

    If eps1 <= 0# Or eps2 <= eps1 Or eps3 <= eps2 Then Err.Raise vbObjectError + 3107, "CMaterialDiagram", "Растянутые деформации должны возрастать и быть положительными."
    If stress1 <= 0# Or stress2 <= 0# Or stress3 <= 0# Then Err.Raise vbObjectError + 3108, "CMaterialDiagram", "Растянутые напряжения должны быть положительными."

    strains(1) = -eps3: stresses(1) = -stress3
    strains(2) = -eps2: stresses(2) = -stress2
    strains(3) = -eps1: stresses(3) = -stress1
    strains(4) = 0#: stresses(4) = 0#
    strains(5) = eps1: stresses(5) = stress1
    strains(6) = eps2: stresses(6) = stress2
    strains(7) = eps3: stresses(7) = stress3
    InitializeFromArrays strains, stresses, 7
End Sub

' Короткий конструктор для старых тестов: при четырех аргументах задается
' сжатая ветвь, при трех - симметричная двухлинейная ветвь арматуры.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, ByVal eps2 As Double, _
        Optional ByVal stress2 As Variant)
    If IsMissing(stress2) Then
        InitializeFromTwoTensionPoints eps1, stress1, eps2, stress1
    Else
        InitializeFromTwoCompressionPoints eps1, stress1, eps2, CDbl(stress2)
    End If
End Sub

' ------------------------------
' Доступ к точкам и пределам
' ------------------------------
' Возвращает крайнюю сжатую точку диаграммы. Capacity использует ее как
' предельную сжатую деформацию, потому что отдельные Capacity.*Limit удалены.
Public Property Get UltimateCompressionStrain() As Double
    EnsureInitialized
    UltimateCompressionStrain = mStrains(1)
End Property

' Возвращает крайнюю растянутую точку диаграммы. Если диаграмма не имеет
' положительной ветви, предел растяжения равен нулю и не участвует в проверке.
Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    If mStrains(mPointCount) > 0# Then UltimateTensionStrain = mStrains(mPointCount)
End Property

Public Property Get PointCount() As Long
    PointCount = mPointCount
End Property

' Возвращает деформацию узла по индексу для тестов и контрольных таблиц.
Public Function PointStrain(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3109, "CMaterialDiagram", "Индекс точки диаграммы вне диапазона."
    PointStrain = mStrains(index)
End Function

' Возвращает напряжение узла по индексу для тестов и контрольных таблиц.
Public Function PointStress(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3110, "CMaterialDiagram", "Индекс точки диаграммы вне диапазона."
    PointStress = mStresses(index)
End Function

' ------------------------------
' Расчетный контракт материала
' ------------------------------
' Возвращает напряжение по заданной деформации. За пределами последнего узла
' напряжение удерживается на крайнем значении, что моделирует плато диаграммы.
Public Function StressAtStrain(ByVal strain As Double) As Double
    EnsureInitialized

    If strain <= mStrains(1) Then
        StressAtStrain = mStresses(1)
    ElseIf strain >= mStrains(mPointCount) Then
        StressAtStrain = mStresses(mPointCount)
    Else
        StressAtStrain = InterpolateStress(strain)
    End If
End Function

' Совместимое имя для существующего solver API.
Public Function GetStress(ByVal strain As Double) As Double
    GetStress = StressAtStrain(strain)
End Function

' Возвращает касательный модуль текущего отрезка диаграммы. За пределами
' заданной области модуль равен нулю, потому что напряжение уже на плато.
Public Function GetTangentModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If strain < mStrains(1) Or strain > mStrains(mPointCount) Then Exit Function
    GetTangentModulus = SegmentTangent(strain)
End Function

' Возвращает секущий модуль stress/strain; около нуля используется касательный
' модуль, чтобы не делить на почти нулевую деформацию.
Public Function GetSecantModulus(ByVal strain As Double) As Double
    EnsureInitialized
    If Abs(strain) <= 0.000000000001 Then
        GetSecantModulus = GetTangentModulus(strain)
    Else
        GetSecantModulus = StressAtStrain(strain) / strain
    End If
End Function

' Линейно интерполирует напряжение между соседними узлами диаграммы.
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

' Возвращает наклон того же отрезка, который содержит текущую деформацию.
Private Function SegmentTangent(ByVal strain As Double) As Double
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            SegmentTangent = (mStresses(i + 1) - mStresses(i)) / _
                (mStrains(i + 1) - mStrains(i))
            Exit Function
        End If
    Next i
End Function

' Сортирует точки по деформации, чтобы provider мог добавлять их в удобном
' инженерном порядке, а расчет всегда шел по монотонной оси strain.
Private Sub SortPointsByStrain()
    Dim i As Long
    Dim j As Long
    For i = 1 To UBound(mStrains) - 1
        For j = i + 1 To UBound(mStrains)
            If mStrains(j) < mStrains(i) Then SwapPoints i, j
        Next j
    Next i
End Sub

' Меняет местами две точки при сортировке.
Private Sub SwapPoints(ByVal firstIndex As Long, ByVal secondIndex As Long)
    Dim tmp As Double
    tmp = mStrains(firstIndex)
    mStrains(firstIndex) = mStrains(secondIndex)
    mStrains(secondIndex) = tmp
    tmp = mStresses(firstIndex)
    mStresses(firstIndex) = mStresses(secondIndex)
    mStresses(secondIndex) = tmp
End Sub

' Проверяет, что после сортировки нет повторяющихся деформаций. Две разные
' точки с одинаковым strain дали бы вертикальный участок и бесконечный модуль.
Private Sub ValidatePointOrder()
    Dim i As Long
    For i = 1 To UBound(mStrains) - 1
        If mStrains(i + 1) <= mStrains(i) Then
            Err.Raise vbObjectError + 3111, "CMaterialDiagram", "Деформации в диаграмме материала не должны повторяться."
        End If
    Next i
End Sub

' Проверяет, что диаграмма уже получила точки.
Private Sub EnsureInitialized()
    If Not mInitialized Then Err.Raise vbObjectError + 3112, "CMaterialDiagram", "Диаграмма материала не инициализирована."
End Sub
