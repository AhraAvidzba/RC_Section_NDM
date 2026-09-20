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
Private mSegmentTangent() As Double ' Заранее рассчитанный наклон каждого линейного участка диаграммы, МПа.
Private mSegmentIntercept() As Double ' Свободный член sigma = E_t * epsilon + b для быстрого расчета stress.
Private mPointCount As Long ' Количество активных точек диаграммы.
Private mInitialized As Boolean ' Защита от вызова расчетных методов до загрузки точек.
Private mPhysicalCompressionStrain As Double ' Физический сжатый предел; extension-точки его не меняют.
Private mPhysicalTensionStrain As Double ' Физический растянутый предел, если у диаграммы есть растянутая ветвь.
Private mHasPhysicalTensionLimit As Boolean ' False для диаграмм, где растянутый бетон принят с нулевым сопротивлением без eps-лимита.
Private mHasCompressionExtension As Boolean ' True, если слева от физического сжатого предела добавлена техническая область.
Private mHasTensionExtension As Boolean ' True, если справа от физического растянутого предела добавлена техническая область.
Private mCompressionExtensionStartStrain As Double ' Граница, после которой сжатая ветвь уже использует numerical extension.
Private mTensionExtensionStartStrain As Double ' Граница, после которой растянутая ветвь уже использует numerical extension.
Private Const RANGE_TOLERANCE As Double = 0.000000000001 ' Малый допуск для сравнения деформаций с границами диаграммы.

' ------------------------------
' Инициализация
' ------------------------------
' Загружает готовые массивы точек диаграммы. Provider передает сюда уже
' рассчитанные и знаково оформленные координаты: сжатие в проекте имеет
' отрицательную деформацию и отрицательное напряжение, растяжение - положительное.
Public Sub InitializeFromArrays(ByRef strains() As Double, ByRef stresses() As Double, ByVal pointCount As Long, _
        Optional ByVal physicalCompressionStrain As Variant, Optional ByVal physicalTensionStrain As Variant, _
        Optional ByVal hasCompressionExtension As Boolean = False, Optional ByVal hasTensionExtension As Boolean = False, _
        Optional ByVal compressionExtensionStartStrain As Variant, Optional ByVal tensionExtensionStartStrain As Variant)
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
    BuildSegmentCache
    If IsMissing(physicalCompressionStrain) Then
        mPhysicalCompressionStrain = mStrains(1)
    Else
        mPhysicalCompressionStrain = CDbl(physicalCompressionStrain)
    End If
    If IsMissing(physicalTensionStrain) Then
        mPhysicalTensionStrain = IIf(mStrains(mPointCount) > 0#, mStrains(mPointCount), 0#)
    Else
        mPhysicalTensionStrain = CDbl(physicalTensionStrain)
    End If
    mHasPhysicalTensionLimit = (mPhysicalTensionStrain > 0#)
    mHasCompressionExtension = hasCompressionExtension
    mHasTensionExtension = hasTensionExtension
    If IsMissing(compressionExtensionStartStrain) Then
        mCompressionExtensionStartStrain = mPhysicalCompressionStrain
    Else
        mCompressionExtensionStartStrain = CDbl(compressionExtensionStartStrain)
    End If
    If IsMissing(tensionExtensionStartStrain) Then
        mTensionExtensionStartStrain = mPhysicalTensionStrain
    Else
        mTensionExtensionStartStrain = CDbl(tensionExtensionStartStrain)
    End If
    mInitialized = True
End Sub

' ------------------------------
' Доступ к расчетным пределам диаграммы
' ------------------------------
' Возвращает крайнюю сжатую точку диаграммы. Capacity использует ее как
' предельную сжатую деформацию, потому что отдельные Capacity.*Limit удалены.
Public Property Get UltimateCompressionStrain() As Double
    EnsureInitialized
    UltimateCompressionStrain = mPhysicalCompressionStrain
End Property

' Возвращает крайнюю растянутую точку диаграммы. Если диаграмма не имеет
' положительной ветви, предел растяжения равен нулю и не участвует в проверке.
Public Property Get UltimateTensionStrain() As Double
    EnsureInitialized
    If mHasPhysicalTensionLimit Then UltimateTensionStrain = mPhysicalTensionStrain
End Property

' Возвращает физический сжатый предел диаграммы. Если provider добавил
' техническую extension-точку, это свойство все равно остается на реальном eps_ult.
Public Property Get PhysicalCompressionStrain() As Double
    EnsureInitialized
    PhysicalCompressionStrain = mPhysicalCompressionStrain
End Property

' Возвращает физический растянутый предел, если он есть. Для бетона с
' отключенным растяжением предела нет: положительная деформация означает
' неработающую растянутую зону, а не автоматический выход в extension.
Public Property Get PhysicalTensionStrain() As Double
    EnsureInitialized
    PhysicalTensionStrain = mPhysicalTensionStrain
End Property

Public Property Get HasPhysicalTensionLimit() As Boolean
    EnsureInitialized
    HasPhysicalTensionLimit = mHasPhysicalTensionLimit
End Property

Public Property Get HasCompressionExtension() As Boolean
    EnsureInitialized
    HasCompressionExtension = mHasCompressionExtension
End Property

Public Property Get HasTensionExtension() As Boolean
    EnsureInitialized
    HasTensionExtension = mHasTensionExtension
End Property

' Проверяет, остается ли деформация в физической области диаграммы.
' Отсутствие растянутого предела трактуется как неработающая растянутая зона
' с нулевым сопротивлением, поэтому положительные eps в этом случае допустимы.
Public Function IsInPhysicalRange(ByVal strain As Double) As Boolean
    EnsureInitialized
    If strain < mPhysicalCompressionStrain - RANGE_TOLERANCE Then Exit Function
    If mHasPhysicalTensionLimit Then
        If strain > mPhysicalTensionStrain + RANGE_TOLERANCE Then Exit Function
    End If
    IsInPhysicalRange = True
End Function

' True означает, что деформация вышла за физический eps_ult именно в ту
' техническую область, которую provider добавил для прямого StateSolution.
' Горизонтальное физическое плато до eps_ult не считается extension: оно
' остается допустимой частью диаграммы материала и не должно давать FAIL.
Public Function IsInExtensionRange(ByVal strain As Double) As Boolean
    EnsureInitialized
    If mHasCompressionExtension And strain < mPhysicalCompressionStrain - RANGE_TOLERANCE Then
        IsInExtensionRange = True
    ElseIf mHasTensionExtension And mHasPhysicalTensionLimit And _
            strain > mPhysicalTensionStrain + RANGE_TOLERANCE Then
        IsInExtensionRange = True
    End If
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
        Dim segmentIndex As Long
        segmentIndex = FindSegmentIndex(strain)
        If segmentIndex > 0 Then StressAtStrain = mSegmentTangent(segmentIndex) * strain + mSegmentIntercept(segmentIndex)
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
    Dim segmentIndex As Long
    segmentIndex = FindSegmentIndex(strain)
    If segmentIndex > 0 Then GetTangentModulus = mSegmentTangent(segmentIndex)
End Function

' Возвращает напряжение и касательный модуль за один проход по отрезкам.
' В горячих циклах CSectionSolver это убирает два независимых поиска одного
' и того же участка диаграммы для каждого бетонного волокна и стержня.
Public Sub EvaluateAtStrain(ByVal strain As Double, ByRef stress As Double, ByRef tangent As Double)
    EnsureInitialized

    If strain < mStrains(1) Then
        stress = mStresses(1)
        tangent = 0#
    ElseIf strain > mStrains(mPointCount) Then
        stress = mStresses(mPointCount)
        tangent = 0#
    Else
        Dim segmentIndex As Long
        segmentIndex = FindSegmentIndex(strain)
        If segmentIndex > 0 Then
            tangent = mSegmentTangent(segmentIndex)
            stress = mSegmentTangent(segmentIndex) * strain + mSegmentIntercept(segmentIndex)
        End If
    End If
End Sub

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

' Подготавливает быстрые коэффициенты линейных участков диаграммы.
Private Sub BuildSegmentCache()
    Dim i As Long
    ReDim mSegmentTangent(1 To mPointCount - 1)
    ReDim mSegmentIntercept(1 To mPointCount - 1)

    ' В горячих циклах solver многократно спрашивает одну и ту же кусочно-линейную
    ' диаграмму. Поэтому наклон и свободный член каждого участка считаются один раз
    ' при инициализации, а не заново для каждого бетонного волокна и стержня.
    For i = 1 To mPointCount - 1
        mSegmentTangent(i) = (mStresses(i + 1) - mStresses(i)) / (mStrains(i + 1) - mStrains(i))
        mSegmentIntercept(i) = mStresses(i) - mSegmentTangent(i) * mStrains(i)
    Next i
End Sub

' Находит участок диаграммы, которому принадлежит заданная деформация.
Private Function FindSegmentIndex(ByVal strain As Double) As Long
    Dim i As Long
    For i = 1 To mPointCount - 1
        If strain >= mStrains(i) And strain <= mStrains(i + 1) Then
            FindSegmentIndex = i
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

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Методы ниже оставлены как тестовый API. Рабочий расчет получает диаграммы
' через CMaterialModelProvider и основной InitializeFromArrays, а unit-тестам
' эти shortcut-конструкторы позволяют проверять интерполяцию без Excel и Config.

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

' ДЛЯ ТЕСТОВ: создает сжатую трехучастковую диаграмму с нулевой точкой.
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

' ДЛЯ ТЕСТОВ: создает симметричную растянутую двухточечную диаграмму.
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

' ДЛЯ ТЕСТОВ: создает симметричную растянутую трехточечную диаграмму.
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

' ДЛЯ ТЕСТОВ: совместимый короткий initializer для старых unit-тестов диаграммы.
Public Sub Initialize(ByVal eps1 As Double, ByVal stress1 As Double, ByVal eps2 As Double, _
        Optional ByVal stress2 As Variant)
    If IsMissing(stress2) Then
        InitializeFromTwoTensionPoints eps1, stress1, eps2, stress1
    Else
        InitializeFromTwoCompressionPoints eps1, stress1, eps2, CDbl(stress2)
    End If
End Sub

Public Property Get PointCount() As Long
    PointCount = mPointCount
End Property

' ДЛЯ ТЕСТОВ: возвращает деформацию контрольной точки диаграммы по индексу.
Public Function PointStrain(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3109, "CMaterialDiagram", "Индекс точки диаграммы вне диапазона."
    PointStrain = mStrains(index)
End Function

' ДЛЯ ТЕСТОВ: возвращает напряжение контрольной точки диаграммы по индексу.
Public Function PointStress(ByVal index As Long) As Double
    EnsureInitialized
    If index < 1 Or index > mPointCount Then Err.Raise vbObjectError + 3110, "CMaterialDiagram", "Индекс точки диаграммы вне диапазона."
    PointStress = mStresses(index)
End Function
