VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionLongitudinalDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' Распределение продольной тяги одноконтурной фермы между стержнями.
' Используются стержни внутри оси и в заданной полосе вблизи неё.
' Подготовленные коэффициенты зависят от геометрии, но не от момента T.
' =====================================================================

Private mStressPerForce() As Double ' Напряжение стержня в МПа на 1 Н суммарной продольной тяги.

' Распределяет периметр по соседним станциям. Площадь одного стержня
' делится между равноправными проекциями, а усилия складываются по его ID.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal hoop As CTorsionHoopGeometry, ByVal bandWidth As Double)
    If section Is Nothing Or hoop Is Nothing Then AllocationError "Не заданы сечение или ось замкнутого хомута."
    If bandWidth <= 0# Then AllocationError "Расчётная полоса продольной арматуры должна иметь положительную ширину."
    Dim nBars As Long: nBars = section.RebarCount
    If nBars < 3 Then AllocationError "Для распределения тяги кручения нужны продольные стержни минимум на трёх разных станциях контура."

    Dim perimeter As Double: perimeter = hoop.Length
    Dim sortedBars() As Long, stations() As Double, weights() As Double
    ReDim sortedBars(1 To nBars), stations(1 To nBars), weights(1 To nBars)
    Dim projections() As CTorsionProjectionStations, continuousBars() As Long, continuousCount As Long, allocated As Boolean
    Dim i As Long, j As Long, count As Long, selected As Long, x As Double, y As Double, station As Double, distance As Double, ambiguous As Boolean
    For i = 1 To nBars
        x = section.RebarX(i): y = section.RebarY(i)
        If hoop.ContainsInside(x, y) And hoop.IsOuterRowPoint(x, y) Then
            If Not hoop.TryProjectStation(x, y, station, distance, ambiguous) Then AllocationError "Недоступна проекция стержня " & section.RebarID(i) & "."
            If distance <= bandWidth + GEOM_BOUNDARY_TOLERANCE Then
                selected = selected + 1
                If ambiguous Then
                    If Not allocated Then
                        ReDim projections(1 To nBars), continuousBars(1 To nBars): allocated = True
                    End If
                    Set projections(i) = hoop.NearestProjectionStations(x, y, distance)
                    For j = 1 To projections(i).PointCount
                        InsertStation sortedBars, stations, weights, count, i, projections(i).PointAt(j), projections(i).PointWeight
                    Next j
                    If projections(i).ArcCount > 0 Then
                        continuousCount = continuousCount + 1: continuousBars(continuousCount) = i
                    End If
                Else
                    InsertStation sortedBars, stations, weights, count, i, station, 1#
                End If
            End If
        End If
    Next i
    If selected < 3 Then AllocationError "В расчётной полосе хомута найдено меньше трёх продольных стержней."

    Dim firstOfGroup() As Long, groupStation() As Double, groups As Long
    ReDim firstOfGroup(1 To count + 1), groupStation(0 To count) ' Нулевая граница сохраняет допустимый пустой набор при одних дугах.
    If count > 0 Then groups = 1: firstOfGroup(1) = 1: groupStation(1) = stations(1)
    For i = 2 To count
        If stations(i) - groupStation(groups) > GEOM_BOUNDARY_TOLERANCE Then
            groups = groups + 1: firstOfGroup(groups) = i: groupStation(groups) = stations(i)
        End If
    Next i
    firstOfGroup(groups + 1) = count + 1
    If groups < 3 And continuousCount = 0 Then AllocationError "Продольная арматура охватывает меньше трёх разных станций хомута."

    Dim result() As Double: ReDim result(1 To nBars)
    Dim g As Long, previous As Double, nextStation As Double, fraction As Double
    Dim groupArea() As Double, groupWidth() As Double
    ReDim groupArea(0 To groups), groupWidth(0 To groups)
    If groups > 0 Then previous = groupStation(groups) - perimeter ' Станция слева от шва в развёрнутой координате.
    For g = 1 To groups
        nextStation = groupStation(1) + perimeter
        If g < groups Then nextStation = groupStation(g + 1)
        fraction = (nextStation - previous) / (2# * perimeter)
        previous = groupStation(g)
        groupWidth(g) = fraction * perimeter
        groupArea(g) = 0#
        For j = firstOfGroup(g) To firstOfGroup(g + 1) - 1
            groupArea(g) = groupArea(g) + section.RebarArea(sortedBars(j)) * weights(j)
        Next j
        If groupArea(g) <= 0# Then AllocationError "Площадь продольной арматуры на станции хомута должна быть больше нуля."
        If continuousCount = 0 Then
            For j = firstOfGroup(g) To firstOfGroup(g + 1) - 1
                i = sortedBars(j)
                result(i) = result(i) + fraction * weights(j) / groupArea(g)
            Next j
        End If
    Next g
    If continuousCount > 0 Then DistributeContinuous section, perimeter, sortedBars, stations, weights, firstOfGroup, groupStation, _
        groupArea, groupWidth, groups, projections, continuousBars, continuousCount, result
    Dim sumFraction As Double
    For i = 1 To nBars
        sumFraction = sumFraction + result(i) * section.RebarArea(i)
    Next i
    If Abs(sumFraction - 1#) > 1E-10 Then AllocationError "Доли продольной тяги не покрывают полный контур хомута."
    mStressPerForce = result
End Sub

' Вставляет станцию в отсортированную часть подготовительного массива.
Private Sub InsertStation(ByRef order() As Long, ByRef stations() As Double, ByRef weights() As Double, ByRef count As Long, _
        ByVal rebarIndex As Long, ByVal station As Double, ByVal weight As Double)
    If count = UBound(order) Then
        Dim capacity As Long: capacity = 2 * UBound(order)
        ReDim Preserve order(1 To capacity), stations(1 To capacity), weights(1 To capacity)
    End If
    Dim index As Long: index = count + 1
    Do While index > 1
        If stations(index - 1) <= station Then Exit Do
        stations(index) = stations(index - 1): order(index) = order(index - 1): weights(index) = weights(index - 1)
        index = index - 1
    Loop
    stations(index) = station: order(index) = rebarIndex: weights(index) = weight
    count = count + 1
End Sub

' Распределяет тягу при непрерывных дуговых проекциях. В каждой ячейке
' точечной станции её площадь равномерно отнесена к прилегающей длине;
' дуговая доля площади распределена по своей дуге. Интегрирование выполняется
' точно по участкам постоянной плотности, без выборки и зависимости от рёбер.
Private Sub DistributeContinuous(ByVal section As CSectionModel, ByVal perimeter As Double, ByRef order() As Long, _
        ByRef stations() As Double, ByRef weights() As Double, ByRef firstOfGroup() As Long, ByRef groupStation() As Double, _
        ByRef groupArea() As Double, ByRef groupWidth() As Double, ByVal groups As Long, _
        ByRef projections() As CTorsionProjectionStations, ByRef continuousBars() As Long, ByVal continuousCount As Long, ByRef result() As Double)
    Dim boundaries() As Double, boundaryCount As Long, i As Long, j As Long, g As Long, bar As Long, nextStation As Double
    ReDim boundaries(1 To 2): boundaries(1) = 0#: boundaries(2) = perimeter: boundaryCount = 2
    For g = 1 To groups
        nextStation = groupStation(1) + perimeter
        If g < groups Then nextStation = groupStation(g + 1)
        InsertBoundary boundaries, boundaryCount, 0.5 * (groupStation(g) + nextStation), perimeter
    Next g
    For i = 1 To continuousCount
        bar = continuousBars(i)
        For j = 1 To projections(bar).ArcCount
            InsertBoundary boundaries, boundaryCount, projections(bar).ArcStartAt(j), perimeter
            InsertBoundary boundaries, boundaryCount, projections(bar).ArcEndAt(j), perimeter
        Next j
    Next i
    Dim densityWeights() As Double: ReDim densityWeights(1 To continuousCount)
    Dim midpoint As Double, density As Double, pieceFactor As Double, length As Double
    For j = 1 To boundaryCount - 1
        length = boundaries(j + 1) - boundaries(j): midpoint = 0.5 * (boundaries(j + 1) + boundaries(j))
        g = NearestGroup(midpoint, groupStation, groups, perimeter)
        density = 0#
        If g > 0 Then density = groupArea(g) / groupWidth(g)
        For i = 1 To continuousCount
            bar = continuousBars(i): densityWeights(i) = projections(bar).WeightDensityAt(midpoint)
            density = density + section.RebarArea(bar) * densityWeights(i)
        Next i
        If density <= 0# Then AllocationError "Выбранная продольная арматура не покрывает периметр оси хомута."
        pieceFactor = length / (perimeter * density)
        If g > 0 Then
            For i = firstOfGroup(g) To firstOfGroup(g + 1) - 1
                bar = order(i): result(bar) = result(bar) + pieceFactor * weights(i) / groupWidth(g)
            Next i
        End If
        For i = 1 To continuousCount
            bar = continuousBars(i): result(bar) = result(bar) + pieceFactor * densityWeights(i)
        Next i
    Next j
End Sub

' Добавляет границу участка на периоде; шов 0/perimeter остаётся явным.
Private Sub InsertBoundary(ByRef values() As Double, ByRef count As Long, ByVal value As Double, ByVal perimeter As Double)
    value = value - perimeter * Int(value / perimeter)
    Dim i As Long
    For i = 1 To count
        If Abs(values(i) - value) <= GEOM_BOUNDARY_TOLERANCE Then Exit Sub
    Next i
    count = count + 1: ReDim Preserve values(1 To count): i = count
    Do While i > 1
        If values(i - 1) <= value Then Exit Do
        values(i) = values(i - 1): i = i - 1
    Loop
    values(i) = value
End Sub

' Находит ближайшую точечную группу по длине оси двоичным поиском.
Private Function NearestGroup(ByVal station As Double, ByRef values() As Double, ByVal count As Long, ByVal perimeter As Double) As Long
    If count = 0 Then Exit Function
    Dim low As Long, high As Long, middle As Long, previous As Long, following As Long
    low = 1: high = count + 1
    Do While low < high
        middle = (low + high) \ 2
        If middle <= count Then
            If values(middle) < station Then low = middle + 1 Else high = middle
        Else
            high = middle
        End If
    Loop
    following = low: If following > count Then following = 1
    previous = low - 1: If previous < 1 Then previous = count
    Dim leftGap As Double, rightGap As Double
    leftGap = Abs(station - values(previous)): leftGap = GeomMin(leftGap, perimeter - leftGap)
    rightGap = Abs(station - values(following)): rightGap = GeomMin(rightGap, perimeter - rightGap)
    If leftGap <= rightGap Then NearestGroup = previous Else NearestGroup = following
End Function

' Копирует коэффициенты напряжений в порядке исходных стержней.
Friend Sub CopyStressPerUnitForce(ByRef coefficients() As Double)
    coefficients = mStressPerForce
End Sub

' Возвращает понятную причину неприменимости распределения.
Private Sub AllocationError(ByVal description As String)
    Err.Raise vbObjectError + 4505, TypeName(Me), description
End Sub
