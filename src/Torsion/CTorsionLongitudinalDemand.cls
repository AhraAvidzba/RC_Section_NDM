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

' Распределяет периметр по серединам дуговых расстояний между соседними
' станциями. Стержни на общей станции получают доли пропорционально площади.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal hoop As CTorsionHoopGeometry, ByVal bandWidth As Double)
    If section Is Nothing Or hoop Is Nothing Then AllocationError "Не заданы сечение или ось замкнутого хомута."
    If bandWidth <= 0# Then AllocationError "Расчётная полоса продольной арматуры должна иметь положительную ширину."
    Dim nBars As Long: nBars = section.RebarCount
    If nBars < 3 Then AllocationError "Для распределения тяги кручения нужны продольные стержни минимум на трёх разных станциях контура."

    Dim perimeter As Double: perimeter = hoop.Length
    Dim sortedBars() As Long, stations() As Double
    ReDim sortedBars(1 To nBars), stations(1 To nBars)
    Dim i As Long, count As Long, x As Double, y As Double, station As Double, distance As Double, ambiguous As Boolean
    For i = 1 To nBars
        x = section.RebarX(i): y = section.RebarY(i)
        If hoop.ContainsInside(x, y) And hoop.IsOuterRowPoint(x, y) Then
            If Not hoop.TryProjectStation(x, y, station, distance, ambiguous) Then AllocationError "Недоступна проекция стержня " & section.RebarID(i) & "."
            If distance <= bandWidth + GEOM_BOUNDARY_TOLERANCE Then
                If ambiguous Then AllocationError "Стержень " & section.RebarID(i) & _
                    " одинаково близок к разным участкам хомута. Уточните ось хомута или положение стержня."
                InsertStation sortedBars, stations, count, i, station
            End If
        End If
    Next i
    If count < 3 Then AllocationError "В расчётной полосе хомута найдено меньше трёх продольных стержней."

    Dim firstOfGroup() As Long, groupStation() As Double, groups As Long
    ReDim firstOfGroup(1 To count + 1), groupStation(1 To count)
    groups = 1: firstOfGroup(1) = 1: groupStation(1) = stations(1)
    For i = 2 To count
        If stations(i) - groupStation(groups) > GEOM_BOUNDARY_TOLERANCE Then
            groups = groups + 1: firstOfGroup(groups) = i: groupStation(groups) = stations(i)
        End If
    Next i
    firstOfGroup(groups + 1) = count + 1
    If groups < 3 Then AllocationError "Продольная арматура охватывает меньше трёх разных станций хомута."

    Dim result() As Double: ReDim result(1 To nBars)
    Dim g As Long, previous As Double, nextStation As Double, fraction As Double, area As Double, j As Long
    Dim sumFraction As Double
    previous = groupStation(groups) - perimeter ' Станция слева от шва в развёрнутой координате.
    For g = 1 To groups
        nextStation = groupStation(1) + perimeter
        If g < groups Then nextStation = groupStation(g + 1)
        fraction = (nextStation - previous) / (2# * perimeter)
        previous = groupStation(g)
        area = 0#
        For j = firstOfGroup(g) To firstOfGroup(g + 1) - 1
            area = area + section.RebarArea(sortedBars(j))
        Next j
        If area <= 0# Then AllocationError "Площадь продольной арматуры на станции хомута должна быть больше нуля."
        For j = firstOfGroup(g) To firstOfGroup(g + 1) - 1
            i = sortedBars(j)
            result(i) = fraction / area
            sumFraction = sumFraction + result(i) * section.RebarArea(i)
        Next j
    Next g
    If Abs(sumFraction - 1#) > 1E-10 Then AllocationError "Доли продольной тяги не покрывают полный контур хомута."
    mStressPerForce = result
End Sub

' Вставляет станцию в отсортированную часть подготовительного массива.
Private Sub InsertStation(ByRef order() As Long, ByRef stations() As Double, ByRef count As Long, _
        ByVal rebarIndex As Long, ByVal station As Double)
    Dim index As Long: index = count + 1
    Do While index > 1
        If stations(index - 1) <= station Then Exit Do
        stations(index) = stations(index - 1): order(index) = order(index - 1)
        index = index - 1
    Loop
    stations(index) = station: order(index) = rebarIndex
    count = count + 1
End Sub

' Копирует коэффициенты напряжений в порядке исходных стержней.
Friend Sub CopyStressPerUnitForce(ByRef coefficients() As Double)
    coefficients = mStressPerForce
End Sub

' Возвращает понятную причину неприменимости распределения.
Private Sub AllocationError(ByVal description As String)
    Err.Raise vbObjectError + 4505, TypeName(Me), description
End Sub
