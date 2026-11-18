VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionProjectionStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Равноправные проекции одного продольного стержня на ось хомута.
' Площадь делится поровну между отдельными точками и связными дугами;
' на каждой непрерывной дуге её доля распределяется равномерно по длине.
' После Seal набор неизменяем; координата станции измеряется вдоль оси, мм.
' ==========================================================================
Private mLength As Double, mSealed As Boolean
Private mPoints() As Double, mPointCount As Long
Private mStarts() As Double, mEnds() As Double, mArcCount As Long
Private mComponents As Long

' Задаёт период замкнутой оси до добавления точек и интервалов.
Friend Sub Initialize(ByVal perimeter As Double)
    If mLength > 0# Or perimeter <= 0# Then ProjectionError "Не задан единственный положительный периметр проекций."
    mLength = perimeter
End Sub

' Добавляет одну станцию без повторного учёта общего конца соседних рёбер.
Friend Sub AddPoint(ByVal station As Double)
    RequireBuilding
    station = Normalize(station)
    Dim i As Long, gap As Double
    For i = 1 To mPointCount
        gap = Abs(mPoints(i) - station)
        If GeomMin(gap, mLength - gap) <= GEOM_BOUNDARY_TOLERANCE Then Exit Sub
    Next i
    mPointCount = mPointCount + 1
    ReDim Preserve mPoints(1 To mPointCount)
    mPoints(mPointCount) = station
End Sub

' Принимает непрерывный равноудалённый участок в порядке обхода оси.
' Соседние дуги объединяются независимо от разбиения исходного контура.
Friend Sub AddArc(ByVal firstStation As Double, ByVal lastStation As Double)
    RequireBuilding
    If lastStation <= firstStation Then ProjectionError "Равноудалённая дуга должна иметь положительную длину."
    If mArcCount > 0 Then
        If Abs(firstStation - mEnds(mArcCount)) <= GEOM_BOUNDARY_TOLERANCE Then
            mEnds(mArcCount) = lastStation
            Exit Sub
        End If
    End If
    mArcCount = mArcCount + 1
    ReDim Preserve mStarts(1 To mArcCount), mEnds(1 To mArcCount)
    mStarts(mArcCount) = firstStation: mEnds(mArcCount) = lastStation
End Sub

' Замыкает интервалы через шов и исключает точки, уже принадлежащие дугам.
Friend Sub Seal()
    RequireBuilding
    Dim i As Long, j As Long, firstEnd As Double, lastStart As Double
    If mArcCount > 1 Then
        If mStarts(1) <= GEOM_BOUNDARY_TOLERANCE And mEnds(mArcCount) >= mLength - GEOM_BOUNDARY_TOLERANCE Then
            firstEnd = mEnds(1): lastStart = mStarts(mArcCount)
            For i = 1 To mArcCount - 2
                mStarts(i) = mStarts(i + 1): mEnds(i) = mEnds(i + 1)
            Next i
            mArcCount = mArcCount - 1
            mStarts(mArcCount) = lastStart: mEnds(mArcCount) = firstEnd + mLength
        End If
    End If
    j = 0
    For i = 1 To mPointCount
        If Not OnAnyArc(mPoints(i)) Then j = j + 1: mPoints(j) = mPoints(i)
    Next i
    mPointCount = j: mComponents = mPointCount + mArcCount
    If mComponents = 0 Then ProjectionError "Не найдены ближайшие проекции стержня."
    mSealed = True
End Sub

' Возвращает число отдельных точечных проекций после объединения границ.
Friend Property Get PointCount() As Long
    RequireSealed
    PointCount = mPointCount
End Property

' Возвращает точечную станцию, нормированную на период оси.
Friend Property Get PointAt(ByVal index As Long) As Double
    RequireSealed
    PointAt = mPoints(index)
End Property

' Возвращает долю физической площади на одну отдельную точечную проекцию.
Friend Property Get PointWeight() As Double
    RequireSealed
    PointWeight = 1# / mComponents
End Property

' Возвращает число связных непрерывных дуговых проекций.
Friend Property Get ArcCount() As Long
    RequireSealed
    ArcCount = mArcCount
End Property

' Возвращает начало дугового интервала в координате вдоль оси.
Friend Property Get ArcStartAt(ByVal index As Long) As Double
    RequireSealed
    ArcStartAt = mStarts(index)
End Property

' Возвращает конец интервала; при пересечении шва он может превышать периметр.
Friend Property Get ArcEndAt(ByVal index As Long) As Double
    RequireSealed
    ArcEndAt = mEnds(index)
End Property

' Возвращает долю площади на единицу длины равноудалённой дуги, 1/мм.
Friend Function WeightDensityAt(ByVal station As Double) As Double
    RequireSealed
    Dim i As Long
    For i = 1 To mArcCount
        If OnArc(station, i) Then
            WeightDensityAt = 1# / (mComponents * (mEnds(i) - mStarts(i)))
            Exit Function
        End If
    Next i
End Function

' Проверяет принадлежность дуге с учётом периодического шва.
Private Function OnArc(ByVal station As Double, ByVal index As Long) As Boolean
    Dim offset As Double: offset = Normalize(station) - mStarts(index)
    If offset < -GEOM_BOUNDARY_TOLERANCE Then offset = offset + mLength
    OnArc = offset >= -GEOM_BOUNDARY_TOLERANCE And offset <= mEnds(index) - mStarts(index) + GEOM_BOUNDARY_TOLERANCE
End Function

' Исключает повтор площади, если точка совпала с концом непрерывной проекции.
Private Function OnAnyArc(ByVal station As Double) As Boolean
    Dim i As Long
    For i = 1 To mArcCount
        If OnArc(station, i) Then OnAnyArc = True: Exit Function
    Next i
End Function

' Сводит станцию к периоду и объединяет численно совпадающие концы оси.
Private Function Normalize(ByVal station As Double) As Double
    station = station - mLength * Int(station / mLength)
    If station <= GEOM_BOUNDARY_TOLERANCE Or mLength - station <= GEOM_BOUNDARY_TOLERANCE Then station = 0#
    Normalize = station
End Function

' Запрещает менять подготовленный набор проекций.
Private Sub RequireBuilding()
    If mLength <= 0# Or mSealed Then ProjectionError "Набор проекций не подготовлен к заполнению."
End Sub

' Запрещает читать незавершённое распределение площади.
Private Sub RequireSealed()
    If Not mSealed Then ProjectionError "Завершите набор проекций стержня."
End Sub

' Возвращает ошибку геометрической подготовки без привязки к Excel.
Private Sub ProjectionError(ByVal description As String)
    Err.Raise vbObjectError + 4505, TypeName(Me), description
End Sub
