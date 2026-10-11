VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarRuleExecutor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Универсальное исполнение правил: аналитические станции, зависимости,
' конечные проекции и отдельный порядок публикации физических стержней.
Private mDefinition As CSectionDefinition
Private mRules() As TRebarRule, mState() As Long, mStart() As Long, mCount() As Long
Private mPositions() As TRebarPosition, mPositionCount As Long, mCapacity As Long
Private mSkippedProjections As Long

Public Property Get RequestedPositions() As Long: RequestedPositions = mPositionCount + mSkippedProjections: End Property
Public Property Get CreatedPositions() As Long: CreatedPositions = mPositionCount: End Property
Public Property Get SkippedProjections() As Long: SkippedProjections = mSkippedProjections: End Property

Public Function Build(ByVal definition As CSectionDefinition, ByVal geometry As CGeometryRegion) As CRebarLayout
    If definition Is Nothing Or geometry Is Nothing Then RuleError 0, "Определение и бетонная область обязательны."
    If Not definition.IsSealed Then RuleError 0, "Определение еще не опубликовано."
    Set mDefinition = definition
    mPositionCount = 0: mCapacity = 0: mSkippedProjections = 0: Erase mPositions
    Dim ruleCount As Long, i As Long, result As CRebarLayout
    ruleCount = definition.RuleCount
    Set result = New CRebarLayout
    If ruleCount = 0 Then Set Build = result: Exit Function
    ReDim mRules(1 To ruleCount): ReDim mState(1 To ruleCount)
    ReDim mStart(1 To ruleCount): ReDim mCount(1 To ruleCount)
    For i = 1 To ruleCount: definition.CopyRule i, mRules(i): Next i
    For i = 1 To ruleCount: PrepareRule i: Next i
    Dim order() As Long, scratch() As Long, index As Long, position As TRebarPosition
    Dim region As CConcreteRegion, query As CSectionGeometryQuery
    Set region = definition.ConcreteDomain: Set query = New CSectionGeometryQuery
    If mPositionCount > 0 Then
        ReDim order(1 To mPositionCount): ReDim scratch(1 To mPositionCount)
        For i = 1 To mPositionCount: order(i) = i: Next i
        SortOrder order, scratch, 1, mPositionCount
        For i = 1 To mPositionCount
            position = mPositions(order(i)): index = position.RuleIndex
            If Not geometry.ContainsPoint(position.X, position.Y) Then RuleError index, "Ось стержня '" & position.SourcePositionID & "' вне бетона."
            If mRules(index).PlacementCheck = CheckFullDisk Then
                If query.BoundaryDistance(region, position.X, position.Y) + GEOM_TOLERANCE < mRules(index).Diameter / 2# Then _
                    RuleError index, "Стержень '" & position.SourcePositionID & "' не помещается целиком внутри бетона."
            End If
            result.AddBar position.SourcePositionID, position.X, position.Y, mRules(index).Diameter, 0#, _
                mRules(index).SteelClass, mRules(index).Comment, geometry, mRules(index).AnnotationGroup
        Next i
    End If
    Set Build = result
End Function

' Topological preparation сохраняет исходный геометрический ordinal;
' публикация и пропущенные проекции не меняют EverySecond.
Private Sub PrepareRule(ByVal index As Long)
    If mState(index) = 2 Then Exit Sub
    If mState(index) = 1 Then RuleError index, "Циклическая зависимость рядов."
    mState(index) = 1
    If Not mRules(index).Enabled Then mState(index) = 2: Exit Sub
    ValidateOptions index
    Dim source As Long
    If mRules(index).Mode = DependentOffset Or mRules(index).Mode = ProjectFrom Then
        source = mDefinition.RuleIndex(mRules(index).SourceRuleID)
        PrepareRule source
    End If
    mStart(index) = mPositionCount + 1
    Select Case mRules(index).Mode
        Case ByCount, ByMaxStep, ByFixedStep, ExplicitPositions
            PrepareIndependent index
        Case DependentOffset
            PrepareDependent index, source
        Case ProjectFrom
            PrepareProjected index, source
        Case Else
            RuleError index, "Неизвестный режим раскладки."
    End Select
    mCount(index) = mPositionCount - mStart(index) + 1
    mState(index) = 2
End Sub

Private Sub PrepareIndependent(ByVal index As Long)
    Dim ids As Variant, id As Variant, path As CGeometryPath, stations As Variant, ordinal As Long, i As Long
    Dim position As TRebarPosition
    ids = mDefinition.TargetPaths(mRules(index).Target)
    For Each id In ids
        Set path = mDefinition.Path(CStr(id))
        stations = StationsFor(index, path)
        If Not IsEmpty(stations) Then
            For i = LBound(stations) To UBound(stations)
                ordinal = ordinal + 1
                path.AtDistance CDbl(stations(i)), position.X, position.Y, position.TangentX, position.TangentY
                position.Station = CDbl(stations(i)): position.RuleIndex = index
                position.RootRuleIndex = index: position.RootOrdinal = ordinal: position.SourceOrdinal = ordinal
                position.PathID = CStr(id)
                position.SourcePositionID = PositionID(index, ordinal, "")
                AppendPosition position
            Next i
        End If
    Next id
End Sub

Private Sub PrepareDependent(ByVal index As Long, ByVal source As Long)
    Dim i As Long, ordinal As Long, position As TRebarPosition
    For i = mStart(source) To mStart(source) + mCount(source) - 1
        ordinal = ordinal + 1
        position = mPositions(i)
        If Not mRules(index).EverySecond Or position.SourceOrdinal Mod 2 = 1 Then
            position.X = position.X + mRules(index).TangentOffset * position.TangentX - mRules(index).NormalOffset * position.TangentY
            position.Y = position.Y + mRules(index).TangentOffset * position.TangentY + mRules(index).NormalOffset * position.TangentX
            position.RuleIndex = index
            position.SourcePositionID = PositionID(index, ordinal, mPositions(i).SourcePositionID)
            AppendPosition position
        End If
    Next i
End Sub

' Выбор ближайшего положительного пересечения только на заявленных путях.
' Setback фильтрует hit; полученные позиции не перераспределяются.
Private Sub PrepareProjected(ByVal index As Long, ByVal source As Long)
    Dim ids As Variant, id As Variant, path As CGeometryPath, i As Long, ordinal As Long, targetIndex As Long
    Dim prepared As Collection
    Dim position As TRebarPosition, station As Double, ray As Double, bestRay As Double, bestStation As Double, bestID As String
    ids = mDefinition.TargetPaths(mRules(index).Target)
    Set prepared = New Collection
    For Each id In ids
        Set path = mDefinition.Path(CStr(id))
        ValidateWorkInterval index, path
        If mRules(index).StartSetback <> 0# Or mRules(index).EndSetback <> 0# Then _
            Set path = path.Slice(mRules(index).StartSetback, path.Length - mRules(index).EndSetback)
        prepared.Add path
    Next id
    For i = mStart(source) To mStart(source) + mCount(source) - 1
        ordinal = ordinal + 1
        position = mPositions(i)
        If Not mRules(index).EverySecond Or position.SourceOrdinal Mod 2 = 1 Then
            bestID = "": bestRay = 1E+100
            For targetIndex = 1 To prepared.Count
                Set path = prepared(targetIndex)
                If path.RayHit(position.X, position.Y, mRules(index).DirectionX, mRules(index).DirectionY, station, ray) Then
                    If ray < bestRay Then
                        bestRay = ray: bestStation = station + mRules(index).StartSetback: bestID = path.ID
                    End If
                End If
            Next targetIndex
            If Len(bestID) = 0 Then
                If Not mRules(index).SkipMissingProjection Then RuleError index, "Нет пересечения исходной позиции с целевым рабочим путем."
                mSkippedProjections = mSkippedProjections + 1
            Else
                Set path = mDefinition.Path(bestID)
                path.AtDistance bestStation, position.X, position.Y, position.TangentX, position.TangentY
                position.Station = bestStation: position.PathID = bestID: position.RuleIndex = index
                position.SourceOrdinal = ordinal
                position.SourcePositionID = PositionID(index, ordinal, mPositions(i).SourcePositionID)
                AppendPosition position
            End If
        End If
    Next i
End Sub

Private Function StationsFor(ByVal index As Long, ByVal path As CGeometryPath) As Variant
    ValidateWorkInterval index, path
    Dim length As Double, first As Double, last As Double, count As Double, intervals As Double
    Dim stations() As Double, i As Long, start As Double, step As Double, kept As Long
    first = mRules(index).StartSetback: last = path.Length - mRules(index).EndSetback
    length = last - first
    If mRules(index).Mode = ExplicitPositions Then
        If Not IsArray(mRules(index).ExplicitStations) Then RuleError index, "Не заданы явные станции."
        count = UBound(mRules(index).ExplicitStations) - LBound(mRules(index).ExplicitStations) + 1
        ReDim stations(0 To CLng(count) - 1)
        For i = 0 To CLng(count) - 1
            stations(i) = CDbl(mRules(index).ExplicitStations(LBound(mRules(index).ExplicitStations) + i))
            If stations(i) < first Or stations(i) > last Then RuleError index, "Явная станция вне целевого рабочего участка."
            If path.IsClosed And stations(i) = path.Length Then RuleError index, "Последняя станция замкнутого пути дублирует первую."
        Next i
        StationsFor = stations: Exit Function
    End If
    If length <= GEOM_TOLERANCE Then
        If mRules(index).Mode <> ByCount Or mRules(index).Count <> 1 Then RuleError index, "Вырожденный рабочий участок допускает только один явно заданный стержень."
    End If
    If mRules(index).Mode = ByFixedStep Then
        StationsFor = FixedStations(index, path, first, length): Exit Function
    End If
    count = mRules(index).Count
    If mRules(index).Mode = ByMaxStep Then
        intervals = -Int(-length / mRules(index).Step)
        Select Case mRules(index).EndMode
            Case IncludeBoth: count = intervals + 1#
            Case ExcludeBoth: count = GeomMax(0#, intervals - 1#)
            Case Else: count = intervals
        End Select
    End If
    If count < 0# Or count > GEOM_MAX_REBAR_COUNT Then RuleError index, "Количество позиций превышает диапазон счетчиков."
    If count = 0# Then Exit Function
    ReDim stations(0 To CLng(count) - 1)
    If count = 1# And mRules(index).Mode = ByCount And mRules(index).EndMode <> ClosedUniform Then
        stations(0) = first + length / 2#
    Else
        Select Case mRules(index).EndMode
            Case IncludeBoth
                step = length / (count - 1#): start = first
            Case ExcludeBoth
                step = length / (count + 1#): start = first + step
            Case IncludeStart, ClosedUniform
                step = length / count: start = first
            Case IncludeEnd
                step = length / count: start = first + step
        End Select
        For i = 0 To CLng(count) - 1: stations(i) = start + CDbl(i) * step: Next i
    End If
    StationsFor = stations
End Function

Private Function FixedStations(ByVal index As Long, ByVal path As CGeometryPath, ByVal first As Double, ByVal length As Double) As Variant
    Dim count As Double, intervals As Double, start As Double, step As Double, stations() As Double, i As Long, kept As Long, station As Double
    step = mRules(index).Step: intervals = length / step
    If mRules(index).EndMode = ClosedUniform Or mRules(index).EndMode = IncludeBoth Then
        If Abs(intervals - Round(intervals)) > 0.0000000001 Then RuleError index, "Точный шаг несовместим с фиксированными обоими концами или полным замкнутым обходом."
        count = Round(intervals)
        If mRules(index).EndMode = IncludeBoth Then count = count + 1#
    Else
        count = Fix(intervals) + 1#
    End If
    If count > GEOM_MAX_REBAR_COUNT Then RuleError index, "Количество позиций превышает диапазон счетчиков."
    If count = 0# Then Exit Function
    ReDim stations(0 To CLng(count) - 1)
    start = first
    Select Case mRules(index).EndMode
        Case IncludeEnd: start = first + length - (count - 1#) * step
        Case ExcludeBoth
            Select Case mRules(index).RemainderPlacement
                Case RemainderStart: start = first + length - (count - 1#) * step
                Case RemainderCenter: start = first + (length - (count - 1#) * step) / 2#
            End Select
    End Select
    For i = 0 To CLng(count) - 1
        station = start + CDbl(i) * step
        If Not ((mRules(index).EndMode = ExcludeBoth Or mRules(index).EndMode = IncludeEnd) And station <= first + GEOM_TOLERANCE) And _
            Not ((mRules(index).EndMode = ExcludeBoth Or mRules(index).EndMode = IncludeStart) And station >= first + length - GEOM_TOLERANCE) Then
            stations(kept) = station: kept = kept + 1
        End If
    Next i
    If kept = 0 Then Exit Function
    ReDim Preserve stations(0 To kept - 1): FixedStations = stations
End Function

Private Sub ValidateOptions(ByVal index As Long)
    With mRules(index)
        If .Diameter <= 0# Then RuleError index, "Диаметр активного ряда должен быть положительным."
        If .StartSetback < 0# Or .EndSetback < 0# Then RuleError index, "Концевые отступы не могут быть отрицательными."
        If .EndMode < IncludeBoth Or .EndMode > ClosedUniform Then RuleError index, "Неизвестное правило концов."
        If .PlacementCheck <> CheckCenter And .PlacementCheck <> CheckFullDisk Then RuleError index, "Неизвестная проверка размещения."
        If .Publication <> PublishByRule And .Publication <> PublishPerSourcePosition Then RuleError index, "Неизвестный порядок публикации."
        If .RemainderPlacement < RemainderStart Or .RemainderPlacement > RemainderCenter Then RuleError index, "Неизвестное размещение остатка шага."
        If .Mode = ByMaxStep Or .Mode = ByFixedStep Then
            If .Step <= 0# Then RuleError index, "Шаг должен быть положительным."
        End If
        If .Mode = ProjectFrom Then
            If Sqr(.DirectionX * .DirectionX + .DirectionY * .DirectionY) <= GEOM_TOLERANCE Then RuleError index, "Не задано направление проекции."
        End If
    End With
End Sub

Private Sub ValidateWorkInterval(ByVal index As Long, ByVal path As CGeometryPath)
    If mRules(index).StartSetback + mRules(index).EndSetback > path.Length + GEOM_TOLERANCE Then RuleError index, "Концевые отступы превышают длину пути."
    If path.IsClosed Then
        If mRules(index).StartSetback <> 0# Or mRules(index).EndSetback <> 0# Then RuleError index, "Для полного замкнутого пути концевые отступы неприменимы."
        If mRules(index).EndMode <> ClosedUniform Then RuleError index, "Замкнутый путь требует ClosedUniform; для иных концов задайте открытый участок."
    ElseIf mRules(index).EndMode = ClosedUniform Then
        RuleError index, "ClosedUniform требует замкнутого пути."
    End If
End Sub

Private Function PositionID(ByVal index As Long, ByVal ordinal As Long, ByVal sourceID As String) As String
    Dim prefix As String
    prefix = mRules(index).SourcePrefix
    If Len(prefix) = 0 And Len(sourceID) > 0 Then
        PositionID = sourceID & mRules(index).SourceSuffix
    Else
        If Len(prefix) = 0 Then prefix = mRules(index).RuleID & "."
        PositionID = prefix & CStr(ordinal) & mRules(index).SourceSuffix
    End If
End Function

Private Sub AppendPosition(ByRef value As TRebarPosition)
    If mPositionCount >= GEOM_MAX_REBAR_COUNT Then RuleError value.RuleIndex, "Суммарное количество позиций превышает диапазон счетчиков."
    If mPositionCount = mCapacity Then
        mCapacity = CLng(GeomMin(GEOM_MAX_REBAR_COUNT, GeomMax(16#, CDbl(mCapacity) * 2#)))
        ReDim Preserve mPositions(1 To mCapacity)
    End If
    mPositionCount = mPositionCount + 1: mPositions(mPositionCount) = value
End Sub

' Устойчивый merge sort: порядок публикации не влияет на построение ряда.
Private Sub SortOrder(ByRef order() As Long, ByRef scratch() As Long, ByVal first As Long, ByVal last As Long)
    If first >= last Then Exit Sub
    Dim middle As Long, a As Long, b As Long, out As Long
    middle = first + (last - first) \ 2
    SortOrder order, scratch, first, middle: SortOrder order, scratch, middle + 1, last
    a = first: b = middle + 1
    For out = first To last
        If a > middle Then
            scratch(out) = order(b): b = b + 1
        ElseIf b > last Then
            scratch(out) = order(a): a = a + 1
        ElseIf BeforePosition(order(a), order(b)) Then
            scratch(out) = order(a): a = a + 1
        Else
            scratch(out) = order(b): b = b + 1
        End If
    Next out
    For out = first To last: order(out) = scratch(out): Next out
End Sub

Private Function BeforePosition(ByVal a As Long, ByVal b As Long) As Boolean
    Dim keyA As Long, keyB As Long, ordinalA As Long, ordinalB As Long
    keyA = mPositions(a).RuleIndex: keyB = mPositions(b).RuleIndex
    If mRules(keyA).Publication = PublishPerSourcePosition Then keyA = mPositions(a).RootRuleIndex: ordinalA = mPositions(a).RootOrdinal
    If mRules(keyB).Publication = PublishPerSourcePosition Then keyB = mPositions(b).RootRuleIndex: ordinalB = mPositions(b).RootOrdinal
    If keyA <> keyB Then BeforePosition = (keyA < keyB): Exit Function
    If ordinalA <> ordinalB Then BeforePosition = (ordinalA < ordinalB): Exit Function
    If mPositions(a).RuleIndex <> mPositions(b).RuleIndex Then BeforePosition = (mPositions(a).RuleIndex < mPositions(b).RuleIndex): Exit Function
    BeforePosition = (a <= b)
End Function

Private Sub RuleError(ByVal index As Long, ByVal reason As String)
    Dim reference As String
    If index > 0 Then reference = mRules(index).SourceReference: reason = "Ряд '" & mRules(index).RuleID & "': " & reason
    SectionRecipeError "CRebarRuleExecutor", reason, reference
End Sub
