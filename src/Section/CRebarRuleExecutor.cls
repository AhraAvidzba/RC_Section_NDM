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
Private mMergedPositions As Long, mConflicts As Long

Public Property Get RequestedPositions() As Long: RequestedPositions = mPositionCount + mSkippedProjections: End Property
Public Property Get CreatedPositions() As Long: CreatedPositions = mPositionCount - mMergedPositions: End Property
Public Property Get SkippedProjections() As Long: SkippedProjections = mSkippedProjections: End Property
Public Property Get MergedPositions() As Long: MergedPositions = mMergedPositions: End Property
Public Property Get Conflicts() As Long: Conflicts = mConflicts: End Property

Public Function Build(ByVal definition As CSectionDefinition, ByVal geometry As CGeometryRegion) As CRebarLayout
    If definition Is Nothing Or geometry Is Nothing Then RuleError 0, "Определение и бетонная область обязательны."
    If Not definition.IsSealed Then RuleError 0, "Определение еще не опубликовано."
    Set mDefinition = definition
    mPositionCount = 0: mCapacity = 0: mSkippedProjections = 0: Erase mPositions
    mMergedPositions = 0: mConflicts = 0
    Dim ruleCount As Long, i As Long, result As CRebarLayout
    ruleCount = definition.RuleCount
    Set result = New CRebarLayout
    If ruleCount = 0 Then Set Build = result: Exit Function
    ReDim mRules(1 To ruleCount): ReDim mState(1 To ruleCount)
    ReDim mStart(1 To ruleCount): ReDim mCount(1 To ruleCount)
    For i = 1 To ruleCount: definition.CopyRule i, mRules(i): Next i
    Preflight
    For i = 1 To ruleCount: PrepareRule i: Next i
    Dim order() As Long, scratch() As Long, index As Long, position As TRebarPosition
    Dim region As CConcreteRegion, query As CSectionGeometryQuery, owner As Long, identical As Boolean
    Set region = definition.ConcreteDomain: Set query = New CSectionGeometryQuery
    If mPositionCount > 0 Then
        ReDim order(1 To mPositionCount): ReDim scratch(1 To mPositionCount)
        For i = 1 To mPositionCount: order(i) = i: Next i
        SortOrder order, scratch, 1, mPositionCount
        For i = 1 To mPositionCount
            position = mPositions(order(i)): index = position.RuleIndex
            owner = result.FindBarAt(position.X, position.Y)
            If owner > 0 Then
                identical = (result.Diameter(owner) = mRules(index).Diameter And _
                    StrComp(result.SteelClass(owner), mRules(index).SteelClass, vbTextCompare) = 0)
                If mRules(index).Coincidence = RejectCoincidence Then _
                    RuleError index, "Позиция совпадает со стержнем '" & result.BarID(owner) & "'; политика объединения не задана."
                If Not identical And mRules(index).Coincidence = MergeIdentical Then _
                    RuleError index, "Совпавшие позиции имеют разные диаметры или материалы."
                ValidatePlacement index, result.X(owner), result.Y(owner), result.Diameter(owner), position.SourcePositionID, geometry, region, query
                If Not identical Then mConflicts = mConflicts + 1
                mMergedPositions = mMergedPositions + 1
                result.AddMembership owner, mRules(index).AnnotationGroup
                GoTo NextPosition
            End If
            ValidatePlacement index, position.X, position.Y, mRules(index).Diameter, position.SourcePositionID, geometry, region, query
            result.AddBar position.SourcePositionID, position.X, position.Y, mRules(index).Diameter, 0#, _
                mRules(index).SteelClass, mRules(index).Comment, geometry, mRules(index).AnnotationGroup
NextPosition:
        Next i
    End If
    Set Build = result
End Function

Private Sub ValidatePlacement(ByVal index As Long, ByVal x As Double, ByVal y As Double, ByVal diameter As Double, _
        ByVal positionID As String, ByVal geometry As CGeometryRegion, ByVal region As CConcreteRegion, ByVal query As CSectionGeometryQuery)
    If Not geometry.ContainsPoint(x, y) Then RuleError index, "Ось стержня '" & positionID & "' вне бетона."
    If mRules(index).PlacementCheck = CheckFullDisk Then
        If query.BoundaryDistance(region, x, y) + GEOM_TOLERANCE < diameter / 2# Then _
            RuleError index, "Стержень '" & positionID & "' не помещается целиком внутри бетона."
    End If
End Sub

' Счетчики всех независимых путей и зависимостей проверяются до массивов
' позиций. Для проекций верхняя граница включает еще не отфильтрованные лучи.
Private Sub Preflight()
    Dim state() As Long, counts() As Double, total As Double, i As Long
    ReDim state(1 To mDefinition.RuleCount)
    ReDim counts(1 To mDefinition.RuleCount)
    For i = 1 To mDefinition.RuleCount
        EstimateRule i, state, counts
        total = total + counts(i)
        If total > GEOM_MAX_REBAR_COUNT Then RuleError i, "Суммарное число запрошенных позиций превышает диапазон счетчиков; уменьшите количество или число активных рядов."
    Next i
End Sub

Private Sub EstimateRule(ByVal index As Long, ByRef state() As Long, ByRef counts() As Double)
    If state(index) = 2 Then Exit Sub
    If state(index) = 1 Then RuleError index, "Циклическая зависимость рядов."
    state(index) = 1
    If mRules(index).Enabled Then
        ValidateOptions index
        Dim source As Long, ids As Variant, id As Variant, path As CGeometryPath
        If mRules(index).Mode = ProjectFrom Or mRules(index).Mode = DependentOffset Then
            source = mDefinition.RuleIndex(mRules(index).SourceRuleID)
            EstimateRule source, state, counts
            counts(index) = counts(source)
            If mRules(index).EverySecond And mRules(source).Mode <= ExplicitPositions Then counts(index) = -Int(-counts(index) / 2#)
        Else
            ids = mDefinition.TargetPaths(mRules(index).Target)
            For Each id In ids
                Set path = mDefinition.Path(CStr(id))
                ValidateWorkInterval index, path
                counts(index) = counts(index) + PositionCountFor(index, path)
            Next id
        End If
        If counts(index) > GEOM_MAX_REBAR_COUNT Then RuleError index, "Количество позиций превышает диапазон счетчиков."
    End If
    state(index) = 2
End Sub

Private Function PositionCountFor(ByVal index As Long, ByVal path As CGeometryPath) As Double
    Dim intervals As Double, length As Double
    length = path.Length - mRules(index).StartSetback - mRules(index).EndSetback
    Select Case mRules(index).Mode
        Case ByCount
            PositionCountFor = mRules(index).Count
        Case ExplicitPositions
            If Not IsArray(mRules(index).ExplicitStations) Then RuleError index, "Не заданы явные станции."
            PositionCountFor = UBound(mRules(index).ExplicitStations) - LBound(mRules(index).ExplicitStations) + 1#
        Case ByMaxStep
            intervals = -Int(-length / mRules(index).Step)
            Select Case mRules(index).EndMode
                Case IncludeBoth: PositionCountFor = intervals + 1#
                Case ExcludeBoth: PositionCountFor = GeomMax(0#, intervals - 1#)
                Case Else: PositionCountFor = intervals
            End Select
        Case ByFixedStep
            intervals = length / mRules(index).Step
            If mRules(index).EndMode = IncludeBoth Or mRules(index).EndMode = ClosedUniform Then
                If Abs(intervals - Round(intervals)) > 0.0000000001 Then RuleError index, "Точный шаг несовместим с фиксированными обоими концами или полным замкнутым обходом."
                PositionCountFor = Round(intervals)
                If mRules(index).EndMode = IncludeBoth Then PositionCountFor = PositionCountFor + 1#
            Else
                PositionCountFor = Fix(intervals) + 1#
            End If
        Case Else
            RuleError index, "Неизвестный режим раскладки."
    End Select
    If PositionCountFor < 0# Or PositionCountFor > GEOM_MAX_REBAR_COUNT Then RuleError index, "Количество позиций превышает диапазон счетчиков."
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
    Dim position As TRebarPosition, first As Long
    ids = mDefinition.TargetPaths(mRules(index).Target)
    For Each id In ids
        Set path = mDefinition.Path(CStr(id))
        stations = StationsFor(index, path)
        If Not IsEmpty(stations) Then
            first = mPositionCount + 1
            For i = LBound(stations) To UBound(stations)
                ordinal = ordinal + 1
                path.AtDistance CDbl(stations(i)), position.X, position.Y, position.TangentX, position.TangentY, mRules(index).IncomingTangentAtJoint
                position.Station = CDbl(stations(i)): position.RuleIndex = index
                position.RootRuleIndex = index: position.RootOrdinal = ordinal: position.SourceOrdinal = ordinal
                position.RootPublicationRank = ordinal
                position.PathID = CStr(id)
                position.SourcePositionID = PositionID(index, ordinal, "")
                AppendPosition position
            Next i
            If mRules(index).PositionOrder = FromPathMiddle Then RankFromMiddle first, mPositionCount, path.Length
        End If
    Next id
End Sub

' Устойчивый O(n log n) порядок по расстоянию до середины полного пути.
' Геометрический ordinal остается прежним, в том числе при неравных setback.
Private Sub RankFromMiddle(ByVal first As Long, ByVal last As Long, ByVal length As Double)
    Dim order() As Long, scratch() As Long, i As Long, base As Long
    ReDim order(first To last): ReDim scratch(first To last)
    For i = first To last: order(i) = i: Next i
    SortMiddle order, scratch, first, last, length / 2#
    base = mPositions(first).RootOrdinal - 1
    For i = first To last: mPositions(order(i)).RootPublicationRank = base + i - first + 1: Next i
End Sub

Private Sub SortMiddle(ByRef order() As Long, ByRef scratch() As Long, ByVal first As Long, ByVal last As Long, ByVal middleStation As Double)
    If first >= last Then Exit Sub
    Dim middle As Long, a As Long, b As Long, out As Long, takeA As Boolean
    middle = first + (last - first) \ 2
    SortMiddle order, scratch, first, middle, middleStation
    SortMiddle order, scratch, middle + 1, last, middleStation
    a = first: b = middle + 1
    For out = first To last
        If a > middle Then
            takeA = False
        ElseIf b > last Then
            takeA = True
        Else
            takeA = (Abs(mPositions(order(a)).Station - middleStation) <= Abs(mPositions(order(b)).Station - middleStation))
        End If
        If takeA Then scratch(out) = order(a): a = a + 1 Else scratch(out) = order(b): b = b + 1
    Next out
    For out = first To last: order(out) = scratch(out): Next out
End Sub

Friend Function RuleEnds(ByVal ruleID As String, ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double) As Boolean
    Dim index As Long, last As Long
    index = mDefinition.RuleIndex(ruleID)
    If mCount(index) = 0 Then Exit Function
    last = mStart(index) + mCount(index) - 1
    startX = mPositions(mStart(index)).X: startY = mPositions(mStart(index)).Y
    endX = mPositions(last).X: endY = mPositions(last).Y
    RuleEnds = True
End Function

' Аннотация измеряет подготовленный ряд; зависимые дополнительные позиции
' не меняют его базовую привязку. Координаты стержней не пересчитываются.
Friend Function RuleBounds(ByVal ruleID As String, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double) As Boolean
    Dim index As Long, i As Long
    index = mDefinition.RuleIndex(ruleID)
    If mCount(index) = 0 Then Exit Function
    minX = 1E+100: minY = 1E+100: maxX = -1E+100: maxY = -1E+100
    For i = mStart(index) To mStart(index) + mCount(index) - 1
        minX = GeomMin(minX, mPositions(i).X): maxX = GeomMax(maxX, mPositions(i).X)
        minY = GeomMin(minY, mPositions(i).Y): maxY = GeomMax(maxY, mPositions(i).Y)
    Next i
    RuleBounds = True
End Function

' Привязка к самому удаленному ряду вдоль заданной нормали. Стержни на
' наклонных/дуговых участках не расширяют подпись прямого крайнего ряда.
Friend Function RuleExtremeEnds(ByVal ruleID As String, ByVal nx As Double, ByVal ny As Double, _
        ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double) As Boolean
    Dim index As Long, i As Long, projection As Double, best As Double, along As Double, first As Double, last As Double
    index = mDefinition.RuleIndex(ruleID): best = -1E+100
    For i = mStart(index) To mStart(index) + mCount(index) - 1
        projection = mPositions(i).X * nx + mPositions(i).Y * ny
        along = mPositions(i).X * ny - mPositions(i).Y * nx
        If Not RuleExtremeEnds Or projection > best + GEOM_TOLERANCE Then
            best = projection: first = along: last = along
            startX = mPositions(i).X: startY = mPositions(i).Y: endX = startX: endY = startY
            RuleExtremeEnds = True
        ElseIf Abs(projection - best) <= GEOM_TOLERANCE Then
            If along < first Then first = along: startX = mPositions(i).X: startY = mPositions(i).Y
            If along > last Then last = along: endX = mPositions(i).X: endY = mPositions(i).Y
        End If
    Next i
End Function

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
    Dim directionX As Double, directionY As Double, successful As Long
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
            directionX = mRules(index).DirectionX: directionY = mRules(index).DirectionY
            If mRules(index).SourceNormalSign <> 0# Then
                directionX = -position.TangentY * mRules(index).SourceNormalSign
                directionY = position.TangentX * mRules(index).SourceNormalSign
            End If
            For targetIndex = 1 To prepared.Count
                Set path = prepared(targetIndex)
                If path.RayHit(position.X, position.Y, directionX, directionY, station, ray) Then
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
                path.AtDistance bestStation, position.X, position.Y, position.TangentX, position.TangentY, mRules(index).IncomingTangentAtJoint
                position.Station = bestStation: position.PathID = bestID: position.RuleIndex = index
                position.SourceOrdinal = ordinal
                successful = successful + 1
                If mRules(index).SuccessfulProjectionOrdinal Then position.SourceOrdinal = successful
                If mRules(index).NewPublicationRoot Then
                    position.RootRuleIndex = index: position.RootOrdinal = successful: position.RootPublicationRank = successful
                End If
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
        If .PositionOrder <> AlongPath And .PositionOrder <> FromPathMiddle Then RuleError index, "Неизвестный порядок позиций."
        If .Coincidence < RejectCoincidence Or .Coincidence > KeepPublishedOwner Then RuleError index, "Неизвестная политика совпадений."
        If .RemainderPlacement < RemainderStart Or .RemainderPlacement > RemainderCenter Then RuleError index, "Неизвестное размещение остатка шага."
        If .Mode = ByMaxStep Or .Mode = ByFixedStep Then
            If .Step <= 0# Then RuleError index, "Шаг должен быть положительным."
        End If
        If .Mode = ProjectFrom Then
            If .SourceNormalSign <> -1# And .SourceNormalSign <> 0# And .SourceNormalSign <> 1# Then RuleError index, "Знак местной нормали должен быть -1, 0 или 1."
            If .SourceNormalSign = 0# And Sqr(.DirectionX * .DirectionX + .DirectionY * .DirectionY) <= GEOM_TOLERANCE Then RuleError index, "Не задано направление проекции."
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
    If Len(prefix) = 0 And Len(sourceID) > 0 And Len(mRules(index).SourceSuffix) > 0 Then
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
    ordinalA = mPositions(a).RootPublicationRank: ordinalB = mPositions(b).RootPublicationRank
    If mRules(keyA).Publication = PublishPerSourcePosition Then keyA = mPositions(a).RootRuleIndex
    If mRules(keyB).Publication = PublishPerSourcePosition Then keyB = mPositions(b).RootRuleIndex
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
