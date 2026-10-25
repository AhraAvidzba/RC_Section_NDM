VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Готовое описание геометрии, армирования и аннотаций сечения
' ==========================================================================
' Хранит линии, роли бетонных границ, правила размещения арматуры и связи
' аннотаций. Заполняется через CSectionDefinitionBuilder и после Seal не изменяется.
' Геометрию проверяет CSectionGeometryQuery; фактическое размещение стержней
' проверяет CRebarRuleExecutor при сборке. Ссылки на Excel не сохраняются.

Private mTypeID As String, mVersion As Long, mSealed As Boolean ' Имя типа сечения, версия описания и признак завершения его заполнения.
Private mPaths As Object, mGroups As Object, mAnchors As Object, mSources As Object ' Именованные линии, группы, точки и адреса ввода.
Private mLoops() As TSectionLoop, mLoopCount As Long ' Описание наружных границ и отверстий, их принадлежность и число записей.
Private mRules() As TRebarRule, mRuleCount As Long ' Копии правил армирования и их количество; координаты стержней рассчитывает CRebarRuleExecutor.
Private mAnnotations() As TSectionAnnotation, mAnnotationCount As Long ' Описания размеров и подписей, их количество; сами стержни хранятся отдельно.
Private mRegion As CConcreteRegion ' Бетонная область, проверенная при Seal; при чтении возвращается её независимая копия.
Private mDiagnostic As CGeometryDiagnostic ' Последняя адресная ошибка проверки определения.
Private mRebarContours() As TRebarContourLink, mRebarContourCount As Long ' Явные связи Main/R2/R3 с исходными контурами и число этих связей.
Private mRebarContourCapacity As Long ' Выделенная ёмкость связей, увеличивается с удвоением без копирования на каждую запись.
Private mRebarContourMains As Object ' Точные ID основных рядов с уже объявленной связью; повтор проверяется без прохода по массиву.

' Последняя машинная диагностика; Nothing при отсутствии записанного события.
Public Property Get LastDiagnostic() As CGeometryDiagnostic
    Set LastDiagnostic = mDiagnostic
End Property

' Задаёт уникальное имя нового определения сечения; повторная инициализация готового определения
' запрещена.
Friend Sub Initialize(ByVal typeID As String, ByVal version As Long)
    If Len(mTypeID) > 0 Then DefinitionError "Определение уже создано."
    If Len(Trim$(typeID)) = 0 Or version < 1 Then DefinitionError "Не заданы тип и версия определения."
    mTypeID = typeID: mVersion = version
    Set mPaths = CreateObject("Scripting.Dictionary"): mPaths.CompareMode = vbBinaryCompare
    Set mGroups = CreateObject("Scripting.Dictionary"): mGroups.CompareMode = vbBinaryCompare
    Set mAnchors = CreateObject("Scripting.Dictionary"): mAnchors.CompareMode = vbBinaryCompare
    Set mSources = CreateObject("Scripting.Dictionary"): mSources.CompareMode = vbBinaryCompare
    Set mRebarContourMains = CreateObject("Scripting.Dictionary"): mRebarContourMains.CompareMode = vbBinaryCompare
End Sub

' Добавляет именованную точку привязки; повторный ID отклоняется.
Friend Sub AddAnchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
    RequireMutable
    If Len(Trim$(anchorID)) = 0 Or mAnchors.Exists(anchorID) Then DefinitionError "Пустой или повторный ID привязки '" & anchorID & "'."
    mAnchors.Add anchorID, Array(x, y)
End Sub

' Возвращает координаты именованной точки отдельными числами;
' изменение полученных переменных не меняет сохранённое описание.
Public Sub AnchorCoordinates(ByVal anchorID As String, ByRef x As Double, ByRef y As Double)
    If Not mAnchors.Exists(anchorID) Then DefinitionError "Не найдена привязка '" & anchorID & "'.", anchorID, "MISSING_ANCHOR"
    x = CDbl(mAnchors(anchorID)(0)): y = CDbl(mAnchors(anchorID)(1))
End Sub

' Связывает ID линии, правила или аннотации с адресом исходного параметра для сообщения об
' ошибке.
Friend Sub AddSource(ByVal entityID As String, ByVal reference As String)
    RequireMutable
    If Len(Trim$(entityID)) = 0 Or Len(Trim$(reference)) = 0 Then DefinitionError "ID и исходная ссылка обязательны."
    mSources(entityID) = reference
End Sub

' Возвращает адрес исходного параметра по ID; при отсутствии связи возвращает пустую строку.
Public Function SourceReference(ByVal entityID As String) As String
    If mSources.Exists(entityID) Then SourceReference = CStr(mSources(entityID))
End Function

' Возвращает имя типа сечения, указанное при создании описания.
Public Property Get TypeID() As String: TypeID = mTypeID: End Property
' Возвращает положительный номер версии формата описания сечения.
Public Property Get SchemaVersion() As Long: SchemaVersion = mVersion: End Property
' Показывает, завершено ли описание; после завершения его геометрия не изменяется.
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
' Возвращает количество линий, добавленных в описание сечения.
Public Property Get PathCount() As Long: PathCount = mPaths.Count: End Property
' Возвращает число замкнутых границ бетона: наружных контуров и отверстий.
Public Property Get LoopCount() As Long: LoopCount = mLoopCount: End Property
' Число объявленных правил, включая выключенные.
Public Property Get RuleCount() As Long: RuleCount = mRuleCount: End Property
' Возвращает число объявленных размеров, подписей и текстовых аннотаций.
Public Property Get AnnotationCount() As Long: AnnotationCount = mAnnotationCount: End Property

' Возвращает завершённую линию по её ID; изменять полученную геометрию нельзя.
Public Function Path(ByVal pathID As String) As CGeometryPath
    If Not mPaths.Exists(pathID) Then DefinitionError "Не найдена геометрическая линия '" & pathID & "'.", pathID, "MISSING_PATH"
    Set Path = mPaths(pathID)
End Function

' Возвращает имена отдельных линий группы в сохранённом порядке.
' Группа не создаёт соединяющую линию между раздельными участками.
Public Function TargetPaths(ByVal targetID As String) As Variant
    If mPaths.Exists(targetID) Then
        TargetPaths = Array(targetID)
    ElseIf mGroups.Exists(targetID) Then
        TargetPaths = mGroups(targetID)
    Else
        DefinitionError "Не найдена геометрическая линия или группа линий '" & targetID & "'."
    End If
End Function

' Возвращает независимую копию подготовленной области бетона
' без повторного чтения или выполнения рецепта.
Public Function ConcreteDomain() As CConcreteRegion
    RequireSealed
    Set ConcreteDomain = mRegion.Clone
End Function

' Добавляет аналитический путь с уникальным ID в определение сечения.
Friend Sub AddPath(ByVal path As CGeometryPath)
    RequireMutable
    If path Is Nothing Then DefinitionError "Не передана геометрическая линия размещения."
    If Not path.IsSealed Then DefinitionError "Геометрическая линия '" & path.ID & "' еще не подготовлен."
    If mPaths.Exists(path.ID) Or mGroups.Exists(path.ID) Then DefinitionError "Повторный ID геометрической линии '" & path.ID & "'."
    mPaths.Add path.ID, path
End Sub

' Добавляет именованную группу существующих путей, сохраняя их порядок.
Friend Sub AddGroup(ByVal groupID As String, ByVal pathIDs As Variant)
    RequireMutable
    If Len(Trim$(groupID)) = 0 Then DefinitionError "Не задан ID группы."
    If mPaths.Exists(groupID) Or mGroups.Exists(groupID) Then DefinitionError "Повторный ID группы '" & groupID & "'."
    If Not IsArray(pathIDs) Then DefinitionError "Для группы геометрических линий нужен массив их идентификаторов (ID)."
    Dim names() As String, i As Long, j As Long, count As Long
    count = UBound(pathIDs) - LBound(pathIDs) + 1
    If count < 1 Then DefinitionError "Группа не может быть пустой."
    ReDim names(0 To count - 1)
    For i = 0 To count - 1
        names(i) = CStr(pathIDs(LBound(pathIDs) + i))
        If Not mPaths.Exists(names(i)) Then DefinitionError "Не найден путь группы '" & names(i) & "'."
        For j = 0 To i - 1
            If names(i) = names(j) Then DefinitionError "Повторный путь в группе '" & groupID & "'."
        Next j
    Next i
    mGroups.Add groupID, names
End Sub

' Добавляет замкнутую бетонную границу с ролью Outer или Opening.
Friend Sub AddLoop(ByRef value As TSectionLoop)
    RequireMutable
    If value.Role <> LoopOuter And value.Role <> LoopOpening Then DefinitionError "Неизвестная роль контура."
    If Len(Trim$(value.LoopID)) = 0 Then DefinitionError "Не задан ID контура."
    Dim i As Long, path As CGeometryPath
    Set path = Me.Path(value.PathID)
    If Not path.IsClosed Then DefinitionError "Контур '" & value.LoopID & "' использует открытый путь."
    For i = 1 To mLoopCount
        If mLoops(i).LoopID = value.LoopID Then DefinitionError "Повторный ID контура '" & value.LoopID & "'."
    Next i
    If value.Role = LoopOuter And Len(value.ParentID) > 0 Then DefinitionError "Наружный контур не имеет родительского отверстия."
    If value.Role = LoopOpening And Len(value.ParentID) = 0 Then DefinitionError "Для отверстия не указан наружный контур."
    If value.NamingPolicy <> PreserveSegmentIDs And value.NamingPolicy <> SequentialByKind Then _
        DefinitionError "Неизвестная политика имён контурных сегментов."
    If value.NamingPolicy = PreserveSegmentIDs And Len(value.NamingPrefix) > 0 Then _
        DefinitionError "Префикс технических имён применяется только с SequentialByKind."
    If value.NamingPolicy = SequentialByKind Then
        If Not IsEmpty(value.ContourIDs) Then DefinitionError "SequentialByKind нельзя совмещать с явным массивом ID сегментов."
        If Len(Trim$(value.NamingPrefix)) = 0 Then value.NamingPrefix = "CONTOUR"
        For i = 1 To mLoopCount
            If mLoops(i).NamingPolicy = SequentialByKind And mLoops(i).NamingPrefix = value.NamingPrefix Then _
                DefinitionError "Для разных границ задайте разные префиксы технических имён: '" & value.NamingPrefix & "'."
        Next i
    End If
    If IsArray(value.ContourIDs) Then
        If UBound(value.ContourIDs) - LBound(value.ContourIDs) + 1 <> path.SegmentCount Then _
            DefinitionError "Число ID контурных сегментов не совпадает с путем."
        For i = LBound(value.ContourIDs) To UBound(value.ContourIDs)
            If VarType(value.ContourIDs(i)) <> vbString Then DefinitionError "ID контурного сегмента должен быть текстом."
        Next i
    ElseIf Not IsEmpty(value.ContourIDs) Then
        DefinitionError "ID контурных сегментов должны быть массивом строк."
    End If
    mLoopCount = mLoopCount + 1: ReDim Preserve mLoops(1 To mLoopCount)
    mLoops(mLoopCount) = value
End Sub

' Сохраняет правило размещения арматуры с независимой копией дополнительных параметров.
Friend Sub AddRule(ByRef value As TRebarRule)
    RequireMutable
    If value.DistributionScope <> PerPath Then DefinitionError "Правило '" & value.RuleID & "': поддерживается DistributionScope=PerPath."
    If Len(Trim$(value.RuleID)) = 0 Then DefinitionError "Не задан ID правила арматуры."
    Dim i As Long, parsedStation As Double, stations As Variant
    For i = 1 To mRuleCount
        If mRules(i).RuleID = value.RuleID Then DefinitionError "Повторный ID правила '" & value.RuleID & "'."
    Next i
    If IsArray(value.ExplicitStations) Then
        stations = value.ExplicitStations
        For i = LBound(value.ExplicitStations) To UBound(value.ExplicitStations)
            If IsObject(value.ExplicitStations(i)) Or IsArray(value.ExplicitStations(i)) Then _
                DefinitionError "Явные позиции должны содержать только числа."
            If Not TryReadStrictDouble(value.ExplicitStations(i), parsedStation) Then _
                DefinitionError "Расстояние от начала линии до места установки стержня должно быть числом без разделителей разрядов."
            stations(i) = parsedStation ' Каноническая копия не оставляет локализованный текст для повторного CDbl в исполнителе и не меняет массив вызывающего кода.
        Next i
    ElseIf Not IsEmpty(value.ExplicitStations) Then
        DefinitionError "Положения стержней должны задаваться массивом числовых расстояний от начала линии."
    End If
    mRuleCount = mRuleCount + 1: ReDim Preserve mRules(1 To mRuleCount)
    mRules(mRuleCount) = value
    If IsArray(stations) Then mRules(mRuleCount).ExplicitStations = stations
End Sub

' Добавляет декларативное описание размера или подписи без дублирования бетонного контура.
Friend Sub AddAnnotation(ByRef value As TSectionAnnotation)
    RequireMutable
    If Len(Trim$(value.AnnotationID)) = 0 Then DefinitionError "Не задан ID аннотации."
    Dim i As Long
    For i = 1 To mAnnotationCount
        If mAnnotations(i).AnnotationID = value.AnnotationID Then DefinitionError "Повторный ID аннотации '" & value.AnnotationID & "'."
    Next i
    mAnnotationCount = mAnnotationCount + 1: ReDim Preserve mAnnotations(1 To mAnnotationCount)
    mAnnotations(mAnnotationCount) = value
End Sub

' Сохраняет явную связь семейства арматуры с исходной фигурой и бетонной границей.
Friend Sub AddRebarContour(ByRef value As TRebarContourLink)
    RequireMutable
    If Len(value.MainRuleID) > 0 Then
        If mRebarContourMains.Exists(value.MainRuleID) Then DefinitionError "Повторная связь основного ряда с контуром.", value.MainRuleID
        mRebarContourMains.Add value.MainRuleID, True
    End If
    If mRebarContourCount = mRebarContourCapacity Then
        mRebarContourCapacity = CLng(GeomMax(8#, 2# * mRebarContourCapacity))
        ReDim Preserve mRebarContours(1 To mRebarContourCapacity)
    End If
    mRebarContourCount = mRebarContourCount + 1
    mRebarContours(mRebarContourCount) = value
End Sub

' Число явно объявленных связей семейств арматуры с контурами.
Friend Property Get RebarContourCount() As Long: RebarContourCount = mRebarContourCount: End Property

' Копирует связь с контуром; готовое определение при чтении не меняется.
Friend Sub CopyRebarContour(ByVal index As Long, ByRef value As TRebarContourLink)
    If index < 1 Or index > mRebarContourCount Then DefinitionError "Номер связи арматуры с контуром вне определения."
    value = mRebarContours(index)
End Sub

' Копирует роль и принадлежность бетонной границы для общей проверки рядов.
Friend Sub CopyLoop(ByVal index As Long, ByRef value As TSectionLoop)
    If index < 1 Or index > mLoopCount Then DefinitionError "Номер бетонного контура вне определения."
    value = mLoops(index)
End Sub

' Возвращает отдельную копию правила; изменение этой копии не меняет готовое описание сечения.
Friend Sub CopyRule(ByVal index As Long, ByRef value As TRebarRule)
    If index < 1 Or index > mRuleCount Then DefinitionError "Номер правила вне определения."
    value = mRules(index)
End Sub

' Возвращает отдельную копию описания размера или подписи для общего построителя аннотаций.
Friend Sub CopyAnnotation(ByVal index As Long, ByRef value As TSectionAnnotation)
    If index < 1 Or index > mAnnotationCount Then DefinitionError "Номер аннотации вне определения."
    value = mAnnotations(index)
End Sub

' Находит правило по его уникальному ID. Ссылка на неизвестное исходное правило вызывает ошибку.
Friend Function RuleIndex(ByVal ruleID As String) As Long
    Dim i As Long
    For i = 1 To mRuleCount
        If mRules(i).RuleID = ruleID Then RuleIndex = i: Exit Function
    Next i
    DefinitionError "Не найден исходный ряд '" & ruleID & "'."
End Function

' Проверяет наружные контуры, отверстия, их вложенность и независимые бетонные области.
' После успешной проверки запрещает дальнейшее изменение описания.
Friend Sub Seal()
    RequireMutable
    If mLoopCount = 0 Then DefinitionError "Не задан наружный контур бетона."
    Dim query As CSectionGeometryQuery, outers As Collection, parent As CConcreteRegion
    Dim outerByID As Object, region As CConcreteRegion, one As CConcreteRegion, i As Long
    Set query = New CSectionGeometryQuery: Set outers = New Collection
    Set outerByID = CreateObject("Scripting.Dictionary"): outerByID.CompareMode = vbBinaryCompare
    For i = 1 To mLoopCount
        If mLoops(i).Role = LoopOuter Then
            Set one = query.RegionFromContours(LoopContours(i, True))
            ValidateOneLoop query, one
            outers.Add one: outerByID.Add mLoops(i).LoopID, one
        End If
    Next i
    Set region = query.MergeDisjointRegions(outers)
    For i = 1 To mLoopCount
        If mLoops(i).Role = LoopOpening Then
            If Not outerByID.Exists(mLoops(i).ParentID) Then DefinitionError "Не найден Outer отверстия '" & mLoops(i).LoopID & "'."
            Set one = query.RegionFromContours(LoopContours(i, True))
            ValidateOneLoop query, one
            Set parent = outerByID(mLoops(i).ParentID) ' Отверстие проверяется внутри своего Outer, а не произвольно выбранной внешней границы.
            query.ValidateContainedRegion parent, one
            Set region = query.SubtractContainedRegion(region, one)
        End If
    Next i
    Dim visiting() As Long
    If mRuleCount > 0 Then
        ReDim visiting(1 To mRuleCount)
        For i = 1 To mRuleCount: ValidateRuleGraph i, visiting: Next i
    End If
    For i = 1 To mAnnotationCount: ValidateAnnotation mAnnotations(i): Next i
    Set mRegion = query.RegionFromContours(MaterialContours()) ' Подготавливаем и упорядочиваем точные границы общими геометрическими операциями. Характеристики бетонной сетки при этом не подменяются характеристиками точного контура. Готовая область сохраняется и при Build повторно не строится.
    mSealed = True
End Sub

' Проверяет привязки и размеры до построения сетки и расчёта НДС.
' Размер по исчезнувшей дуге или ссылка на несуществующий арматурный ряд вызывает ошибку.
Private Sub ValidateAnnotation(ByRef value As TSectionAnnotation)
    Dim ids As Variant, path As CGeometryPath, edge As TRegionEdge, unused As Long
    Select Case value.Kind
        Case "Anchors.Horizontal", "Anchors.Vertical", "Anchors.Aligned"
            RequireAnnotationAnchor value.AnnotationID, value.StartAnchorID
            RequireAnnotationAnchor value.AnnotationID, value.EndAnchorID
        Case "Text", "PathLength"
            RequireAnnotationAnchor value.AnnotationID, value.StartAnchorID
            If value.Kind = "PathLength" Then ids = TargetPaths(value.Target)
        Case "Radius", "Diameter"
            Set path = Me.Path(value.Target): path.CopySegment path.SegmentIndex(value.SegmentID), edge
            If edge.Sweep = 0# Then DefinitionError "Измерение радиуса/диаметра требует аналитическую дугу.", value.AnnotationID, "MEASUREMENT_TARGET"
        Case "VerticalExtent", "HorizontalExtent", "RebarLabel"
            ids = TargetPaths(value.Target)
        Case "RebarRowLabel", "RebarRowExtentLabel", "RebarRowExtremeLabel"
            unused = RuleIndex(value.Target)
        Case "BetweenPoints"
        Case Else
            DefinitionError "Не распознан способ построения размера '" & value.Kind & "'.", value.AnnotationID, "MEASUREMENT_KIND"
    End Select
End Sub

' Проверяет наличие точки привязки, необходимой для аннотации.
Private Sub RequireAnnotationAnchor(ByVal annotationID As String, ByVal anchorID As String)
    If Not mAnchors.Exists(anchorID) Then DefinitionError "Не найдена привязка '" & anchorID & "'.", annotationID, "MISSING_ANCHOR", anchorID
End Sub

' Проверяет замкнутость и геометрическую пригодность одной бетонной границы.
Private Sub ValidateOneLoop(ByVal query As CSectionGeometryQuery, ByVal region As CConcreteRegion)
    query.ValidatePreparedBoundary region
    If region.LoopCount <> 1 Then DefinitionError "Один Loop должен содержать одну замкнутую границу."
End Sub

' Проверяет зависимости арматурных правил и отклоняет циклы до исполнения.
Private Sub ValidateRuleGraph(ByVal index As Long, ByRef visiting() As Long)
    If visiting(index) = 2 Then Exit Sub
    If visiting(index) = 1 Then DefinitionError "Циклическая зависимость ряда '" & mRules(index).RuleID & "'."
    visiting(index) = 1 ' Метка текущей цепочки; повторный вход до значения 2 выявляет циклическую зависимость.
    Dim ids As Variant, source As Long
    If mRules(index).Enabled Then
        If mRules(index).Mode <> DependentOffset Then ids = TargetPaths(mRules(index).Target)
        If mRules(index).Mode = DependentOffset Or mRules(index).Mode = ProjectFrom Then
            source = RuleIndex(mRules(index).SourceRuleID)
            ValidateRuleGraph source, visiting
        End If
    End If
    visiting(index) = 2
End Sub

' Возвращает отдельный контейнер точных наружных границ и отверстий бетона.
Public Function MaterialContours() As CSectionContours
    Dim result As CSectionContours, i As Long
    Set result = New CSectionContours
    For i = 1 To mLoopCount: AppendLoop result, i, False: Next i
    Set MaterialContours = result
End Function

' Копирует точные границы бетона в переданный контейнер, предварительно очищая его.
Public Sub WriteContours(ByVal result As CSectionContours)
    RequireSealed
    If result Is Nothing Then DefinitionError "Контейнер контуров не передан."
    result.Clear
    Dim i As Long
    For i = 1 To mLoopCount: AppendLoop result, i, False: Next i
End Sub

' Преобразует аналитические сегменты одной границы в общий контейнер контуров.
Private Function LoopContours(ByVal index As Long, ByVal forceOuter As Boolean) As CSectionContours
    Dim result As CSectionContours
    Set result = New CSectionContours
    AppendLoop result, index, forceOuter
    Set LoopContours = result
End Function

' Передаёт сегменты границы с её ролью и источником в выходной контейнер.
Private Sub AppendLoop(ByVal result As CSectionContours, ByVal index As Long, ByVal forceOuter As Boolean)
    Dim path As CGeometryPath, role As String, edge As TRegionEdge, radius As Double, i As Long, contourID As String, comment As String
    Set path = Me.Path(mLoops(index).PathID)
    role = "Outer"
    If mLoops(index).Role = LoopOpening And Not forceOuter Then role = "Opening"
    For i = 1 To path.SegmentCount
        path.CopySegment i, edge
        comment = mLoops(index).Comment
        If edge.Sweep <> 0# And Len(mLoops(index).ArcComment) > 0 Then comment = mLoops(index).ArcComment
        contourID = path.SegmentID(i)
        If IsArray(mLoops(index).ContourIDs) Then contourID = CStr(mLoops(index).ContourIDs(LBound(mLoops(index).ContourIDs) + i - 1))
        If mLoops(index).NamingPolicy = SequentialByKind Then
            If edge.Sweep = 0# Then contourID = mLoops(index).NamingPrefix & "_LINE_" & CStr(i) Else contourID = mLoops(index).NamingPrefix & "_ARC_" & CStr(i)
        End If
        If edge.Sweep = 0# Then
            result.AddContourLine contourID, edge.X1, edge.Y1, edge.X2, edge.Y2, comment, mLoops(index).LoopID, role
        Else
            radius = Sqr((edge.X1 - edge.CenterX) ^ 2 + (edge.Y1 - edge.CenterY) ^ 2)
            If Abs(edge.Sweep) = 2# * GEOM_PI Then
                result.AddContourCircle contourID, edge.CenterX, edge.CenterY, radius, comment, mLoops(index).LoopID, role
            Else
                result.AddContourArc contourID, edge.X1, edge.Y1, edge.X2, edge.Y2, edge.CenterX, edge.CenterY, radius, _
                    edge.Sweep, comment, mLoops(index).LoopID, role
            End If
        End If
    Next i
End Sub

' Запрещает изменение описания после его завершения и проверки методом Seal.
Private Sub RequireMutable()
    If Len(mTypeID) = 0 Or mSealed Then DefinitionError "Определение не открыто для заполнения."
End Sub

' Требует завершённую проверенную геометрию перед её использованием.
Private Sub RequireSealed()
    If Not mSealed Then DefinitionError "Определение еще не опубликовано."
End Sub

' Выдаёт понятную ошибку определения сечения с сохранённой связью с вводом.
Private Sub DefinitionError(ByVal reason As String, Optional ByVal entityID As String = "", _
        Optional ByVal code As String = "DEFINITION_INVALID", Optional ByVal relatedEntityID As String = "")
    Dim reference As String
    If Len(entityID) = 0 Then entityID = mTypeID
    If Not mSources Is Nothing Then
        If mSources.Exists(entityID) Then reference = CStr(mSources(entityID))
    End If
    Set mDiagnostic = New CGeometryDiagnostic
    mDiagnostic.Initialize code, entityID, relatedEntityID, reason, 0#, 0#, reference
    SectionRecipeError "CSectionDefinition", "Определение '" & mTypeID & "': " & reason, reference
End Sub
