VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionDefinitionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Команды декларативного описания сечения
' ==========================================================================
' Позволяет рецепту задать границы бетона и отверстий, размещение арматуры,
' размеры и подписи. Геометрические вычисления выполняют общие классы.
' Initialize начинает описание, Finish проверяет и завершает его.
' Команда создания линии сразу сохраняет её в описании; возвращённый объект
' уже зарегистрирован. Другие команды связываются с этой линией по её ID.

Private mDefinition As CSectionDefinition ' Описание сечения, которое заполняется командами этого объекта; после Finish изменять его нельзя.
Private mQuery As CSectionGeometryQuery ' Общие геометрические операции над линиями и областями для выполнения команд рецепта.
Private mFinished As Boolean ' Признак завершения описания: после Finish повторное завершение и новые команды запрещены.

' Начинает описание сечения с именем TypeID и версией схемы; повторный вызов запрещён.
Public Sub Initialize(ByVal typeID As String, Optional ByVal schemaVersion As Long = 1)
    If Not mDefinition Is Nothing Then SectionRecipeError "CSectionDefinitionBuilder", "Фасад уже создан."
    Set mDefinition = New CSectionDefinition: mDefinition.Initialize typeID, schemaVersion
    Set mQuery = New CSectionGeometryQuery
End Sub

' Создаёт и сохраняет окружность: обход против часовой стрелки от направления +X.
' Окружность содержит один сегмент с тем же ID, что и вся линия.
Public Function CirclePath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddArc pathID, centerX, centerY, radius, 0#, 2# * GEOM_PI
    result.Seal True: mDefinition.AddPath result
    Set CirclePath = result
End Function

' Создаёт и сохраняет отрезок от первой заданной точки ко второй.
Public Function LinePath(ByVal pathID As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddLine pathID, x1, y1, x2, y2
    result.Seal: mDefinition.AddPath result
    Set LinePath = result
End Function

' Создаёт и сохраняет дугу по центру, радиусу, начальному углу и углу обхода со знаком.
' Полный оборот образует замкнутую окружность.
Public Function ArcPath(ByVal pathID As String, ByVal centerX As Double, ByVal centerY As Double, _
        ByVal radius As Double, ByVal startAngle As Double, ByVal sweep As Double) As CGeometryPath
    RequireOpen
    Dim result As CGeometryPath
    Set result = New CGeometryPath: result.Initialize pathID
    result.AddArc pathID, centerX, centerY, radius, startAngle, sweep
    result.Seal (Abs(sweep) = 2# * GEOM_PI): mDefinition.AddPath result
    Set ArcPath = result
End Function

' Создаёт и сохраняет скруглённый контур по вершинам и радиусам CPolygonSpec.
' Имена прямых участков и скруглений задаются в AddVertex; пустые имена дают Edge.ID и Corner.ID.
' Дуга существует только при положительном радиусе; исчезнувший участок сохраняет имя привязки.
Public Function RoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec) As CGeometryPath
    RequireOpen
    Set RoundedPolygon = mQuery.RoundedPolygon(pathID, polygon)
    mDefinition.AddPath RoundedPolygon
End Function

' Строит смещённый скруглённый контур по тому же описанию CPolygonSpec.
' Положительное расстояние направлено внутрь многоугольника, отрицательное — наружу.
' Для отступа от отверстия в сторону бетона задаётся отрицательное расстояние.
Public Function OffsetRoundedPolygon(ByVal pathID As String, ByVal polygon As CPolygonSpec, ByVal distance As Double) As CGeometryPath
    RequireOpen
    Set OffsetRoundedPolygon = mQuery.OffsetRoundedPolygon(pathID, polygon, distance)
    mDefinition.AddPath OffsetRoundedPolygon
End Function

' Добавляет в описание готовую непрерывную линию CGeometryPath, созданную отдельно.
' Чтобы она стала границей бетона, требуется Outer или Opening. Линии, полученные
' командами CirclePath, LinePath, ArcPath, RectanglePath и другими командами создания,
' уже добавлены в описание: повторный AddPath для них не требуется.
Public Sub AddPath(ByVal path As CGeometryPath)
    RequireOpen
    mDefinition.AddPath path
End Sub

' Сохраняет именованную точку привязки в координатах сечения, мм.
Public Sub Anchor(ByVal anchorID As String, ByVal x As Double, ByVal y As Double)
    RequireOpen
    mDefinition.AddAnchor anchorID, x, y
End Sub

' Создаёт именованную точку на выбранном сегменте: параметр от 0 до 1 идёт от начала к концу.
Public Sub PathAnchor(ByVal anchorID As String, ByVal pathID As String, ByVal segmentID As String, Optional ByVal parameter As Double = 0#)
    RequireOpen
    Dim path As CGeometryPath, x As Double, y As Double, tx As Double, ty As Double
    Set path = mDefinition.Path(pathID)
    path.AtDistance path.AnchorStation(segmentID, parameter), x, y, tx, ty
    mDefinition.AddAnchor anchorID, x, y
End Sub

' Связывает ID геометрии или правила с адресом исходного параметра без ссылки на Excel.
Public Sub Source(ByVal entityID As String, ByVal sourceReference As String)
    RequireOpen
    mDefinition.AddSource entityID, sourceReference
End Sub

' Создаёт новую точку привязки из существующей: отражает относительно локальной оси X,
' поворачивает вокруг (0, 0), затем переносит на заданное смещение.
Public Sub TransformAnchor(ByVal anchorID As String, ByVal sourceAnchorID As String, ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False)
    RequireOpen
    Dim x As Double, y As Double
    mDefinition.AnchorCoordinates sourceAnchorID, x, y
    mQuery.TransformCoordinates x, y, shiftX, shiftY, rotation, reflectX
    mDefinition.AddAnchor anchorID, x, y
End Sub

' Создаёт смещённую линию: положительное signedLeftDistance направлено влево по обходу.
' Прямые соединяются пересечением продолжений; стыки с дугами должны быть гладкими.
Public Function Offset(ByVal pathID As String, ByVal sourcePathID As String, ByVal signedLeftDistance As Double) As CGeometryPath
    RequireOpen
    On Error GoTo Failed
    Set Offset = mQuery.OffsetPath(pathID, mDefinition.Path(sourcePathID), signedLeftDistance)
    mDefinition.AddPath Offset
    Exit Function
Failed:
    SectionRecipeGeometryError "CSectionDefinitionBuilder", mDefinition.SourceReference(pathID)
End Function

' Обрезает или продлевает одну прямую до двух поперечных линий nX*x + nY*y = bound.
' Направление обхода исходного отрезка сохраняется.
Public Function ConstrainLine(ByVal pathID As String, ByVal sourcePathID As String, _
        ByVal startNX As Double, ByVal startNY As Double, ByVal startBound As Double, _
        ByVal endNX As Double, ByVal endNY As Double, ByVal endBound As Double) As CGeometryPath
    RequireOpen
    On Error GoTo Failed
    Set ConstrainLine = mQuery.ConstrainLine(pathID, mDefinition.Path(sourcePathID), startNX, startNY, startBound, endNX, endNY, endBound)
    mDefinition.AddPath ConstrainLine
    Exit Function
Failed:
    SectionRecipeGeometryError "CSectionDefinitionBuilder", mDefinition.SourceReference(pathID)
End Function

' Возвращает последнюю причину геометрической ошибки с адресом исходного параметра.
' Если ошибка не записана, возвращает Nothing.
Public Property Get LastGeometryDiagnostic() As CGeometryDiagnostic
    Dim raw As CGeometryDiagnostic, mapped As CGeometryDiagnostic
    If mQuery Is Nothing Then Exit Property
    Set raw = mQuery.LastDiagnostic
    If raw Is Nothing Then Exit Property
    Set mapped = New CGeometryDiagnostic
    mapped.Initialize raw.Code, raw.EntityID, raw.RelatedEntityID, raw.Reason, raw.Detail1, raw.Detail2, mDefinition.SourceReference(raw.EntityID)
    Set LastGeometryDiagnostic = mapped
End Property

' Выбирает непрерывный участок линии между двумя привязками к её сегментам.
' Направление обхода задаётся явно; fullLoopIfSame разрешает полный оборот
' замкнутого контура, если обе привязки совпали.
Public Function Chain(ByVal pathID As String, ByVal sourcePathID As String, _
        ByVal startSegmentID As String, ByVal startParameter As Double, ByVal endSegmentID As String, ByVal endParameter As Double, _
        Optional ByVal forward As Boolean = True, Optional ByVal fullLoopIfSame As Boolean = False) As CGeometryPath
    RequireOpen
    Dim source As CGeometryPath
    Set source = mDefinition.Path(sourcePathID)
    Set Chain = source.BetweenAnchors(pathID, startSegmentID, startParameter, endSegmentID, endParameter, forward, fullLoopIfSame)
    mDefinition.AddPath Chain
End Function

' Создаёт новую линию отражением, поворотом и переносом исходной.
' Сохраняет ID сегментов и сведения об источнике; отражение меняет знак угла дуг.
Public Function Transform(ByVal pathID As String, ByVal sourcePathID As String, ByVal shiftX As Double, ByVal shiftY As Double, _
        Optional ByVal rotation As Double = 0#, Optional ByVal reflectX As Boolean = False) As CGeometryPath
    RequireOpen
    Dim source As CGeometryPath, result As CGeometryPath
    Set source = mDefinition.Path(sourcePathID): Set result = source.Transformed(shiftX, shiftY, rotation, reflectX)
    Set Transform = result.WithID(pathID): mDefinition.AddPath Transform
End Function

' Создаёт и сохраняет замкнутый прямоугольник от точки (x, y).
' Ширина width и высота height задаются положительными, мм.
' Обход: низ слева направо, правая сторона снизу вверх, верх справа налево, левая сверху вниз.
' Имена относятся к исходным сторонам и сохраняются после поворота или отражения.
Public Function RectanglePath(ByVal pathID As String, ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, _
        Optional ByVal bottomEdgeID As String = "", Optional ByVal rightEdgeID As String = "", _
        Optional ByVal topEdgeID As String = "", Optional ByVal leftEdgeID As String = "") As CGeometryPath
    RequireOpen
    Dim original As CGeometryPath, result As CGeometryPath, edge As TRegionEdge, i As Long, name As String, assigned As String
    Set original = mQuery.AdjacentRectangles(pathID, Array(Array(x, y, width, height)))
    Set result = original
    If Len(Trim$(bottomEdgeID & rightEdgeID & topEdgeID & leftEdgeID)) > 0 Then
        Set result = New CGeometryPath: result.Initialize pathID
        For i = 1 To original.SegmentCount
            original.CopySegment i, edge
            ' Сопоставляем имя с координатами стороны, а не с номером после нормализации.
            assigned = ""
            If edge.Y1 = y And edge.Y2 = y Then assigned = bottomEdgeID
            If edge.X1 = x + width And edge.X2 = x + width Then assigned = rightEdgeID
            If edge.Y1 = y + height And edge.Y2 = y + height Then assigned = topEdgeID
            If edge.X1 = x And edge.X2 = x Then assigned = leftEdgeID
            name = original.SegmentID(i)
            If Len(Trim$(assigned)) > 0 Then name = assigned
            result.Append name, edge
        Next i
        result.Seal True
    End If
    Set RectanglePath = result
    mDefinition.AddPath RectanglePath
End Function

' Строит общую наружную границу смежных прямоугольников.
' Прямоугольники должны соприкасаться без наложения площадей;
' объединение произвольных пересекающихся фигур эта команда не выполняет.
Public Function AdjacentRectangles(ByVal pathID As String, ByVal rectangles As Variant) As CGeometryPath
    RequireOpen
    Set AdjacentRectangles = mQuery.AdjacentRectangles(pathID, rectangles)
    mDefinition.AddPath AdjacentRectangles
End Function

' Объединяет ID независимых линий в именованную группу, сохраняя их порядок.
' Соединяющие отрезки между линиями не добавляются.
Public Sub PathGroup(ByVal groupID As String, ByVal pathIDs As Variant)
    RequireOpen
    mDefinition.AddGroup groupID, pathIDs
End Sub

' Назначает замкнутую линию наружной границей бетона. Арматуру можно задать отдельно позднее.
' loopID именует границу, pathID — её линию; политика публикации не меняет локальные SegmentID.
Public Sub Outer(ByVal loopID As String, ByVal pathID As String, Optional ByVal comment As String = "", _
        Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "", _
        Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs, Optional ByVal namingPrefix As String = "")
    RequireOpen
    Dim value As TSectionLoop
    value.LoopID = loopID: value.PathID = pathID: value.Role = LoopOuter: value.Comment = comment: value.ArcComment = arcComment
    value.NamingPolicy = namingPolicy: value.NamingPrefix = namingPrefix
    If Not IsMissing(contourIDs) Then value.ContourIDs = contourIDs
    mDefinition.AddLoop value
End Sub

' Назначает замкнутую линию отверстием внутри наружного контура parentOuterID.
' Правильность расположения отверстия проверяется при Finish.
' parentOuterID относится к LoopID наружной границы, а не к имени её линии или сегмента.
Public Sub Opening(ByVal loopID As String, ByVal pathID As String, ByVal parentOuterID As String, _
        Optional ByVal comment As String = "", Optional ByVal contourIDs As Variant, Optional ByVal arcComment As String = "", _
        Optional ByVal namingPolicy As EContourNamingPolicy = PreserveSegmentIDs, Optional ByVal namingPrefix As String = "")
    RequireOpen
    Dim value As TSectionLoop
    value.LoopID = loopID: value.PathID = pathID: value.Role = LoopOpening
    value.ParentID = parentOuterID: value.Comment = comment: value.ArcComment = arcComment
    value.NamingPolicy = namingPolicy: value.NamingPrefix = namingPrefix
    If Not IsMissing(contourIDs) Then value.ContourIDs = contourIDs
    mDefinition.AddLoop value
End Sub

' Задаёт n мест установки стержней на рабочей части линии с выбранным EndMode.
' Для группы линий n применяется к каждой линии; n = 0 или d = 0 отключает правило.
Public Sub RebarByCount(ByVal ruleID As String, ByVal target As String, ByVal count As Long, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByCount, diameter, steelClass, options
    If count < 0 Then SectionRecipeError "CSectionDefinitionBuilder", "Количество стержней не может быть отрицательным.", value.SourceReference
    value.Count = count: value.Enabled = (count > 0 And diameter > 0#)
    mDefinition.AddRule value
End Sub

' Задаёт размещение стержней с шагом не больше указанного.
' Общий исполнитель выбирает достаточное количество стержней и фактический шаг.
Public Sub RebarByMaxStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByMaxStep, diameter, steelClass, options
    value.Step = step: mDefinition.AddRule value
End Sub

' Задаёт постоянный шаг стержней и способ распределения остатка длины.
' Режимы IncludeBoth и ClosedUniform требуют совместимости длины с этим шагом.
Public Sub RebarByFixedStep(ByVal ruleID As String, ByVal target As String, ByVal step As Double, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ByFixedStep, diameter, steelClass, options
    value.Step = step: mDefinition.AddRule value
End Sub

' Задаёт стержни в перечисленных расстояниях от начала линии, мм.
' Порядок значений в массиве определяет исходный порядок позиций.
Public Sub RebarAtStations(ByVal ruleID As String, ByVal target As String, ByVal stations As Variant, ByVal diameter As Double, _
        Optional ByVal options As CRebarRuleOptions = Nothing, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ExplicitPositions, diameter, steelClass, options
    value.ExplicitStations = stations: mDefinition.AddRule value
End Sub

' Задаёт зависимый ряд смещениями TangentOffset и NormalOffset от исходных стержней.
' Зависимые стержни сохраняют связь с исходными, без новой равномерной расстановки.
Public Sub RebarDependent(ByVal ruleID As String, ByVal sourceRuleID As String, ByVal diameter As Double, _
        ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, "", DependentOffset, diameter, steelClass, options
    value.SourceRuleID = sourceRuleID: mDefinition.AddRule value
End Sub

' Задаёт проекцию исходных позиций на целевую линию. Концевые отступы ограничивают
' допустимый участок; поведение при отсутствии пересечения задаётся параметрами правила.
Public Sub RebarProjected(ByVal ruleID As String, ByVal target As String, ByVal sourceRuleID As String, ByVal diameter As Double, _
        ByVal options As CRebarRuleOptions, Optional ByVal steelClass As String = "Rebar")
    Dim value As TRebarRule
    PrepareRule value, ruleID, target, ProjectFrom, diameter, steelClass, options
    value.SourceRuleID = sourceRuleID: mDefinition.AddRule value
End Sub

' Заполняет общие поля правила и копирует дополнительные параметры.
' Нулевой диаметр отключает правило, отрицательный вызывает ошибку.
Private Sub PrepareRule(ByRef value As TRebarRule, ByVal ruleID As String, ByVal target As String, ByVal mode As ERebarDistribution, _
        ByVal diameter As Double, ByVal steelClass As String, ByVal options As CRebarRuleOptions)
    RequireOpen
    value.RuleID = ruleID: value.Target = target: value.Mode = mode: value.Diameter = diameter
    value.SteelClass = steelClass: value.Enabled = (diameter > 0#)
    If options Is Nothing Then Set options = New CRebarRuleOptions
    options.CopyTo value
    If diameter < 0# Then SectionRecipeError "CSectionDefinitionBuilder", "Диаметр не может быть отрицательным.", value.SourceReference
End Sub

' Задаёт полный вертикальный габарит выбранных линий по их точной геометрии.
' Сторона размещения размерной линии указывается отдельно.
Public Sub VerticalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, _
        ByVal placementX As Double, Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "VerticalExtent": value.Target = target: value.Name = name
    value.StartX = placementX: value.EndX = placementX: value.NormalX = normalX: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт полный горизонтальный габарит выбранных линий по их точной геометрии.
' Измеряется габарит, а не длина линии или ширина до скругления.
Public Sub HorizontalExtent(ByVal annotationID As String, ByVal target As String, ByVal name As String, _
        ByVal placementY As Double, Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "HorizontalExtent": value.Target = target: value.Name = name
    value.StartY = placementY: value.EndY = placementY: value.NormalY = normalY: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт размер от startX до endX на высоте placementY, мм; значение равно endX - startX.
' DisplayIncrement округляет только подпись. Для положительного размера концы задаются слева направо.
Public Sub HorizontalDistance(ByVal annotationID As String, ByVal name As String, _
        ByVal startX As Double, ByVal endX As Double, ByVal placementY As Double, _
        Optional ByVal normalY As Double = 1#, Optional ByVal comment As String = "", Optional ByVal displayIncrement As Double = 0#)
    RequireOpen
    If displayIncrement < 0# Then SectionRecipeError "CSectionDefinitionBuilder", "Шаг округления подписи не может быть отрицательным."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "BetweenPoints": value.Name = name
    value.StartX = startX: value.EndX = endX: value.StartY = placementY: value.EndY = placementY
    value.NormalY = normalY: value.Value = endX - startX: value.Comment = comment: value.DisplayIncrement = displayIncrement
    mDefinition.AddAnnotation value
End Sub

' Задаёт размер между startY и endY на вертикали placementX, мм.
' Значение равно модулю разности координат; округление подписи не меняет геометрию.
Public Sub VerticalDistance(ByVal annotationID As String, ByVal name As String, _
        ByVal startY As Double, ByVal endY As Double, ByVal placementX As Double, _
        Optional ByVal normalX As Double = -1#, Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "BetweenPoints": value.Name = name
    value.StartX = placementX: value.EndX = placementX: value.StartY = startY: value.EndY = endY
    value.NormalX = normalX: value.Value = Abs(endY - startY): value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт размер между двумя именованными точками: горизонтальный, вертикальный
' или вдоль соединяющего их отрезка, согласно выбранному виду.
Public Sub Dimension(ByVal annotationID As String, ByVal startAnchorID As String, ByVal endAnchorID As String, _
        ByVal name As String, Optional ByVal measurement As String = "Aligned", _
        Optional ByVal normalX As Double = 0#, Optional ByVal normalY As Double = 0#, Optional ByVal comment As String = "")
    RequireOpen
    If measurement <> "Horizontal" And measurement <> "Vertical" And measurement <> "Aligned" Then _
        SectionRecipeError "CSectionDefinitionBuilder", "Measurement должен быть Horizontal, Vertical или Aligned."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Anchors." & measurement: value.Name = name
    value.StartAnchorID = startAnchorID: value.EndAnchorID = endAnchorID
    value.NormalX = normalX: value.NormalY = normalY: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт размер по точной суммарной длине выбранных линий и его текстовую привязку.
' Разрывы между линиями группы не включаются в длину.
Public Sub PathLength(ByVal annotationID As String, ByVal target As String, ByVal placementAnchorID As String, _
        Optional ByVal name As String = "L", Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "PathLength": value.Target = target
    value.StartAnchorID = placementAnchorID: value.Name = name: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт радиус дуги или окружности; параметр сегмента выбирает конец радиусной линии.
' Для прямого отрезка вызов отклоняется.
Public Sub Radius(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "R", Optional ByVal comment As String = "")
    CircularDimension annotationID, pathID, segmentID, parameter, name, comment, False
End Sub

' Задаёт диаметр дуги или окружности; второй конец лежит напротив на той же окружности.
' Для прямого отрезка вызов отклоняется.
Public Sub Diameter(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        Optional ByVal parameter As Double = 0.5, Optional ByVal name As String = "D", Optional ByVal comment As String = "")
    CircularDimension annotationID, pathID, segmentID, parameter, name, comment, True
End Sub

' Добавляет радиус или диаметр аналитической дуги с привязкой по параметру сегмента.
Private Sub CircularDimension(ByVal annotationID As String, ByVal pathID As String, ByVal segmentID As String, _
        ByVal parameter As Double, ByVal name As String, ByVal comment As String, ByVal diameter As Boolean)
    RequireOpen
    If parameter < 0# Or parameter > 1# Then SectionRecipeError "CSectionDefinitionBuilder", "Параметр точки дуги должен лежать в [0;1]."
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Radius": If diameter Then value.Kind = "Diameter"
    value.Target = pathID: value.SegmentID = segmentID: value.Value = parameter
    value.Name = name: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Задаёт текст с привязкой к именованной точке модели.
' Шрифт и положение на диаграмме определяет общее оформление схемы.
Public Sub Text(ByVal annotationID As String, ByVal anchorID As String, ByVal textValue As String, _
        Optional ByVal comment As String = "")
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = annotationID: value.Kind = "Text": value.StartAnchorID = anchorID
    value.Name = textValue: value.Comment = comment
    mDefinition.AddAnnotation value
End Sub

' Привязывает общую подпись арматурной группы к выбранной линии.
' Направление выноски normalX/normalY может быть наклонным.
Public Sub RebarLabel(ByVal groupID As String, ByVal target As String, ByVal normalX As Double, ByVal normalY As Double)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarLabel": value.Target = target: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY
    mDefinition.AddAnnotation value
End Sub

' Привязывает подпись арматурной группы к фактическим концам ряда, заданного ruleID.
Public Sub RebarRowLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowLabel": value.Target = ruleID: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY
    mDefinition.AddAnnotation value
End Sub

' Размещает подпись по крайним координатам стержней основного ряда,
' а не по полной длине исходной линии расстановки.
Public Sub RebarRowExtentLabel(ByVal groupID As String, ByVal ruleID As String, ByVal normalX As Double, ByVal normalY As Double, _
        Optional ByVal reverseDirection As Boolean = False)
    RequireOpen
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowExtentLabel": value.Target = ruleID: value.Name = groupID
    value.NormalX = normalX: value.NormalY = normalY: value.ReverseDirection = reverseDirection
    mDefinition.AddAnnotation value
End Sub

' Проверяет контуры, ссылки, зависимости и аннотации и возвращает готовое описание.
' После этого изменение описания запрещено; фактическую расстановку арматуры
' проверяет CSectionBuilder до построения бетонной сетки.
Public Function Finish() As CSectionDefinition
    RequireOpen
    mDefinition.Seal
    mFinished = True
    Set Finish = mDefinition
End Function

' Отдельно задаёт, с какой стороны выбирать крайний ряд арматуры и куда выносить подпись.
' Например, ряд у отверстия выбирается со стороны бетона, а его подпись выводится в отверстие.
Public Sub RebarRowExtremeLabel(ByVal groupID As String, ByVal ruleID As String, ByVal directionX As Double, ByVal directionY As Double, _
        ByVal normalX As Double, ByVal normalY As Double)
    RequireOpen
    If Abs(directionX * directionX + directionY * directionY - 1#) > GEOM_TOLERANCE Then _
        SectionRecipeError "CSectionDefinitionBuilder", "Направление крайнего ряда должно быть единичным."
    Dim value As TSectionAnnotation
    value.AnnotationID = "REBAR_" & groupID: value.Kind = "RebarRowExtremeLabel": value.Target = ruleID: value.Name = groupID
    value.DirectionX = directionX: value.DirectionY = directionY
    value.NormalX = normalX: value.NormalY = normalY
    mDefinition.AddAnnotation value
End Sub

' Запрещает дальнейшее описание после Finish; готовое определение неизменяемо.
Private Sub RequireOpen()
    If mDefinition Is Nothing Or mFinished Then SectionRecipeError "CSectionDefinitionBuilder", "Фасад не открыт для описания сечения."
End Sub
