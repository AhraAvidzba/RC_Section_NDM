VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAutoCADSectionModelImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Импорт расчетной сетки из AutoCAD Region
' ==========================================================================
' Импортер читает активный чертеж AutoCAD и превращает Region на заданных
' слоях бетона и арматуры в CSectionModel. После этого расчетное ядро уже не
' знает, что геометрия пришла из AutoCAD: оно видит те же бетонные элементы
' и арматуру, что и при автоматической генерации.
'
' Класс не выполняет расчет и не пишет Results. Его ответственность - только
' аккуратно отфильтровать AutoCAD-сущности и заполнить модель сечения.

Private Const REBAR_STEEL_CLASS As String = "Rebar" ' Техническая метка расчетной арматуры в snapshot геометрии.
Private Const LOCAL_INERTIA_SQUARE_REL_TOLERANCE As Double = 0.000001 ' Допуск, при котором I1/I2 Region не задают однозначного угла.
Private Const EXPLICIT_RECTANGLE_AREA_REL_TOLERANCE As Double = 0.02 ' Допуск только для тестовых/подготовленных данных с уже известными Width/Height.

Private mEdgeProbeCount As Long ' ДЛЯ ТЕСТОВ: сколько раз импорт обращался к Explode за углом простой грани.
Private mImportComment As String ' Диагностика достоверных контуров, сформированная ответственным importer-ом.

' Возвращает пояснение последнего успешного импорта контуров для сообщения
' Excel. Writer и оболочка кнопки не угадывают, был ли слой пустым.
Public Property Get ImportComment() As String
    ImportComment = mImportComment
End Property

' Читает активный документ AutoCAD через COM и запускает общий импорт
' ModelSpace. Метод остается тонкой оболочкой: проверяет доступность AutoCAD,
' берет слои из Config и не выполняет расчетную интерпретацию сечения.
Public Function ImportFromActiveDocument(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки Config не переданы."

    Dim acad As Object
    Set acad = ConnectToRunningAutoCAD()

    Dim doc As Object
    On Error Resume Next
    Set doc = acad.ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then Err.Raise vbObjectError + 4402, "CAutoCADSectionModelImporter", _
        "AutoCAD открыт, но активный чертеж не найден. Откройте или создайте чертеж и повторите импорт."

    Set ImportFromActiveDocument = ImportConfiguredModelSpace(doc.ModelSpace, settings, units, doc.Layers)
End Function

' Подготавливает параметры импорта из прочитанного Config и передает ModelSpace
' общему обходу Region. Порог из пользовательской INPUT-площади переводится
' в мм2; координаты и геометрические свойства самих Region уже заданы в мм.
' Получение активного документа остается только в ImportFromActiveDocument.
Public Function ImportConfiguredModelSpace(ByVal modelSpace As Object, ByVal settings As CSystemSettingsReader, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal drawingLayers As Object = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки Config не переданы."
    Dim concreteLayer As String, rebarLayer As String
    ReadAutoCADCommonMaterialLayers settings, concreteLayer, rebarLayer
    Dim importUnits As CUnitSystem
    Set importUnits = units
    If importUnits Is Nothing Then
        Set importUnits = New CUnitSystem
        importUnits.LoadFromSettings settings
    End If

    Dim minArea As Double
    minArea = settings.GetRequiredDouble("AutoCAD.Import.MinArea")
    If minArea < 0# Then Err.Raise vbObjectError + 4404, "CAutoCADSectionModelImporter", _
        settings.InputErrorMessage("AutoCAD.Import.MinArea", "Порог площади не может быть отрицательным.", _
            "Введите ноль для отключения фильтра или положительную площадь в единицах INPUT Area.")
    On Error GoTo InvalidMinimumArea
    minArea = importUnits.InputAreaToInternal(minArea)
    On Error GoTo 0
    ' Эти два поля опциональны по контракту. Без точных контуров сохраняется
    ' расчетная сетка. Непустое имя проверяется, но отсутствие
    ' объектов на существующем слое не делает контур обязательным.
    Dim contourLayer As String, openingLayer As String
    contourLayer = OptionalContourLayer(settings, "AutoCAD.Common.SectionContourLayer")
    openingLayer = OptionalContourLayer(settings, "AutoCAD.Common.OpeningContourLayer")
    Set ImportConfiguredModelSpace = ImportFromModelSpace(modelSpace, _
        concreteLayer, rebarLayer, REBAR_STEEL_CLASS, minArea, importUnits, contourLayer, openingLayer, drawingLayers)
    Exit Function
InvalidMinimumArea:
    Err.Raise vbObjectError + 4404, "CAutoCADSectionModelImporter", _
        settings.InputErrorMessage("AutoCAD.Import.MinArea", "Площадь слишком велика для перевода в мм2.", _
            "Уменьшите число и проверьте выбранную единицу INPUT Area.")
End Function

' Проходит по ModelSpace и отбирает только Region на пользовательских слоях
' бетона и арматуры. На выходе возвращает полностью заполненный CSectionModel,
' пригодный для того же расчетного pipeline, что и встроенные генераторы.
' После чтения всей сетки назначает средний угол оболочкам без направления грани.
Public Function ImportFromModelSpace(ByVal modelSpace As Object, ByVal concreteLayer As String, _
        ByVal rebarLayer As String, ByVal steelClass As String, ByVal minArea As Double, Optional ByVal units As CUnitSystem = Nothing, _
        Optional ByVal contourLayer As String = "", Optional ByVal openingLayer As String = "", Optional ByVal drawingLayers As Object = Nothing) As CSectionModel
    mImportComment = vbNullString
    If modelSpace Is Nothing Then Err.Raise vbObjectError + 4403, "CAutoCADSectionModelImporter", "ModelSpace AutoCAD не передан."
    If Len(Trim$(concreteLayer)) = 0 Or Len(Trim$(rebarLayer)) = 0 Then Err.Raise vbObjectError + 4405, "CAutoCADSectionModelImporter", _
        "Не заданы слои импорта. Укажите отдельный слой бетонных Region и отдельный слой арматурных Region."
    If StrComp(Trim$(concreteLayer), Trim$(rebarLayer), vbTextCompare) = 0 Then Err.Raise vbObjectError + 4405, "CAutoCADSectionModelImporter", _
        "Слои бетона и арматуры совпадают. Укажите разные слои, чтобы однозначно определить тип каждого Region."

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    Dim entity As Object, entityLayer As String, materialEntities As Collection, materialLayers As Collection
    Dim outerEntities As Collection, openingEntities As Collection, i As Long
    Set materialEntities = New Collection: Set materialLayers = New Collection
    Set outerEntities = New Collection: Set openingEntities = New Collection
    mEdgeProbeCount = 0
    ' Читаем слой один раз и сохраняем только ссылки этой операции. Последующее
    ' Copy/Explode уже не меняет коллекцию во время ее COM-обхода.
    For Each entity In modelSpace
        entityLayer = CStr(entity.Layer)
        If SameLayer(entityLayer, concreteLayer) Or SameLayer(entityLayer, rebarLayer) Then
            If IsRegionEntity(entity) Then
                materialEntities.Add entity: materialLayers.Add entityLayer
            End If
        End If
        If Len(contourLayer) > 0 And SameLayer(entityLayer, contourLayer) Then outerEntities.Add entity
        If Len(openingLayer) > 0 And SameLayer(entityLayer, openingLayer) Then openingEntities.Add entity
    Next entity
    For i = 1 To materialEntities.Count
        Set entity = materialEntities(i): entityLayer = CStr(materialLayers(i))
        If SameLayer(entityLayer, concreteLayer) Then
            AddConcreteRegionEntity model, entity, minArea, units, entityLayer
        Else
            AddRebarRegionEntity model, entity, steelClass, minArea, units, entityLayer
        End If
    Next i

    ValidateImportedModel model
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    ImportMaterialContours model, modelSpace, concreteLayer, rebarLayer, contourLayer, openingLayer, drawingLayers, outerEntities, openingEntities
    Set ImportFromModelSpace = model
End Function

' Собирает опциональные достоверные контуры до возврата новой модели.
' Сначала проверяются найденные объекты и отверстия, затем готовая область
' сохраняется вместе с геометрией. Пустые слои допустимы: без наружного
' контура границы определяются по сетке. Ошибка не оставляет частичный импорт.
Private Sub ImportMaterialContours(ByVal model As CSectionModel, ByVal modelSpace As Object, _
        ByVal concreteLayer As String, ByVal rebarLayer As String, ByVal contourLayer As String, _
        ByVal openingLayer As String, ByVal drawingLayers As Object, _
        ByVal outerEntities As Collection, ByVal openingEntities As Collection)
    If Len(contourLayer) = 0 And Len(openingLayer) = 0 Then
        mImportComment = mImportComment & "Границы для расчета трещин определяются по бетонной сетке."
        Exit Sub
    End If
    ValidateContourLayerMapping concreteLayer, rebarLayer, contourLayer, openingLayer
    If drawingLayers Is Nothing Then
        On Error Resume Next
        Set drawingLayers = modelSpace.Document.Layers
        On Error GoTo 0
    End If
    If drawingLayers Is Nothing Then Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", _
        "Не удалось проверить слои достоверных контуров AutoCAD. Передайте коллекцию Layers активного чертежа."
    contourLayer = ExistingContourLayer(drawingLayers, contourLayer)
    openingLayer = ExistingContourLayer(drawingLayers, openingLayer)
    If Len(contourLayer) = 0 Then Set outerEntities = New Collection
    If Len(openingLayer) = 0 Then Set openingEntities = New Collection
    If Len(contourLayer) = 0 And Len(openingLayer) = 0 Then
        mImportComment = mImportComment & "Границы для расчета трещин определяются по бетонной сетке."
        Exit Sub
    End If

    Dim entity As Object
    Dim originalOuter As Collection, originalOpenings As Collection, exportedOuter As Boolean
    Set originalOuter = OriginalContours(outerEntities)
    Set originalOpenings = OriginalContours(openingEntities)
    ' Приоритет исходников действует отдельно на каждом слое. Наличие
    ' исходного outer не отменяет opening, сохранившиеся только в экспорте.
    If originalOuter.Count > 0 Then Set outerEntities = originalOuter
    If originalOpenings.Count > 0 Then Set openingEntities = originalOpenings
    exportedOuter = (originalOuter.Count = 0 And outerEntities.Count > 0)
    If outerEntities.Count = 0 And openingEntities.Count = 0 Then
        mImportComment = mImportComment & "Дополнительные контуры не заданы; границы для расчета трещин определяются по бетонной сетке."
        Exit Sub
    End If
    Dim query As CSectionGeometryQuery, reader As CAutoCADContourReader, regions As Collection
    Dim domain As CConcreteRegion, opening As CConcreteRegion, i As Long, j As Long, commentText As String, openings As Collection
    Dim validatedOpenings As CConcreteRegion
    Set query = New CSectionGeometryQuery: Set reader = New CAutoCADContourReader
    Set openings = New Collection
    If exportedOuter Then
        Set domain = ReadExportedMaterialContours(outerEntities, openingEntities, reader, query)
    ElseIf outerEntities.Count > 0 Then
        Set regions = New Collection
        For Each entity In outerEntities: regions.Add reader.ReadRegion(entity): Next entity
        Set domain = query.MergeDisjointRegions(regions)
    Else
        query.Initialize model
        Set domain = query.ConcreteDomain
    End If
    If Not exportedOuter Then
    For Each entity In openingEntities
        Set opening = reader.ReadRegion(entity)
        ' Region исходного outer может уже содержать это отверстие. Только
        ' собственная экспортная копия той же границы не вычитается повторно.
        If outerEntities.Count > 0 And IsNDMContourOutput(entity) And opening.LoopCount = 1 Then
            Dim matchingOpening As Boolean
            If FindMatchingBoundaryLoop(domain, opening, query, matchingOpening) Then
                If matchingOpening Then GoTo NextOpening
            End If
        End If
        If outerEntities.Count = 0 Then
            For j = 1 To opening.LoopCount
                If query.LoopIsOpening(opening, j) Then Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", _
                    "Внутри Region отверстия есть отдельная бетонная часть, но наружный контур не задан. На AutoCAD.Common.SectionContourLayer разместите наружный Region с вырезанным отверстием и отдельный Region этой бетонной части."
            Next j
        End If
        If outerEntities.Count = 0 Then
            ' Сеточная пустота может совпадать с точным opening, быть меньше
            ' либо больше него. Это не конфликт достоверных исходных границ:
            ' вычитаем opening из сетки, сохраняя его собственный контур.
            ' Между самими точными отверстиями строгая проверка остается.
            openings.Add opening
            Set validatedOpenings = query.MergeDisjointRegions(openings)
            Set domain = query.SubtractRegion(domain, opening)
        Else
            Set domain = query.SubtractContainedRegion(domain, opening)
            openings.Add opening
        End If
NextOpening:
    Next entity
    End If
    For i = 1 To model.RebarCount
        If Not query.ContainsPoint(domain, model.RebarX(i), model.RebarY(i), True) Then Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", _
            "Арматура " & model.RebarID(i) & " находится вне импортированного бетонного контура или в отверстии. Проверьте слои и координаты исходных объектов."
    Next i
    If outerEntities.Count > 0 Then
        commentText = "Достоверный контур AutoCAD: реальные линии/дуги и внутренние контуры отверстий; координаты в мм."
    Else
        commentText = "Наружный контур не задан; границы для расчета трещин определяются по бетонной сетке с учетом импортированных контуров отверстий. Координаты в мм."
    End If
    If Len(openingLayer) > 0 And openingEntities.Count = 0 Then commentText = commentText & _
        " На заданном слое отверстий '" & openingLayer & "' объектов нет; дополнительные контуры отверстий не заданы."
    If outerEntities.Count > 0 Then
        model.Contours.AddMaterialRegion domain, "MATERIAL", commentText
    Else
        ' Только заданные opening не превращают прежнюю оболочку сетки
        ' в достоверный outer. Это важно для выбора якорных групп СП 35.
        For i = 1 To openings.Count
            model.Contours.AddMaterialRegion openings(i), "OPENING_SOURCE_" & CStr(i), commentText, True
        Next i
    End If
    mImportComment = mImportComment & commentText
End Sub

' Пустое опциональное поле допустимо, но TODO и ошибка Excel не равны
' отсутствию запроса. Непустое имя проходит тот же валидатор, что слои сетки.
Private Function OptionalContourLayer(ByVal settings As CSystemSettingsReader, ByVal key As String) As String
    If Len(Trim$(settings.GetRawString(key))) = 0 Then Exit Function
    OptionalContourLayer = RequiredAutoCADLayerName(settings, key)
End Function

' Не позволяет использовать один слой сразу для расчетных Region и
' достоверных границ: это неоднозначный вход, а не повод менять роль объекта.
Private Sub ValidateContourLayerMapping(ByVal concreteLayer As String, ByVal rebarLayer As String, _
        ByVal contourLayer As String, ByVal openingLayer As String)
    Dim layers As Variant, i As Long, j As Long
    layers = Array(concreteLayer, rebarLayer, contourLayer, openingLayer)
    For i = 0 To UBound(layers)
        If Len(CStr(layers(i))) > 0 Then
            For j = 0 To i - 1
                If Len(CStr(layers(j))) > 0 And SameLayer(CStr(layers(i)), CStr(layers(j))) Then _
                    Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", "Слои бетона, арматуры, внешнего контура и отверстий должны иметь разные имена. Исправьте AutoCAD.Common/Import на Config."
            Next j
        End If
    Next i
End Sub

' Возвращает имя только существующего дополнительного слоя. Отсутствие
' слоя допустимо: его имя можно заранее сохранить для других чертежей.
Private Function ExistingContourLayer(ByVal layers As Object, ByVal layerName As String) As String
    If Len(layerName) = 0 Then Exit Function
    Dim layer As Object
    On Error Resume Next
    Set layer = layers.Item(layerName)
    On Error GoTo 0
    If layer Is Nothing Then
        mImportComment = mImportComment & "Дополнительный слой '" & layerName & "' отсутствует в чертеже; контуры с него не импортировались. "
    Else
        ExistingContourLayer = layerName
    End If
End Function

' На одном общем слое исходники имеют приоритет над экспортированными
' копиями. Решение не распространяется на другой слой/роль контура.
Private Function OriginalContours(ByVal entities As Collection) As Collection
    Dim originals As Collection, entity As Object
    Set originals = New Collection
    For Each entity In entities
        If Not IsNDMContourOutput(entity) Then originals.Add entity
    Next entity
    Set OriginalContours = originals
End Function

' Собирает экспортированные наружные границы и выбранные отверстия обратно
' в одну область, сохраняя бетон внутри проема. Отверстия могут быть исходными
' или экспортированными; роль каждого контура проверяется по его слою.
Private Function ReadExportedMaterialContours(ByVal outers As Collection, ByVal openings As Collection, _
        ByVal reader As CAutoCADContourReader, ByVal query As CSectionGeometryQuery) As CConcreteRegion
    Dim raw() As TRegionEdge, one() As TRegionEdge, count As Long, oneCount As Long
    Dim entity As Object, region As CConcreteRegion, group As Long, entities As Collection, i As Long
    Dim sourceRegions As Collection, sourceRoles As Collection, domain As CConcreteRegion, loopID As Long
    Dim boundary As CConcreteRegion, actualOpening As Boolean, expectedOpening As Boolean
    Set sourceRegions = New Collection: Set sourceRoles = New Collection
    For group = 0 To 1
        If group = 0 Then Set entities = outers Else Set entities = openings
        For Each entity In entities
            Set region = reader.ReadRegion(entity): region.CopyEdges one, oneCount
            sourceRegions.Add region: sourceRoles.Add (group = 1)
            ReDim Preserve raw(1 To count + oneCount)
            For i = 1 To oneCount: LSet raw(count + i) = one(i): Next i
            count = count + oneCount
        Next entity
    Next group
    Set domain = query.BoundaryRegion(raw, count, "ValidatedAutoCADExportContours")
    ' Вложенность не должна превращать ошибочно расположенное opening
    ' в новый бетонный участок или наружную границу в отверстие.
    For i = 1 To sourceRegions.Count
        Set region = sourceRegions(i)
        For loopID = 1 To region.LoopCount
            Set boundary = BoundaryLoopRegion(region, loopID, query)
            expectedOpening = region.LoopIsOpening(loopID) Xor CBool(sourceRoles(i))
            If Not FindMatchingBoundaryLoop(domain, boundary, query, actualOpening) Then _
                Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", "Не удалось сохранить границу импортированного контура. Проверьте наружные контуры и отверстия."
            If actualOpening <> expectedOpening Then Err.Raise vbObjectError + 4411, "CAutoCADSectionModelImporter", _
                "Контур находится не на своем слое: отверстия должны быть внутри бетона, а отдельные бетонные части внутри отверстий должны находиться на слое наружного контура."
        Next loopID
    Next i
    Set ReadExportedMaterialContours = domain
End Function

' Возвращает одну замкнутую границу как положительную область для точного
' сравнения линий/дуг. Исходная роль отверстия при этом остается у region.
Private Function BoundaryLoopRegion(ByVal region As CConcreteRegion, ByVal loopID As Long, _
        ByVal query As CSectionGeometryQuery) As CConcreteRegion
    Dim edges() As TRegionEdge, one() As TRegionEdge, count As Long, oneCount As Long, i As Long
    region.CopyEdges edges, count
    For i = 1 To count
        If edges(i).LoopID = loopID Then
            oneCount = oneCount + 1: ReDim Preserve one(1 To oneCount)
            LSet one(oneCount) = edges(i)
        End If
    Next i
    Set BoundaryLoopRegion = query.BoundaryRegion(one, oneCount)
End Function

' Ищет совпадающую аналитическую границу и возвращает ее фактическую роль.
' Совпадение площади само по себе недостаточно: сравнение выполняет query.
Private Function FindMatchingBoundaryLoop(ByVal domain As CConcreteRegion, ByVal boundary As CConcreteRegion, _
        ByVal query As CSectionGeometryQuery, ByRef isOpening As Boolean) As Boolean
    Dim loopID As Long, existing As CConcreteRegion
    isOpening = False
    For loopID = 1 To domain.LoopCount
        Set existing = BoundaryLoopRegion(domain, loopID, query)
        If query.RegionsCoincide(existing, boundary) Then
            isOpening = domain.LoopIsOpening(loopID)
            FindMatchingBoundaryLoop = True
            Exit Function
        End If
    Next loopID
End Function

' Читает один бетонный Region из AutoCAD: площадь, центр тяжести, реальные
' центральные инерции и, только для почти изотропного случая, подсказку угла
' по первой прямой грани. Width/Height здесь не восстанавливаются.
Private Sub AddConcreteRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, ByVal minArea As Double, _
        ByVal units As CUnitSystem, ByVal sourceLayer As String)
    Dim area As Double
    area = InputArea(CDbl(entity.Area), units)
    If area < minArea Then Exit Sub

    Dim centroid As Variant
    centroid = entity.Centroid

    Dim xCoord As Double
    Dim yCoord As Double
    xCoord = InputLength(CDbl(centroid(LBound(centroid))), units)
    yCoord = InputLength(CDbl(centroid(LBound(centroid) + 1)), units)

    Dim localIx As Double
    Dim localIy As Double
    Dim localIxy As Double
    ReadCentralRegionInertia entity, area, xCoord, yCoord, units, localIx, localIy, localIxy

    Dim rotationHint As Double
    Dim rotationKnown As Boolean
    If model.LocalInertiaIsNearlySquare(localIx, localIy, localIxy, LOCAL_INERTIA_SQUARE_REL_TOLERANCE) Then
        rotationKnown = TryReadFirstStraightEdgeRotation(entity, units, rotationHint)
    End If

    AddConcreteRegion model, area, xCoord, yCoord, localIx, localIy, localIxy, _
        0#, 0#, rotationHint, sourceLayer, EntityHandle(entity), rotationKnown
End Sub

' Читает один арматурный Region. Для арматуры программа хранит площадь и
' эквивалентный круговой диаметр, потому что механика стержня задается As,
' а визуальная оболочка стержня должна оставаться круглой.
Private Sub AddRebarRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, _
        ByVal steelClass As String, ByVal minArea As Double, ByVal units As CUnitSystem, ByVal sourceLayer As String)
    Dim area As Double
    area = InputArea(CDbl(entity.Area), units)
    If area < minArea Then Exit Sub

    Dim centroid As Variant
    centroid = entity.Centroid
    AddRebarRegion model, area, InputLength(CDbl(centroid(LBound(centroid))), units), _
        InputLength(CDbl(centroid(LBound(centroid) + 1)), units), steelClass, sourceLayer, EntityHandle(entity)
End Sub

' Тестовый/подготовленный путь импорта бетонных Region из массива значений.
' Он повторяет контракт живого импорта, но позволяет regression-тестам
' проверять CSectionModel без запуска AutoCAD.
' ДЛЯ ТЕСТОВ
Private Sub AddConcreteRegionArray(ByVal model As CSectionModel, ByVal regions As Variant, ByVal minArea As Double)
    If Not HasArrayRows(regions) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(regions, 1) To UBound(regions, 1)
        Dim area As Double
        area = CDbl(regions(rowIndex, 1))
        If area >= minArea Then
            AddConcreteRegion model, area, CDbl(regions(rowIndex, 2)), CDbl(regions(rowIndex, 3)), _
                OptionalRegionDouble(regions, rowIndex, 4), OptionalRegionDouble(regions, rowIndex, 5), _
                OptionalRegionDouble(regions, rowIndex, 6), OptionalRegionDouble(regions, rowIndex, 8), _
                OptionalRegionDouble(regions, rowIndex, 9), OptionalRegionDouble(regions, rowIndex, 10), _
                "Concrete", OptionalRegionString(regions, rowIndex, 7)
        End If
    Next rowIndex
End Sub

' Тестовый/подготовленный путь импорта арматурных Region из массива значений.
' Используется только для локальных проверок чтения snapshot-подобных данных.
' ДЛЯ ТЕСТОВ
Private Sub AddRebarRegionArray(ByVal model As CSectionModel, ByVal regions As Variant, _
        ByVal steelClass As String, ByVal minArea As Double)
    If Not HasArrayRows(regions) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(regions, 1) To UBound(regions, 1)
        Dim area As Double
        area = CDbl(regions(rowIndex, 1))
        If area >= minArea Then
            AddRebarRegion model, area, CDbl(regions(rowIndex, 2)), CDbl(regions(rowIndex, 3)), _
                steelClass, "Reinf", OptionalRegionString(regions, rowIndex, 4)
        End If
    Next rowIndex
End Sub

' Передает бетонный элемент в CSectionModel, сохраняя реальные A/I отдельно
' от подсказок геометрической оболочки. Именно CSectionModel решает, будет ли
' оболочка Rectangle, Equivalent rectangle или Equivalent square.
Private Sub AddConcreteRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal realWidth As Double, ByVal realHeight As Double, ByVal realRotation As Double, _
        ByVal sourceName As String, ByVal sourceHandle As String, Optional ByVal rotationKnown As Boolean = False)
    Dim commentText As String
    Dim shapeHint As String
    Dim widthHint As Double
    Dim heightHint As Double
    Dim rotationHint As Double

    shapeHint = "Region"
    rotationHint = realRotation
    If HasExplicitRectangleGeometry(area, realWidth, realHeight) Then
        ' Этот путь нужен только для тестовых/подготовленных массивов и
        ' встроенных адаптеров, которые уже знают реальные стороны. Живой
        ' AutoCAD importer Width/Height не вычисляет.
        shapeHint = "Rectangle"
        widthHint = realWidth
        heightHint = realHeight
        commentText = "Region: прямоугольная оболочка была явно передана источником; расчетные A/I хранит CSectionModel."
    ElseIf model.LocalInertiaIsNearlySquare(localIx, localIy, localIxy, LOCAL_INERTIA_SQUARE_REL_TOLERANCE) Then
        commentText = "AutoCAD Region: реальные A/I прочитаны; почти изотропная оболочка строится в CSectionModel по площади и углу простой грани."
    ElseIf IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        commentText = "AutoCAD Region: реальные A/I прочитаны; прямоугольную оболочку по ним строит CSectionModel."
    Else
        commentText = "AutoCAD Region: A/I не получены; CSectionModel применит последний fallback-квадрат по площади."
    End If

    model.AddConcreteElement xCoord, yCoord, area, 1, sourceName, sourceHandle, _
        shapeHint, widthHint, heightHint, rotationHint, commentText, _
        localIx, localIy, localIxy, rotationKnown
End Sub

' Добавляет арматурный Region как расчетный стержень с площадью AutoCAD.
' Диаметр нужен для схемы и AutoCAD export; расчетная площадь не выводится
' обратно из диаметра, а берется из исходной площади Region.
Private Sub AddRebarRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal steelClass As String, _
        ByVal sourceName As String, ByVal sourceHandle As String)
    Dim diameter As Double
    diameter = Sqr(4# * area / GEOM_PI)
    model.AddRebarElement xCoord, yCoord, diameter, area, steelClass, 1, sourceName, sourceHandle, _
        "AutoCAD Region: стержень арматуры; диаметр и радиус восстановлены по площади как эквивалентный круг."
End Sub

' Приводит инерции AutoCAD Region к центральным локальным характеристикам
' элемента. Если AutoCAD не отдал пригодные A/I или матрица физически
' невозможна, значения сбрасываются, а CSectionModel применит последний
' fallback только для такого элемента.
Private Sub ReadCentralRegionInertia(ByVal entity As Object, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal units As CUnitSystem, ByRef localIx As Double, ByRef localIy As Double, ByRef localIxy As Double)
    localIx = 0#
    localIy = 0#
    localIxy = 0#

    On Error Resume Next
    Dim moments As Variant
    moments = entity.MomentOfInertia
    If Err.Number = 0 Then
        localIx = InputFourthPowerLength(CDbl(moments(LBound(moments))), units) - area * yCoord * yCoord
        localIy = InputFourthPowerLength(CDbl(moments(LBound(moments) + 1)), units) - area * xCoord * xCoord
    End If
    Err.Clear

    Dim product As Variant
    product = entity.ProductOfInertia
    If Err.Number = 0 Then
        localIxy = CentralProductOfInertia(product, area, xCoord, yCoord, units)
    End If
    On Error GoTo 0

    If Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        localIx = 0#
        localIy = 0#
        localIxy = 0#
    End If
End Sub

' Преобразует XY-компоненту тензора AutoCAD в принятое ядром интегральное
' произведение Ixy = integral((x-xc)*(y-yc)*dA). AutoCAD возвращает отрицательный
' интеграл относительно начала координат: сначала меняем знак, затем снимаем
' перенос A*xc*yc. Положительная определенность проверяется для всего тензора
' вызывающим методом; она не позволяет выбирать знак или величину Ixy.
Private Function CentralProductOfInertia(ByVal product As Variant, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal units As CUnitSystem) As Double
    Dim rawValue As Double
    If IsArray(product) Then
        rawValue = InputFourthPowerLength(CDbl(product(LBound(product))), units)
    Else
        rawValue = InputFourthPowerLength(CDbl(product), units)
    End If
    CentralProductOfInertia = -rawValue - area * xCoord * yCoord
End Function

' Для почти изотропного Region тензор инерции не задает направления.
' В этом единственном случае importer дешево берет угол первой прямой грани:
' без проверки прямоугольности, без BoundingBox и без анализа всего контура.
Private Function TryReadFirstStraightEdgeRotation(ByVal entity As Object, _
        ByVal units As CUnitSystem, ByRef rotation As Double) As Boolean
    Dim tempEntity As Object
    Dim pieces As Variant

    On Error GoTo Failed
    Set tempEntity = entity.Copy
    pieces = tempEntity.Explode
    DeleteAcadObjectQuietly tempEntity
    mEdgeProbeCount = mEdgeProbeCount + 1

    If Not IsArray(pieces) Then GoTo Cleanup

    Dim i As Long
    For i = LBound(pieces) To UBound(pieces)
        If IsLineEntity(pieces(i)) Then
            If TryReadLineDirection(pieces(i), units, rotation) Then
                TryReadFirstStraightEdgeRotation = True
                Exit For
            End If
        End If
    Next i

Cleanup:
    DeleteExplodedPiecesQuietly pieces
    Exit Function

Failed:
    On Error Resume Next
    DeleteAcadObjectQuietly tempEntity
    DeleteExplodedPiecesQuietly pieces
End Function

' Читает направление одной прямой AutoCAD-линии. Длина нужна только как
' фильтр от вырожденных сегментов; форму Region по этой линии не выводим.
Private Function TryReadLineDirection(ByVal lineEntity As Object, ByVal units As CUnitSystem, _
        ByRef rotation As Double) As Boolean
    Dim startPoint As Variant
    Dim endPoint As Variant
    startPoint = lineEntity.StartPoint
    endPoint = lineEntity.EndPoint

    Dim x1 As Double
    Dim y1 As Double
    Dim x2 As Double
    Dim y2 As Double
    x1 = InputLength(CDbl(startPoint(LBound(startPoint))), units)
    y1 = InputLength(CDbl(startPoint(LBound(startPoint) + 1)), units)
    x2 = InputLength(CDbl(endPoint(LBound(endPoint))), units)
    y2 = InputLength(CDbl(endPoint(LBound(endPoint) + 1)), units)

    Dim dx As Double
    Dim dy As Double
    dx = x2 - x1
    dy = y2 - y1
    If Sqr(dx * dx + dy * dy) <= GEOM_TOLERANCE Then Exit Function

    rotation = Atn2Local(dy, dx)
    TryReadLineDirection = True
End Function

' Проверяет, является ли фрагмент Explode простой прямой линией.
' Это вспомогательная COM-проверка для изотропных Region, где тензор инерции
' не задает ориентацию оболочки.
Private Function IsLineEntity(ByVal entity As Object) As Boolean
    On Error Resume Next
    IsLineEntity = (StrComp(CStr(entity.ObjectName), "AcDbLine", vbTextCompare) = 0)
    On Error GoTo 0
End Function

' Удаляет временные объекты AutoCAD, созданные при разложении копии Region.
Private Sub DeleteExplodedPiecesQuietly(ByVal pieces As Variant)
    On Error Resume Next
    If IsArray(pieces) Then
        Dim i As Long
        For i = LBound(pieces) To UBound(pieces)
            DeleteAcadObjectQuietly pieces(i)
        Next i
    End If
    On Error GoTo 0
End Sub

' Безопасно удаляет временный AutoCAD-объект. Ошибки удаления намеренно
' подавляются, чтобы диагностика импорта не ломалась из-за мусорной копии
' Region или уже удаленного фрагмента Explode.
Private Sub DeleteAcadObjectQuietly(ByVal entity As Object)
    On Error Resume Next
    If Not entity Is Nothing Then entity.Delete
    On Error GoTo 0
End Sub

' Проверяет явно переданную прямоугольную оболочку в тестовых/подготовленных
' массивах. Live AutoCAD import не использует BoundingBox и не попадает сюда
' с вычисленными Width/Height.
' ДЛЯ ТЕСТОВ
Private Function HasExplicitRectangleGeometry(ByVal area As Double, ByVal width As Double, _
        ByVal height As Double) As Boolean
    If area <= GEOM_TOLERANCE Or width <= GEOM_TOLERANCE Or height <= GEOM_TOLERANCE Then Exit Function

    Dim rectangleArea As Double
    rectangleArea = width * height
    If rectangleArea <= GEOM_TOLERANCE Then Exit Function

    HasExplicitRectangleGeometry = _
        (Abs(rectangleArea - area) / area <= EXPLICIT_RECTANGLE_AREA_REL_TOLERANCE)
End Function

' Матрица собственных инерций должна быть положительно определенной, иначе
' эквивалентный прямоугольник по A/I физически не восстанавливается надежно.
Private Function IsPositiveDefiniteInertia(ByVal localIx As Double, ByVal localIy As Double, _
        ByVal localIxy As Double) As Boolean
    If localIx <= GEOM_TOLERANCE Or localIy <= GEOM_TOLERANCE Then Exit Function
    IsPositiveDefiniteInertia = (localIx * localIy - localIxy * localIxy > GEOM_TOLERANCE)
End Function

' Проверяет тип AutoCAD-сущности по ObjectName. Импорт сознательно работает
' только с Region, чтобы не смешивать реальные массовые характеристики с
' произвольной графикой чертежа.
Private Function IsRegionEntity(ByVal entity As Object) As Boolean
    On Error Resume Next
    IsRegionEntity = (StrComp(CStr(entity.ObjectName), "AcDbRegion", vbTextCompare) = 0)
    On Error GoTo 0
End Function

' Сравнивает имена слоев без учета регистра и лишних пробелов, чтобы
' пользовательский ввод Config был устойчив к обычным опечаткам форматирования.
Private Function SameLayer(ByVal actualLayer As String, ByVal expectedLayer As String) As Boolean
    SameLayer = (StrComp(Trim$(actualLayer), Trim$(expectedLayer), vbTextCompare) = 0)
End Function

' Возвращает AutoCAD handle исходной сущности для трассировки. Если COM
' не отдал handle, импорт продолжается без него.
Private Function EntityHandle(ByVal entity As Object) As String
    On Error Resume Next
    EntityHandle = CStr(entity.Handle)
    On Error GoTo 0
End Function

' Проверяет результат фильтрации Region по слоям и площади: для принятого
' импорта должны остаться бетон и арматура. Пустой слой получает собственную
' причину, а не передает пустую модель в решатель НДС.
Private Sub ValidateImportedModel(ByVal model As CSectionModel)
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4405, "CAutoCADSectionModelImporter", _
        "Не найдено бетонных AutoCAD Region на заданном слое импорта. Проверьте слой AutoCAD.Common.ConcreteLayer и порог AutoCAD.Import.MinArea на листе Config."
    If model.RebarCount <= 0 Then Err.Raise vbObjectError + 4406, "CAutoCADSectionModelImporter", _
        "Не найдено арматурных AutoCAD Region на заданном слое импорта. Проверьте слой AutoCAD.Common.RebarLayer и порог AutoCAD.Import.MinArea на листе Config."
End Sub

' Проверяет, что двумерный Variant действительно содержит хотя бы одну строку.
' Нужен тестовому импорту из массивов, где пустой массив не должен падать на
' LBound/UBound до валидации модели.
' ДЛЯ ТЕСТОВ
Private Function HasArrayRows(ByVal values As Variant) As Boolean
    On Error GoTo NoRows
    Dim firstRow As Long
    firstRow = LBound(values, 1)
    HasArrayRows = (UBound(values, 1) >= firstRow)
    Exit Function
NoRows:
    HasArrayRows = False
End Function

' Безопасно читает числовую ячейку подготовленного массива Region.
' Отсутствующая, пустая или выходящая за границы колонка трактуется как 0,
' чтобы тестовые массивы могли задавать только нужные поля.
' ДЛЯ ТЕСТОВ
Private Function OptionalRegionDouble(ByVal values As Variant, ByVal rowIndex As Long, _
        ByVal colIndex As Long) As Double
    On Error GoTo MissingValue
    If colIndex > UBound(values, 2) Then GoTo MissingValue
    If Len(Trim$(CStr(values(rowIndex, colIndex)))) = 0 Then GoTo MissingValue
    OptionalRegionDouble = CDbl(values(rowIndex, colIndex))
    Exit Function
MissingValue:
    OptionalRegionDouble = 0#
End Function

' Безопасно читает текстовую ячейку подготовленного массива Region.
' Отсутствующие поля превращаются в пустую строку, например когда source handle
' не нужен конкретному regression-тесту.
' ДЛЯ ТЕСТОВ
Private Function OptionalRegionString(ByVal values As Variant, ByVal rowIndex As Long, _
        ByVal colIndex As Long) As String
    On Error GoTo MissingValue
    If colIndex > UBound(values, 2) Then GoTo MissingValue
    OptionalRegionString = CStr(values(rowIndex, colIndex))
    Exit Function
MissingValue:
    OptionalRegionString = vbNullString
End Function

' Возвращает AutoCAD-длину во внутренних миллиметрах. Параметр units оставлен
' в сигнатуре для единообразия, но импорт DWG уже работает в мм по контракту.
Private Function InputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    ' AutoCAD Region импортируется в миллиметрах независимо от пользовательской
    ' INPUT-единицы длины. Иначе смена Units.Length.Input масштабировала бы
    ' координаты чертежа и сдвигала imported snapshot относительно AutoCAD.
    InputLength = value
End Function

' Возвращает AutoCAD-площадь во внутренних мм2. Пользовательские Input-единицы
' не должны масштабировать геометрию, считанную из чертежа.
Private Function InputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    ' Площадь AutoCAD Region уже считается в мм2 по принятому контракту импорта.
    InputArea = value
End Function

' Возвращает AutoCAD-момент инерции во внутренних мм4. Это сохраняет
' координатную совместимость импорта и обратного AutoCAD export.
Private Function InputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    ' Моменты инерции AutoCAD Region уже приходят в мм4 по принятому контракту
    ' импорта; пользовательские INPUT-единицы не должны менять геометрию DWG.
    InputFourthPowerLength = value
End Function

' Локальная реализация atan2 для VBA. Нужна, чтобы получить устойчивый угол
' прямой грани во всех квадрантах без зависимости от внешних библиотек.
Private Function Atn2Local(ByVal y As Double, ByVal x As Double) As Double
    If Abs(x) <= GEOM_TOLERANCE Then
        If y > 0# Then
            Atn2Local = GEOM_PI / 2#
        ElseIf y < 0# Then
            Atn2Local = -GEOM_PI / 2#
        Else
            Atn2Local = 0#
        End If
    Else
        Atn2Local = Atn(y / x)
        If x < 0# Then
            If y >= 0# Then
                Atn2Local = Atn2Local + GEOM_PI
            Else
                Atn2Local = Atn2Local - GEOM_PI
            End If
        End If
    End If
End Function

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Тесты собирают AutoCAD-import модель из массивов, чтобы проверить фильтрацию,
' заполнение CSectionModel и расчетные поля без запуска AutoCAD COM.

' ДЛЯ ТЕСТОВ
' Создает импортированную модель по двум массивам Region и прогоняет ту же
' валидацию, что живой импорт из ModelSpace.
Public Function BuildFromRegionArrays(ByVal concreteRegions As Variant, ByVal rebarRegions As Variant, _
        ByVal steelClass As String, Optional ByVal minArea As Double = 0.000001) As CSectionModel
    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"
    mEdgeProbeCount = 0

    AddConcreteRegionArray model, concreteRegions, minArea
    AddRebarRegionArray model, rebarRegions, steelClass, minArea
    model.ApplyAverageRotationToEquivalentAreaFallbacks
    ValidateImportedModel model

    Set BuildFromRegionArrays = model
End Function

' ДЛЯ ТЕСТОВ
' Возвращает число обращений к Explode за направлением простой грани.
' Производственный расчет это значение не использует.
Public Property Get DebugEdgeProbeCount() As Long
    DebugEdgeProbeCount = mEdgeProbeCount
End Property

