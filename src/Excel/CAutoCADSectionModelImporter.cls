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

' Читает активный документ AutoCAD через COM и запускает общий импорт
' ModelSpace. Метод остается тонкой оболочкой: проверяет доступность AutoCAD,
' берет слои из Config и не выполняет расчетную интерпретацию сечения.
Public Function ImportFromActiveDocument(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки Config не переданы."

    Dim acad As Object
    On Error Resume Next
    Set acad = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    If acad Is Nothing Then Err.Raise vbObjectError + 4401, "CAutoCADSectionModelImporter", _
        "AutoCAD не открыт. Откройте AutoCAD с нужным чертежом и повторите импорт."

    Dim doc As Object
    On Error Resume Next
    Set doc = acad.ActiveDocument
    On Error GoTo 0
    If doc Is Nothing Then Err.Raise vbObjectError + 4402, "CAutoCADSectionModelImporter", _
        "AutoCAD открыт, но активный чертеж не найден. Откройте или создайте чертеж и повторите импорт."

    Set ImportFromActiveDocument = ImportConfiguredModelSpace(doc.ModelSpace, settings, units)
End Function

' Подготавливает параметры импорта из прочитанного Config и передает ModelSpace
' общему обходу Region. Порог из пользовательской INPUT-площади переводится
' в мм2; координаты и геометрические свойства самих Region уже заданы в мм.
' Получение активного документа остается только в ImportFromActiveDocument.
Public Function ImportConfiguredModelSpace(ByVal modelSpace As Object, ByVal settings As CSystemSettingsReader, _
        Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки Config не переданы."
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
    Set ImportConfiguredModelSpace = ImportFromModelSpace(modelSpace, _
        settings.GetRequiredString("AutoCAD.Import.ConcreteLayer"), _
        settings.GetRequiredString("AutoCAD.Import.RebarLayer"), REBAR_STEEL_CLASS, minArea, importUnits)
    Exit Function
InvalidMinimumArea:
    Err.Raise vbObjectError + 4404, "CAutoCADSectionModelImporter", _
        settings.InputErrorMessage("AutoCAD.Import.MinArea", "Площадь слишком велика для перевода в мм2.", _
            "Уменьшите число и проверьте выбранную единицу INPUT Area.")
End Function

' Проходит по ModelSpace и отбирает только Region на пользовательских слоях
' бетона и арматуры. На выходе возвращает полностью заполненный CSectionModel,
' пригодный для того же расчетного pipeline, что и встроенные генераторы.
Public Function ImportFromModelSpace(ByVal modelSpace As Object, ByVal concreteLayer As String, _
        ByVal rebarLayer As String, ByVal steelClass As String, ByVal minArea As Double, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If modelSpace Is Nothing Then Err.Raise vbObjectError + 4403, "CAutoCADSectionModelImporter", "ModelSpace AutoCAD не передан."

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    Dim entity As Object
    mEdgeProbeCount = 0
    For Each entity In modelSpace
        If IsRegionEntity(entity) Then
            If SameLayer(entity.Layer, concreteLayer) Then
                AddConcreteRegionEntity model, entity, minArea, units
            ElseIf SameLayer(entity.Layer, rebarLayer) Then
                AddRebarRegionEntity model, entity, steelClass, minArea, units
            End If
        End If
    Next entity

    ValidateImportedModel model
    Set ImportFromModelSpace = model
End Function

' Читает один бетонный Region из AutoCAD: площадь, центр тяжести, реальные
' центральные инерции и, только для почти изотропного случая, подсказку угла
' по первой прямой грани. Width/Height здесь не восстанавливаются.
Private Sub AddConcreteRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, ByVal minArea As Double, ByVal units As CUnitSystem)
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
    If model.LocalInertiaIsNearlySquare(localIx, localIy, localIxy, LOCAL_INERTIA_SQUARE_REL_TOLERANCE) Then
        If Not TryReadFirstStraightEdgeRotation(entity, units, rotationHint) Then rotationHint = 0#
    End If

    AddConcreteRegion model, area, xCoord, yCoord, localIx, localIy, localIxy, _
        0#, 0#, rotationHint, entity.Layer, EntityHandle(entity)
End Sub

' Читает один арматурный Region. Для арматуры программа хранит площадь и
' эквивалентный круговой диаметр, потому что механика стержня задается As,
' а визуальная оболочка стержня должна оставаться круглой.
Private Sub AddRebarRegionEntity(ByVal model As CSectionModel, ByVal entity As Object, _
        ByVal steelClass As String, ByVal minArea As Double, ByVal units As CUnitSystem)
    Dim area As Double
    area = InputArea(CDbl(entity.Area), units)
    If area < minArea Then Exit Sub

    Dim centroid As Variant
    centroid = entity.Centroid
    AddRebarRegion model, area, InputLength(CDbl(centroid(LBound(centroid))), units), _
        InputLength(CDbl(centroid(LBound(centroid) + 1)), units), steelClass, entity.Layer, EntityHandle(entity)
End Sub

' Тестовый/подготовленный путь импорта бетонных Region из массива значений.
' Он повторяет контракт живого импорта, но позволяет regression-тестам
' проверять CSectionModel без запуска AutoCAD.
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
        ByVal sourceName As String, ByVal sourceHandle As String)
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
        localIx, localIy, localIxy
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
        localIxy = BestCentralProductOfInertia(product, area, xCoord, yCoord, units, localIx, localIy)
    End If
    On Error GoTo 0

    If Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        localIx = 0#
        localIy = 0#
        localIxy = 0#
    End If
End Sub

' Выбирает физически пригодный центральный Ixy из вариантов AutoCAD.
' В разных версиях COM ProductOfInertia может приходить со знаком или уже в
' центральном виде, поэтому проверяем несколько кандидатов и оставляем тот,
' который дает положительно определенную локальную матрицу.
Private Function BestCentralProductOfInertia(ByVal product As Variant, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal units As CUnitSystem, _
        ByVal localIx As Double, ByVal localIy As Double) As Double
    Dim bestFound As Boolean
    Dim bestValue As Double
    Dim bestScore As Double
    Dim shiftValue As Double
    shiftValue = area * xCoord * yCoord
    Dim inertiaScale As Double
    inertiaScale = GeomMax(Sqr(Abs(localIx * localIy)), 1#)

    Dim firstIndex As Long
    If IsArray(product) Then
        firstIndex = LBound(product)
    Else
        firstIndex = 0
    End If

    Dim rawValue As Double
    If IsArray(product) Then
        rawValue = InputFourthPowerLength(CDbl(product(firstIndex)), units)
    Else
        rawValue = InputFourthPowerLength(CDbl(product), units)
    End If

    TestProductCandidate rawValue - shiftValue, localIx, localIy, bestFound, bestValue, bestScore
    TestProductCandidate -rawValue - shiftValue, localIx, localIy, bestFound, bestValue, bestScore
    TestProductCandidate rawValue + shiftValue, localIx, localIy, bestFound, bestValue, bestScore
    TestProductCandidate -rawValue + shiftValue, localIx, localIy, bestFound, bestValue, bestScore
    If Abs(shiftValue) <= 0.05 * inertiaScale Then
        TestProductCandidate rawValue, localIx, localIy, bestFound, bestValue, bestScore
    End If

    If bestFound Then BestCentralProductOfInertia = bestValue
End Function

' Кандидат Ixy пригоден, только если вся локальная матрица инерций физически
' положительно определена. Минимальный по модулю Ixy обычно соответствует
' центральной системе AutoCAD Region без паразитного переноса.
Private Sub TestProductCandidate(ByVal candidateIxy As Double, ByVal localIx As Double, _
        ByVal localIy As Double, ByRef bestFound As Boolean, _
        ByRef bestValue As Double, ByRef bestScore As Double)
    If Not IsPositiveDefiniteInertia(localIx, localIy, candidateIxy) Then Exit Sub

    Dim scoreValue As Double
    scoreValue = Abs(candidateIxy)
    If Not bestFound Or scoreValue < bestScore Then
        bestFound = True
        bestValue = candidateIxy
        bestScore = scoreValue
    End If
End Sub

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
        "Не найдено бетонных AutoCAD Region на заданном слое импорта. Проверьте слой AutoCAD.Import.ConcreteLayer и порог AutoCAD.Import.MinArea на листе Config."
    If model.RebarCount <= 0 Then Err.Raise vbObjectError + 4406, "CAutoCADSectionModelImporter", _
        "Не найдено арматурных AutoCAD Region на заданном слое импорта. Проверьте слой AutoCAD.Import.RebarLayer и порог AutoCAD.Import.MinArea на листе Config."
End Sub

' Проверяет, что двумерный Variant действительно содержит хотя бы одну строку.
' Нужен тестовому импорту из массивов, где пустой массив не должен падать на
' LBound/UBound до валидации модели.
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

