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

Public Function ImportFromActiveDocument(ByVal settings As CSystemSettingsReader, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If settings Is Nothing Then Err.Raise vbObjectError + 4400, "CAutoCADSectionModelImporter", "Настройки System не переданы."

    Dim acad As Object
    On Error Resume Next
    Set acad = GetObject(, "AutoCAD.Application")
    On Error GoTo 0
    If acad Is Nothing Then Err.Raise vbObjectError + 4401, "CAutoCADSectionModelImporter", _
        "AutoCAD не найден. Откройте AutoCAD с активным чертежом или выберите Geometry.Source = Generated."

    Dim doc As Object
    Set doc = acad.ActiveDocument
    If doc Is Nothing Then Err.Raise vbObjectError + 4402, "CAutoCADSectionModelImporter", _
        "В AutoCAD нет активного чертежа."

    Set ImportFromActiveDocument = ImportFromModelSpace(doc.ModelSpace, _
        settings.GetString("AutoCAD.Import.ConcreteLayer", "Concrete"), _
        settings.GetString("AutoCAD.Import.RebarLayer", "Reinf"), _
        settings.GetString("Steel.RebarProfile", "Ribbed"), _
        InputArea(settings.GetDouble("AutoCAD.Import.MinArea", 0.000001), units), units)
End Function

Public Function ImportFromModelSpace(ByVal modelSpace As Object, ByVal concreteLayer As String, _
        ByVal rebarLayer As String, ByVal steelClass As String, ByVal minArea As Double, Optional ByVal units As CUnitSystem = Nothing) As CSectionModel
    If modelSpace Is Nothing Then Err.Raise vbObjectError + 4403, "CAutoCADSectionModelImporter", "ModelSpace AutoCAD не передан."

    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    Dim entity As Object
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

' Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек.
Public Function BuildFromRegionArrays(ByVal concreteRegions As Variant, ByVal rebarRegions As Variant, _
        ByVal steelClass As String, Optional ByVal minArea As Double = 0.000001) As CSectionModel
    Dim model As CSectionModel
    Set model = New CSectionModel
    model.SourceType = "AutoCADImport"

    AddConcreteRegionArray model, concreteRegions, minArea
    AddRebarRegionArray model, rebarRegions, steelClass, minArea
    ValidateImportedModel model

    Set BuildFromRegionArrays = model
End Function

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
    AddConcreteRegion model, area, xCoord, yCoord, localIx, localIy, localIxy, entity.Layer, EntityHandle(entity)
End Sub

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

Private Sub AddConcreteRegionArray(ByVal model As CSectionModel, ByVal regions As Variant, ByVal minArea As Double)
    If Not HasArrayRows(regions) Then Exit Sub

    Dim rowIndex As Long
    For rowIndex = LBound(regions, 1) To UBound(regions, 1)
        Dim area As Double
        area = CDbl(regions(rowIndex, 1))
        If area >= minArea Then
            AddConcreteRegion model, area, CDbl(regions(rowIndex, 2)), CDbl(regions(rowIndex, 3)), _
                OptionalRegionDouble(regions, rowIndex, 4), OptionalRegionDouble(regions, rowIndex, 5), _
                OptionalRegionDouble(regions, rowIndex, 6), "Concrete", OptionalRegionString(regions, rowIndex, 7)
        End If
    Next rowIndex
End Sub

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

Private Sub AddConcreteRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double, _
        ByVal sourceName As String, ByVal sourceHandle As String)
    Dim width As Double
    Dim height As Double
    EquivalentRectangleDimensions area, localIx, localIy, width, height
    If localIx <= 0# Then localIx = area * height * height / 12#
    If localIy <= 0# Then localIy = area * width * width / 12#

    model.AddConcreteElement xCoord, yCoord, area, 1, sourceName, sourceHandle, _
        "EquivalentRectangle", width, height, 0#, "AutoCAD Region; equivalent rectangle for export", _
        localIx, localIy, localIxy
End Sub

Private Sub AddRebarRegion(ByVal model As CSectionModel, ByVal area As Double, _
        ByVal xCoord As Double, ByVal yCoord As Double, ByVal steelClass As String, _
        ByVal sourceName As String, ByVal sourceHandle As String)
    Dim diameter As Double
    diameter = Sqr(4# * area / GEOM_PI)
    model.AddRebarElement xCoord, yCoord, diameter, area, steelClass, 1, sourceName, sourceHandle, "AutoCAD Region"
End Sub

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
        If IsArray(product) Then
            localIxy = InputFourthPowerLength(CDbl(product(LBound(product))), units) - area * xCoord * yCoord
        Else
            localIxy = InputFourthPowerLength(CDbl(product), units) - area * xCoord * yCoord
        End If
    End If
    On Error GoTo 0

    If localIx <= 0# Or localIy <= 0# Then
        localIx = 0#
        localIy = 0#
        localIxy = 0#
    End If
End Sub

Private Sub EquivalentRectangleDimensions(ByVal area As Double, ByVal localIx As Double, _
        ByVal localIy As Double, ByRef width As Double, ByRef height As Double)
    If area <= 0# Then Err.Raise vbObjectError + 4404, "CAutoCADSectionModelImporter", "Площадь Region должна быть положительной."

    If localIx > 0# And localIy > 0# Then
        height = Sqr(12# * localIx / area)
        width = Sqr(12# * localIy / area)
    Else
        width = Sqr(area)
        height = width
    End If

    If width <= 0# Or height <= 0# Then
        width = Sqr(area)
        height = width
    End If
End Sub

Private Function IsRegionEntity(ByVal entity As Object) As Boolean
    On Error Resume Next
    IsRegionEntity = (StrComp(CStr(entity.ObjectName), "AcDbRegion", vbTextCompare) = 0)
    On Error GoTo 0
End Function

Private Function SameLayer(ByVal actualLayer As String, ByVal expectedLayer As String) As Boolean
    SameLayer = (StrComp(Trim$(actualLayer), Trim$(expectedLayer), vbTextCompare) = 0)
End Function

Private Function EntityHandle(ByVal entity As Object) As String
    On Error Resume Next
    EntityHandle = CStr(entity.Handle)
    On Error GoTo 0
End Function

' Проверяет входные данные и прерывает выполнение понятной ошибкой, если расчетный контракт нарушен.
Private Sub ValidateImportedModel(ByVal model As CSectionModel)
    If model.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4405, "CAutoCADSectionModelImporter", _
        "Не найдено бетонных AutoCAD Region на заданном слое импорта."
    If model.RebarCount <= 0 Then Err.Raise vbObjectError + 4406, "CAutoCADSectionModelImporter", _
        "Не найдено арматурных AutoCAD Region на заданном слое импорта."
End Sub

Private Function HasArrayRows(ByVal values As Variant) As Boolean
    On Error GoTo NoRows
    Dim firstRow As Long
    firstRow = LBound(values, 1)
    HasArrayRows = (UBound(values, 1) >= firstRow)
    Exit Function
NoRows:
    HasArrayRows = False
End Function

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

Private Function OptionalRegionString(ByVal values As Variant, ByVal rowIndex As Long, _
        ByVal colIndex As Long) As String
    On Error GoTo MissingValue
    If colIndex > UBound(values, 2) Then GoTo MissingValue
    OptionalRegionString = CStr(values(rowIndex, colIndex))
    Exit Function
MissingValue:
    OptionalRegionString = vbNullString
End Function

Private Function InputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputLength = value Else InputLength = units.InputLengthToInternal(value)
End Function

Private Function InputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputArea = value Else InputArea = units.InputAreaToInternal(value)
End Function

Private Function InputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then InputFourthPowerLength = value Else InputFourthPowerLength = units.InputFourthPowerLengthToInternal(value)
End Function

