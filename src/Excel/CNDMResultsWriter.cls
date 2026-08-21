VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNDMResultsWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ELEMENT_RESULT_COLUMNS As Long = 6
Private Const GEOMETRY_COLUMNS As Long = 16
Private Const SECTION_PROPERTY_COLUMNS As Long = 6
Private Const ANNOTATION_COLUMNS As Long = 14
Private Const PHYSICAL_TOLERANCE As Double = 0.000000001

Public Sub WriteResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal concreteClass As String, Optional ByVal units As CUnitSystem = Nothing)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4450, "CNDMResultsWriter", "Книга Excel не передана."
    If section Is Nothing Then Err.Raise vbObjectError + 4451, "CNDMResultsWriter", "Модель сечения не передана."
    If concrete Is Nothing Then Err.Raise vbObjectError + 4452, "CNDMResultsWriter", "Материал бетона не передан."
    If steel Is Nothing Then Err.Raise vbObjectError + 4453, "CNDMResultsWriter", "Материал арматуры не передан."
    If batch Is Nothing Then Err.Raise vbObjectError + 4454, "CNDMResultsWriter", "Пакетный расчетчик не передан."

    ClearResults workbook

    Dim runID As String
    runID = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    WriteSectionGeometry workbook, section, runID, concreteClass, units
    WriteElementResults workbook, section, concrete, steel, batch, runID, units
    WriteSectionProperties workbook, section, concrete, steel, batch, runID, units
    WriteSectionAnnotations workbook, section, runID, units
End Sub

Public Sub ClearResults(ByVal workbook As Object)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4455, "CNDMResultsWriter", "Книга Excel не передана."

    ClearNamedRegion workbook, "rngNDMElementResults"
    ClearNamedRegion workbook, "rngNDMSectionGeometry"
    ClearNamedRegion workbook, "rngNDMSectionProperties"
    ClearNamedRegion workbook, "rngNDMSectionAnnotations"
End Sub

Private Sub ClearNamedRegion(ByVal workbook As Object, ByVal rangeName As String)
    On Error GoTo MissingRange
    workbook.Names.Item(rangeName).RefersToRange.CurrentRegion.ClearContents
MissingRange:
End Sub

Private Sub WriteElementResults(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal runID As String, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMElementResults").RefersToRange

    Dim elementCount As Long
    elementCount = section.ConcreteCount + section.RebarCount

    Dim activeCombinationCount As Long
    Dim i As Long
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then activeCombinationCount = activeCombinationCount + 1
    Next i

    Dim rowCount As Long
    rowCount = 1 + activeCombinationCount * elementCount
    If rowCount < 1 Then rowCount = 1

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To ELEMENT_RESULT_COLUMNS)
    PutElementResultHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then
            PutCombinationElementResults data, rowIndex, section, concrete, steel, batch, i, runID, units
        End If
    Next i

    target.Resize(rowCount, ELEMENT_RESULT_COLUMNS).Value2 = data
End Sub

Private Sub WriteSectionGeometry(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As String, ByVal concreteClass As String, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange

    Dim rowCount As Long
    rowCount = 1 + section.ConcreteCount + section.RebarCount

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To GEOMETRY_COLUMNS)
    PutGeometryHeaders data, units

    Dim rowIndex As Long
    rowIndex = 2

    Dim i As Long
    For i = 1 To section.ConcreteCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.ConcreteID(i)
        data(rowIndex, 3) = "Concrete"
        data(rowIndex, 4) = concreteClass
        data(rowIndex, 5) = OutputLength(section.ConcreteX(i), units)
        data(rowIndex, 6) = OutputLength(section.ConcreteY(i), units)
        data(rowIndex, 7) = OutputArea(section.ConcreteArea(i), units)
        data(rowIndex, 8) = section.ConcreteShapeType(i)
        data(rowIndex, 9) = OutputLength(section.ConcreteWidth(i), units)
        data(rowIndex, 10) = OutputLength(section.ConcreteHeight(i), units)
        data(rowIndex, 11) = vbNullString
        data(rowIndex, 12) = section.ConcreteRotation(i)
        data(rowIndex, 13) = OutputFourthPowerLength(section.ConcreteLocalIx(i), units)
        data(rowIndex, 14) = OutputFourthPowerLength(section.ConcreteLocalIy(i), units)
        data(rowIndex, 15) = OutputFourthPowerLength(section.ConcreteLocalIxy(i), units)
        data(rowIndex, 16) = section.ConcreteComment(i)
        rowIndex = rowIndex + 1
    Next i

    For i = 1 To section.RebarCount
        data(rowIndex, 1) = runID
        data(rowIndex, 2) = section.RebarID(i)
        data(rowIndex, 3) = "Rebar"
        data(rowIndex, 4) = section.RebarSteelClass(i)
        data(rowIndex, 5) = OutputLength(section.RebarX(i), units)
        data(rowIndex, 6) = OutputLength(section.RebarY(i), units)
        data(rowIndex, 7) = OutputArea(section.RebarArea(i), units)
        data(rowIndex, 8) = "Circle"
        data(rowIndex, 9) = vbNullString
        data(rowIndex, 10) = vbNullString
        data(rowIndex, 11) = OutputLength(section.RebarDiameter(i), units)
        data(rowIndex, 12) = 0#
        data(rowIndex, 13) = 0#
        data(rowIndex, 14) = 0#
        data(rowIndex, 15) = 0#
        data(rowIndex, 16) = section.RebarComment(i)
        rowIndex = rowIndex + 1
    Next i

    target.Resize(rowCount, GEOMETRY_COLUMNS).Value2 = data
End Sub

Private Sub WriteSectionProperties(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal concrete As Object, ByVal steel As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal runID As String, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionProperties").RefersToRange

    Dim rowCount As Long
    rowCount = 1 + 20 + batch.Count * 12

    Dim data() As Variant
    ReDim data(1 To rowCount, 1 To SECTION_PROPERTY_COLUMNS)
    PutSectionPropertyHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    GetConcreteBounds section, minX, maxX, minY, maxY

    Dim props As CSectionPropertiesCalculator
    Set props = New CSectionPropertiesCalculator
    props.CalculateTransformed section, concrete, steel

    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinX", OutputLength(minX, units), OutputLengthUnit(units), "Минимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxX", OutputLength(maxX, units), OutputLengthUnit(units), "Максимальный X расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MinY", OutputLength(minY, units), OutputLengthUnit(units), "Минимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "Bounds.MaxY", OutputLength(maxY, units), OutputLengthUnit(units), "Максимальный Y расчетной модели."
    PutProperty data, rowIndex, runID, "ALL", "CentroidX", OutputLength(props.CentroidX, units), OutputLengthUnit(units), "Центр тяжести приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "CentroidY", OutputLength(props.CentroidY, units), OutputLengthUnit(units), "Центр тяжести приведенного сечения."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalAngle", props.PrincipalAngleRad, "rad", "Угол главных центральных осей."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalI1", OutputFourthPowerLength(props.PrincipalI1, units), OutputLengthUnit(units) & "4", "Главный момент инерции I1."
    PutProperty data, rowIndex, runID, "ALL", "PrincipalI2", OutputFourthPowerLength(props.PrincipalI2, units), OutputLengthUnit(units) & "4", "Главный момент инерции I2."
    PutProperty data, rowIndex, runID, "ALL", "ConcreteCount", section.ConcreteCount, "шт", "Количество бетонных элементов."
    PutProperty data, rowIndex, runID, "ALL", "RebarCount", section.RebarCount, "шт", "Количество элементов арматуры."
    PutProperty data, rowIndex, runID, "ALL", "Output.LengthUnit", OutputLengthUnit(units), "-", "Единица вывода длины snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.AreaUnit", OutputAreaUnit(units), "-", "Единица вывода площади snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.ForceUnit", OutputForceUnit(units), "-", "Единица вывода силы snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.MomentUnit", OutputMomentUnit(units), "-", "Единица вывода момента snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.StressUnit", OutputStressUnit(units), "-", "Единица вывода напряжения snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.CurvatureUnit", OutputCurvatureUnit(units), "-", "Единица вывода кривизны snapshot."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.N", OutputSignN(units), "-", "Пользовательский знак N на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.Mx", OutputSignMx(units), "-", "Пользовательский знак Mx на момент расчета."
    PutProperty data, rowIndex, runID, "ALL", "Output.SignConvention.My", OutputSignMy(units), "-", "Пользовательский знак My на момент расчета."

    Dim i As Long
    For i = 1 To batch.Count
        If batch.StateAvailable(i) Then
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Epsilon0", batch.Epsilon0(i), "-", "Деформация в начале координат расчетной модели."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaX", OutputCurvature(batch.KappaX(i), units), OutputCurvatureUnit(units), "Кривизна KappaX."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "KappaY", OutputCurvature(batch.KappaY(i), units), OutputCurvatureUnit(units), "Кривизна KappaY."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "LoadReferenceX", OutputLength(batch.LoadReferenceX, units), OutputLengthUnit(units), "X точки приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "LoadReferenceY", OutputLength(batch.LoadReferenceY, units), OutputLengthUnit(units), "Y точки приложения нагрузки."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "N", OutputForce(batch.N(i), units), OutputForceUnit(units), "Заданная продольная сила."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxUser", OutputMx(batch.UserMx(i), units), OutputMomentUnit(units), "Пользовательский Mx."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyUser", OutputMy(batch.UserMy(i), units), OutputMomentUnit(units), "Пользовательский My."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MxInternal", OutputMx(batch.Mx(i), units), OutputMomentUnit(units), "Mx после переноса к координатам элементов."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "MyInternal", OutputMy(batch.My(i), units), OutputMomentUnit(units), "My после переноса к координатам элементов."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Nint", OutputForce(batch.Nint(i), units), OutputForceUnit(units), "Внутренняя продольная сила."
            PutProperty data, rowIndex, runID, batch.CombinationID(i), "Status", batch.Status(i), "-", "Сводный статус сочетания."
        End If
    Next i

    target.Resize(rowCount, SECTION_PROPERTY_COLUMNS).Value2 = data
End Sub

Private Sub WriteSectionAnnotations(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal runID As String, ByVal units As CUnitSystem)
    Dim target As Object
    Set target = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange

    Dim data() As Variant
    ReDim data(1 To 256, 1 To ANNOTATION_COLUMNS)
    PutAnnotationHeaders data

    Dim rowIndex As Long
    rowIndex = 2

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    GetSectionBounds section, minX, maxX, minY, maxY
    Dim outMinX As Double, outMaxX As Double, outMinY As Double, outMaxY As Double
    outMinX = OutputLength(minX, units)
    outMaxX = OutputLength(maxX, units)
    outMinY = OutputLength(minY, units)
    outMaxY = OutputLength(maxY, units)

    Dim dx As Double, dy As Double
    dx = outMaxX - outMinX
    dy = outMaxY - outMinY
    Dim offset As Double
    offset = 0.12 * MaxDouble(dx, dy)
    If offset <= 0# Then offset = 20#

    If Not HasGeneratedSectionAnnotations(section) Then
        target.Resize(UBound(data, 1), ANNOTATION_COLUMNS).Value2 = data
        Exit Sub
    End If

    WriteContourAnnotations data, rowIndex, runID, section, units, outMinX, outMaxX, outMinY, outMaxY
    If IsCircleModel(section) Then
        PutDimensionAnnotation data, rowIndex, runID, "DIM_CIRCLE_D", outMinX, outMinY, outMinX, outMaxY, -1#, 0#, offset, _
            "D=" & FormatNumberInvariant(dy) & " " & OutputLengthUnit(units), dy, OutputLengthUnit(units), "Diameter dimension for generated circle."
        WriteRebarGroupAnnotations data, rowIndex, runID, section, units, outMinX, outMaxX, outMinY, outMaxY, offset
        target.Resize(UBound(data, 1), ANNOTATION_COLUMNS).Value2 = data
        Exit Sub
    End If

    PutDimensionAnnotation data, rowIndex, runID, "DIM_B", outMinX, outMinY, outMaxX, outMinY, 0#, -1#, offset, _
        "B=" & FormatNumberInvariant(dx) & " " & OutputLengthUnit(units), dx, OutputLengthUnit(units), "Overall concrete width by Bounds."
    PutDimensionAnnotation data, rowIndex, runID, "DIM_H", outMinX, outMinY, outMinX, outMaxY, -1#, 0#, offset, _
        "H=" & FormatNumberInvariant(dy) & " " & OutputLengthUnit(units), dy, OutputLengthUnit(units), "Overall concrete height by Bounds."

    WriteDerivedConcreteDimensions data, rowIndex, runID, section, units, outMinX, outMaxX, outMinY, outMaxY, offset
    WriteRebarGroupAnnotations data, rowIndex, runID, section, units, outMinX, outMaxX, outMinY, outMaxY, offset

    target.Resize(UBound(data, 1), ANNOTATION_COLUMNS).Value2 = data
End Sub

Private Sub PutCombinationElementResults(ByRef data() As Variant, ByRef rowIndex As Long, _
        ByVal section As CSectionModel, ByVal concrete As Object, ByVal steel As Object, _
        ByVal batch As CBatchSectionCalculator, ByVal combinationIndex As Long, _
        ByVal runID As String, ByVal units As CUnitSystem)
    Dim i As Long
    Dim strain As Double
    Dim stress As Double

    For i = 1 To section.ConcreteCount
        strain = batch.Epsilon0(combinationIndex) + batch.KappaX(combinationIndex) * section.ConcreteY(i) + _
            batch.KappaY(combinationIndex) * section.ConcreteX(i)
        stress = concrete.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), section.ConcreteID(i), _
            strain, stress, PhysicalState("Concrete", strain, stress), units
        rowIndex = rowIndex + 1
    Next i

    For i = 1 To section.RebarCount
        strain = batch.Epsilon0(combinationIndex) + batch.KappaX(combinationIndex) * section.RebarY(i) + _
            batch.KappaY(combinationIndex) * section.RebarX(i)
        stress = steel.GetStress(strain)
        PutElementRow data, rowIndex, runID, batch.CombinationID(combinationIndex), section.RebarID(i), _
            strain, stress, PhysicalState("Rebar", strain, stress), units
        rowIndex = rowIndex + 1
    Next i
End Sub

Private Sub PutElementRow(ByRef data() As Variant, ByVal rowIndex As Long, ByVal runID As String, _
        ByVal combinationID As String, ByVal elementID As String, _
        ByVal strain As Double, ByVal stress As Double, ByVal state As String, ByVal units As CUnitSystem)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = combinationID
    data(rowIndex, 3) = elementID
    data(rowIndex, 4) = strain
    data(rowIndex, 5) = OutputStress(stress, units)
    data(rowIndex, 6) = state
End Sub

Private Function PhysicalState(ByVal materialType As String, ByVal strain As Double, ByVal stress As Double) As String
    If Abs(stress) <= PHYSICAL_TOLERANCE Then
        If StrComp(materialType, "Concrete", vbTextCompare) = 0 And strain > PHYSICAL_TOLERANCE Then
            PhysicalState = "InactiveTensionConcrete"
        Else
            PhysicalState = "NearZero"
        End If
    ElseIf stress < 0# Then
        PhysicalState = "Compression"
    Else
        PhysicalState = "Tension"
    End If
End Function

Private Sub PutElementResultHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "ElementID"
    data(1, 4) = "Strain"
    data(1, 5) = "Stress, " & OutputStressUnit(units)
    data(1, 6) = "PhysicalState"
End Sub

Private Sub PutGeometryHeaders(ByRef data() As Variant, ByVal units As CUnitSystem)
    data(1, 1) = "RunID"
    data(1, 2) = "ElementID"
    data(1, 3) = "MaterialType"
    data(1, 4) = "MaterialClass"
    data(1, 5) = "X, " & OutputLengthUnit(units)
    data(1, 6) = "Y, " & OutputLengthUnit(units)
    data(1, 7) = "Area, " & OutputAreaUnit(units)
    data(1, 8) = "ShapeType"
    data(1, 9) = "Width, " & OutputLengthUnit(units)
    data(1, 10) = "Height, " & OutputLengthUnit(units)
    data(1, 11) = "Diameter, " & OutputLengthUnit(units)
    data(1, 12) = "Rotation, rad"
    data(1, 13) = "LocalIx, " & OutputLengthUnit(units) & "4"
    data(1, 14) = "LocalIy, " & OutputLengthUnit(units) & "4"
    data(1, 15) = "LocalIxy, " & OutputLengthUnit(units) & "4"
    data(1, 16) = "Comment"
End Sub

Private Sub PutSectionPropertyHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "LoadCase"
    data(1, 3) = "Parameter"
    data(1, 4) = "Value"
    data(1, 5) = "Unit"
    data(1, 6) = "Comment"
End Sub

Private Sub PutAnnotationHeaders(ByRef data() As Variant)
    data(1, 1) = "RunID"
    data(1, 2) = "AnnotationType"
    data(1, 3) = "AnnotationID"
    data(1, 4) = "StartX"
    data(1, 5) = "StartY"
    data(1, 6) = "EndX"
    data(1, 7) = "EndY"
    data(1, 8) = "OutsideNormalX"
    data(1, 9) = "OutsideNormalY"
    data(1, 10) = "Offset"
    data(1, 11) = "Text"
    data(1, 12) = "Value"
    data(1, 13) = "Unit"
    data(1, 14) = "Comment"
End Sub

Private Sub PutProperty(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal loadCase As String, ByVal parameter As String, ByVal value As Variant, _
        ByVal unitText As String, ByVal comment As String)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = loadCase
    data(rowIndex, 3) = parameter
    data(rowIndex, 4) = value
    data(rowIndex, 5) = unitText
    data(rowIndex, 6) = comment
    rowIndex = rowIndex + 1
End Sub

Private Sub PutDimensionAnnotation(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal offset As Double, ByVal textValue As String, ByVal numericValue As Double, _
        ByVal unitText As String, ByVal comment As String)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = "DIMENSION"
    data(rowIndex, 3) = annotationID
    data(rowIndex, 4) = startX
    data(rowIndex, 5) = startY
    data(rowIndex, 6) = endX
    data(rowIndex, 7) = endY
    data(rowIndex, 8) = normalX
    data(rowIndex, 9) = normalY
    data(rowIndex, 10) = offset
    data(rowIndex, 11) = textValue
    data(rowIndex, 12) = numericValue
    data(rowIndex, 13) = unitText
    data(rowIndex, 14) = comment
    rowIndex = rowIndex + 1
End Sub

Private Sub WriteRebarGroupAnnotations(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal section As CSectionModel, ByVal units As CUnitSystem, _
        ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, ByVal offset As Double)
    If section.RebarCount <= 0 Then Exit Sub

    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare

    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")
    stats.CompareMode = vbTextCompare

    Dim i As Long
    For i = 1 To section.RebarCount
        Dim groupName As String
        groupName = RebarGroupName(section.RebarComment(i))
        If Len(groupName) = 0 Then groupName = "Rebar"
        UpdateRebarAnnotationStats stats, groupName, OutputLength(section.RebarX(i), units), OutputLength(section.RebarY(i), units)

        Dim diameterText As String
        diameterText = FormatNumberInvariant(OutputLength(section.RebarDiameter(i), units))
        Dim key As String
        key = groupName & "|" & diameterText
        If groups.Exists(key) Then
            groups(key) = CLng(groups(key)) + 1
        Else
            groups.Add key, 1
        End If
    Next i

    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    labels.CompareMode = vbTextCompare
    Dim keyVariant As Variant
    For Each keyVariant In groups.Keys
        Dim parts As Variant
        parts = Split(CStr(keyVariant), "|")
        Dim labelText As String
        labelText = CStr(groups(keyVariant)) & ChrW$(216) & CStr(parts(1))
        If labels.Exists(CStr(parts(0))) Then
            labels(CStr(parts(0))) = CStr(labels(CStr(parts(0)))) & " + " & labelText
        Else
            labels.Add CStr(parts(0)), labelText
        End If
    Next keyVariant

    Dim topRight As Double, rightTop As Double, hasLShape As Boolean
    DetectLShapeBreak section, units, minX, maxX, minY, maxY, topRight, rightTop, hasLShape

    For Each keyVariant In labels.Keys
        PutRebarAnnotationForGroup data, rowIndex, runID, CStr(keyVariant), CStr(labels(keyVariant)), _
            stats, section, units, minX, maxX, minY, maxY, offset, topRight, rightTop, hasLShape
    Next keyVariant
End Sub

Private Sub WriteDerivedConcreteDimensions(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal section As CSectionModel, ByVal units As CUnitSystem, _
        ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, ByVal offset As Double)
    Dim dx As Double, dy As Double
    dx = maxX - minX
    dy = maxY - minY
    If dx <= 0# Or dy <= 0# Then Exit Sub

    Dim tol As Double
    tol = MaxDouble(dx, dy) * 0.03

    Dim topRight As Double
    topRight = minX
    Dim rightTop As Double
    rightTop = minY

    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim cx As Double, cy As Double, w As Double, h As Double
        cx = OutputLength(section.ConcreteX(i), units)
        cy = OutputLength(section.ConcreteY(i), units)
        w = OutputLength(ConcreteWidth(section, i), units)
        h = OutputLength(ConcreteHeight(section, i), units)
        If cy + h / 2# >= maxY - tol Then topRight = MaxDouble(topRight, cx + w / 2#)
        If cx + w / 2# >= maxX - tol Then rightTop = MaxDouble(rightTop, cy + h / 2#)
    Next i

    If topRight > minX + tol And topRight < maxX - tol Then
        PutDimensionAnnotation data, rowIndex, runID, "DIM_TOP_WIDTH", minX, maxY, topRight, maxY, _
            0#, 1#, offset * 0.55, "B1=" & FormatNumberInvariant(topRight - minX) & " " & OutputLengthUnit(units), _
            topRight - minX, OutputLengthUnit(units), "Top horizontal size detected from concrete bounds."
    End If

    If rightTop > minY + tol And rightTop < maxY - tol Then
        PutDimensionAnnotation data, rowIndex, runID, "DIM_RIGHT_HEIGHT", maxX, minY, maxX, rightTop, _
            1#, 0#, offset * 0.55, "H2=" & FormatNumberInvariant(rightTop - minY) & " " & OutputLengthUnit(units), _
            rightTop - minY, OutputLengthUnit(units), "Right vertical size detected from concrete bounds."
    End If
End Sub

Private Sub UpdateRebarAnnotationStats(ByVal stats As Object, ByVal groupName As String, _
        ByVal xCoord As Double, ByVal yCoord As Double)
    Dim values As Variant
    If stats.Exists(groupName) Then
        values = stats(groupName)
        values(0) = MinDouble(CDbl(values(0)), xCoord)
        values(1) = MaxDouble(CDbl(values(1)), xCoord)
        values(2) = MinDouble(CDbl(values(2)), yCoord)
        values(3) = MaxDouble(CDbl(values(3)), yCoord)
        values(4) = CDbl(values(4)) + xCoord
        values(5) = CDbl(values(5)) + yCoord
        values(6) = CLng(values(6)) + 1
        stats(groupName) = values
    Else
        stats.Add groupName, Array(xCoord, xCoord, yCoord, yCoord, xCoord, yCoord, 1)
    End If
End Sub

Private Sub PutRebarAnnotationForGroup(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal groupName As String, ByVal textValue As String, ByVal stats As Object, _
        ByVal section As CSectionModel, ByVal units As CUnitSystem, _
        ByVal sectionMinX As Double, ByVal sectionMaxX As Double, ByVal sectionMinY As Double, _
        ByVal sectionMaxY As Double, ByVal offset As Double, ByVal topRight As Double, _
        ByVal rightTop As Double, ByVal hasLShape As Boolean)
    If Not stats.Exists(groupName) Then Exit Sub

    Dim values As Variant
    values = stats(groupName)

    Dim minX As Double, maxX As Double, minY As Double, maxY As Double, avgX As Double, avgY As Double
    minX = CDbl(values(0))
    maxX = CDbl(values(1))
    minY = CDbl(values(2))
    maxY = CDbl(values(3))
    avgX = CDbl(values(4)) / CLng(values(6))
    avgY = CDbl(values(5)) / CLng(values(6))

    Dim centerX As Double, centerY As Double
    centerX = 0.5 * (sectionMinX + sectionMaxX)
    centerY = 0.5 * (sectionMinY + sectionMaxY)

    Dim startX As Double, startY As Double, endX As Double, endY As Double, normalX As Double, normalY As Double
    Dim axisDistance As Double
    If SectionRebarAnnotationAnchor(groupName, section, units, startX, startY, endX, endY, normalX, normalY, axisDistance) Then
        ' Anchor line is supplied by the geometry/rebar generator.
    ElseIf hasLShape And LShapeRebarAnchor(groupName, sectionMinX, sectionMaxX, sectionMinY, sectionMaxY, _
            topRight, rightTop, startX, startY, endX, endY, normalX, normalY) Then
        ' Fallback anchor is defined by semantic L-shape face name.
    ElseIf StrComp(groupName, "Circle", vbTextCompare) = 0 Then
        startX = sectionMinX
        startY = sectionMaxY
        endX = sectionMaxX
        endY = sectionMaxY
        normalY = 1#
    ElseIf (maxY - minY) > (maxX - minX) Then
        startX = avgX
        startY = minY
        endX = avgX
        endY = maxY
        If avgX < centerX Then normalX = -1# Else normalX = 1#
    Else
        startX = minX
        startY = avgY
        endX = maxX
        endY = avgY
        If avgY < centerY Then normalY = -1# Else normalY = 1#
    End If

    PutRebarAnnotation data, rowIndex, runID, "REBAR_" & groupName, startX, startY, _
        endX, endY, normalX, normalY, axisDistance, textValue, "Grouped rebar annotation."
End Sub

Private Function LShapeRebarAnchor(ByVal groupName As String, ByVal minX As Double, ByVal maxX As Double, _
        ByVal minY As Double, ByVal maxY As Double, ByVal topRight As Double, ByVal rightTop As Double, _
        ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double, _
        ByRef normalX As Double, ByRef normalY As Double) As Boolean
    Select Case LCase$(Trim$(groupName))
        Case "h1.as_1"
            startX = minX
            startY = rightTop
            endX = minX
            endY = maxY
            normalX = -1#
        Case "h1.as_2"
            startX = topRight
            startY = rightTop
            endX = topRight
            endY = maxY
            normalX = 1#
        Case "b1.as_1"
            startX = minX
            startY = maxY
            endX = topRight
            endY = maxY
            normalY = 1#
        Case "b1.as_2"
            startX = minX
            startY = rightTop
            endX = topRight
            endY = rightTop
            normalY = -1#
        Case "h2.as_1"
            startX = minX
            startY = minY
            endX = minX
            endY = rightTop
            normalX = -1#
        Case "h2.as_2"
            startX = maxX
            startY = minY
            endX = maxX
            endY = rightTop
            normalX = 1#
        Case "b2.as_1"
            startX = minX
            startY = rightTop
            endX = maxX
            endY = rightTop
            normalY = 1#
        Case "b2.as_2"
            startX = minX
            startY = minY
            endX = maxX
            endY = minY
            normalY = -1#
        Case Else
            Exit Function
    End Select
    LShapeRebarAnchor = True
End Function

Private Function SectionRebarAnnotationAnchor(ByVal groupName As String, ByVal section As CSectionModel, _
        ByVal units As CUnitSystem, _
        ByRef startX As Double, ByRef startY As Double, ByRef endX As Double, ByRef endY As Double, _
        ByRef normalX As Double, ByRef normalY As Double, ByRef axisDistance As Double) As Boolean
    Dim i As Long
    For i = 1 To section.RebarAnnotationCount
        If StrComp(section.RebarAnnotationGroupName(i), groupName, vbTextCompare) = 0 Then
            startX = OutputLength(section.RebarAnnotationStartX(i), units)
            startY = OutputLength(section.RebarAnnotationStartY(i), units)
            endX = OutputLength(section.RebarAnnotationEndX(i), units)
            endY = OutputLength(section.RebarAnnotationEndY(i), units)
            normalX = section.RebarAnnotationNormalX(i)
            normalY = section.RebarAnnotationNormalY(i)
            axisDistance = OutputLength(section.RebarAnnotationAxisDistance(i), units)
            SectionRebarAnnotationAnchor = True
            Exit Function
        End If
    Next i
End Function

Private Function RebarGroupName(ByVal commentText As String) As String
    Dim textValue As String
    textValue = Trim$(commentText)
    If InStr(1, textValue, "auto lshape ", vbTextCompare) = 1 Then
        Dim tail As String
        tail = Mid$(textValue, Len("auto lshape ") + 1)
        Dim rowPos As Long
        rowPos = InStr(1, tail, " row ", vbTextCompare)
        If rowPos > 1 Then RebarGroupName = Trim$(Left$(tail, rowPos - 1))
    ElseIf InStr(1, textValue, "auto circle", vbTextCompare) = 1 Then
        RebarGroupName = "Circle"
    End If
End Function

Private Sub PutRebarAnnotation(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal annotationID As String, ByVal startX As Double, ByVal startY As Double, _
        ByVal endX As Double, ByVal endY As Double, ByVal normalX As Double, ByVal normalY As Double, _
        ByVal offset As Double, ByVal textValue As String, ByVal comment As String)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = "REBAR_ANNOTATION"
    data(rowIndex, 3) = annotationID
    data(rowIndex, 4) = startX
    data(rowIndex, 5) = startY
    data(rowIndex, 6) = endX
    data(rowIndex, 7) = endY
    data(rowIndex, 8) = normalX
    data(rowIndex, 9) = normalY
    data(rowIndex, 10) = offset
    data(rowIndex, 11) = textValue
    data(rowIndex, 12) = vbNullString
    data(rowIndex, 13) = vbNullString
    data(rowIndex, 14) = comment
    rowIndex = rowIndex + 1
End Sub

Private Function HasGeneratedSectionAnnotations(ByVal section As CSectionModel) As Boolean
    Select Case LCase$(Trim$(section.SourceType))
        Case "circle", "lshape", "roundedrectangle", "generated"
            HasGeneratedSectionAnnotations = True
    End Select
End Function

Private Function IsCircleModel(ByVal section As CSectionModel) As Boolean
    IsCircleModel = (StrComp(Trim$(section.SourceType), "Circle", vbTextCompare) = 0)
End Function

Private Sub WriteContourAnnotations(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal section As CSectionModel, ByVal units As CUnitSystem, ByVal minX As Double, ByVal maxX As Double, _
        ByVal minY As Double, ByVal maxY As Double)
    If IsCircleModel(section) Then
        PutContourCircleAnnotation data, rowIndex, runID, 0.5 * (minX + maxX), 0.5 * (minY + maxY), _
            0.5 * MinDouble(maxX - minX, maxY - minY)
        Exit Sub
    End If

    Dim topRight As Double, rightTop As Double, hasLShape As Boolean
    DetectLShapeBreak section, units, minX, maxX, minY, maxY, topRight, rightTop, hasLShape
    If hasLShape Then
        PutContourLineAnnotation data, rowIndex, runID, minX, minY, maxX, minY
        PutContourLineAnnotation data, rowIndex, runID, maxX, minY, maxX, rightTop
        PutContourLineAnnotation data, rowIndex, runID, maxX, rightTop, topRight, rightTop
        PutContourLineAnnotation data, rowIndex, runID, topRight, rightTop, topRight, maxY
        PutContourLineAnnotation data, rowIndex, runID, topRight, maxY, minX, maxY
        PutContourLineAnnotation data, rowIndex, runID, minX, maxY, minX, minY
    Else
        PutContourLineAnnotation data, rowIndex, runID, minX, minY, maxX, minY
        PutContourLineAnnotation data, rowIndex, runID, maxX, minY, maxX, maxY
        PutContourLineAnnotation data, rowIndex, runID, maxX, maxY, minX, maxY
        PutContourLineAnnotation data, rowIndex, runID, minX, maxY, minX, minY
    End If
End Sub

Private Sub PutContourLineAnnotation(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal startX As Double, ByVal startY As Double, ByVal endX As Double, ByVal endY As Double)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = "CONTOUR_LINE"
    data(rowIndex, 3) = "CONTOUR_LINE_" & CStr(rowIndex)
    data(rowIndex, 4) = startX
    data(rowIndex, 5) = startY
    data(rowIndex, 6) = endX
    data(rowIndex, 7) = endY
    data(rowIndex, 14) = "Generated section contour segment."
    rowIndex = rowIndex + 1
End Sub

Private Sub PutContourCircleAnnotation(ByRef data() As Variant, ByRef rowIndex As Long, ByVal runID As String, _
        ByVal centerX As Double, ByVal centerY As Double, ByVal radius As Double)
    data(rowIndex, 1) = runID
    data(rowIndex, 2) = "CONTOUR_CIRCLE"
    data(rowIndex, 3) = "CONTOUR_CIRCLE"
    data(rowIndex, 4) = centerX
    data(rowIndex, 5) = centerY
    data(rowIndex, 6) = radius
    data(rowIndex, 7) = 0#
    data(rowIndex, 14) = "Generated circle contour."
    rowIndex = rowIndex + 1
End Sub

Private Sub DetectLShapeBreak(ByVal section As CSectionModel, ByVal units As CUnitSystem, _
        ByVal minX As Double, ByVal maxX As Double, ByVal minY As Double, ByVal maxY As Double, _
        ByRef topRight As Double, ByRef rightTop As Double, ByRef hasLShape As Boolean)
    topRight = minX
    rightTop = minY
    hasLShape = False

    Dim dx As Double, dy As Double
    dx = maxX - minX
    dy = maxY - minY
    If dx <= 0# Or dy <= 0# Then Exit Sub

    Dim tol As Double
    tol = MaxDouble(dx, dy) * 0.03

    Dim i As Long
    For i = 1 To section.ConcreteCount
        Dim cx As Double, cy As Double, w As Double, h As Double
        cx = OutputLength(section.ConcreteX(i), units)
        cy = OutputLength(section.ConcreteY(i), units)
        w = OutputLength(ConcreteWidth(section, i), units)
        h = OutputLength(ConcreteHeight(section, i), units)
        If cy + h / 2# >= maxY - tol Then topRight = MaxDouble(topRight, cx + w / 2#)
        If cx + w / 2# >= maxX - tol Then rightTop = MaxDouble(rightTop, cy + h / 2#)
    Next i

    hasLShape = (topRight > minX + tol And topRight < maxX - tol And _
        rightTop > minY + tol And rightTop < maxY - tol)
End Sub

Private Sub GetConcreteBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4460, "CNDMResultsWriter", "No concrete elements in section model."

    minX = section.ConcreteX(1) - ConcreteWidth(section, 1) / 2#
    maxX = section.ConcreteX(1) + ConcreteWidth(section, 1) / 2#
    minY = section.ConcreteY(1) - ConcreteHeight(section, 1) / 2#
    maxY = section.ConcreteY(1) + ConcreteHeight(section, 1) / 2#

    Dim i As Long
    For i = 1 To section.ConcreteCount
        minX = MinDouble(minX, section.ConcreteX(i) - ConcreteWidth(section, i) / 2#)
        maxX = MaxDouble(maxX, section.ConcreteX(i) + ConcreteWidth(section, i) / 2#)
        minY = MinDouble(minY, section.ConcreteY(i) - ConcreteHeight(section, i) / 2#)
        maxY = MaxDouble(maxY, section.ConcreteY(i) + ConcreteHeight(section, i) / 2#)
    Next i
End Sub

Private Sub GetSectionBounds(ByVal section As CSectionModel, ByRef minX As Double, ByRef maxX As Double, _
        ByRef minY As Double, ByRef maxY As Double)
    If section.ConcreteCount <= 0 Then Err.Raise vbObjectError + 4460, "CNDMResultsWriter", "В модели нет бетонных элементов."

    minX = section.ConcreteX(1) - ConcreteWidth(section, 1) / 2#
    maxX = section.ConcreteX(1) + ConcreteWidth(section, 1) / 2#
    minY = section.ConcreteY(1) - ConcreteHeight(section, 1) / 2#
    maxY = section.ConcreteY(1) + ConcreteHeight(section, 1) / 2#

    Dim i As Long
    For i = 1 To section.ConcreteCount
        minX = MinDouble(minX, section.ConcreteX(i) - ConcreteWidth(section, i) / 2#)
        maxX = MaxDouble(maxX, section.ConcreteX(i) + ConcreteWidth(section, i) / 2#)
        minY = MinDouble(minY, section.ConcreteY(i) - ConcreteHeight(section, i) / 2#)
        maxY = MaxDouble(maxY, section.ConcreteY(i) + ConcreteHeight(section, i) / 2#)
    Next i
    For i = 1 To section.RebarCount
        minX = MinDouble(minX, section.RebarX(i) - section.RebarDiameter(i) / 2#)
        maxX = MaxDouble(maxX, section.RebarX(i) + section.RebarDiameter(i) / 2#)
        minY = MinDouble(minY, section.RebarY(i) - section.RebarDiameter(i) / 2#)
        maxY = MaxDouble(maxY, section.RebarY(i) + section.RebarDiameter(i) / 2#)
    Next i
End Sub

Private Function ConcreteWidth(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteWidth = section.ConcreteWidth(index)
    If ConcreteWidth <= 0# Then ConcreteWidth = Sqr(section.ConcreteArea(index))
End Function

Private Function ConcreteHeight(ByVal section As CSectionModel, ByVal index As Long) As Double
    ConcreteHeight = section.ConcreteHeight(index)
    If ConcreteHeight <= 0# Then ConcreteHeight = Sqr(section.ConcreteArea(index))
End Function

Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

Private Function OutputArea(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputArea = value Else OutputArea = units.InternalAreaToOutput(value)
End Function

Private Function OutputFourthPowerLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputFourthPowerLength = value Else OutputFourthPowerLength = units.InternalFourthPowerLengthToOutput(value)
End Function

Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value Else OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

Private Function OutputAreaUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputAreaUnit = "mm2" Else OutputAreaUnit = units.OutputAreaUnit
End Function

Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

Private Function OutputSignN(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignN = "Tension" Else OutputSignN = units.OutputSignN
End Function

Private Function OutputSignMx(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMx = "+Y tension" Else OutputSignMx = units.OutputSignMx
End Function

Private Function OutputSignMy(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputSignMy = "+X tension" Else OutputSignMy = units.OutputSignMy
End Function

Private Function FormatNumberInvariant(ByVal value As Double) As String
    Dim textValue As String
    textValue = Replace$(Format$(value, "0.###"), ",", ".")
    Do While InStr(1, textValue, ".", vbTextCompare) > 0 And Right$(textValue, 1) = "0"
        textValue = Left$(textValue, Len(textValue) - 1)
    Loop
    If Right$(textValue, 1) = "." Then textValue = Left$(textValue, Len(textValue) - 1)
    FormatNumberInvariant = textValue
End Function

Private Function MaxDouble(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxDouble = a Else MaxDouble = b
End Function

Private Function MinDouble(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinDouble = a Else MinDouble = b
End Function
