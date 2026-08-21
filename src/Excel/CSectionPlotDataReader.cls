VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionPlotDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mResultType As String
Private mLoadCase As String
Private mResultUnit As String
Private mResultCaption As String
Private mLengthUnit As String
Private mAreaUnit As String
Private mStressUnit As String
Private mCurvatureUnit As String
Private mCount As Long
Private mElementID() As String
Private mMaterialType() As String
Private mX() As Double
Private mY() As Double
Private mArea() As Double
Private mDiameter() As Double
Private mWidth() As Double
Private mHeight() As Double
Private mResultValue() As Double
Private mPhysicalState() As String
Private mMinX As Double
Private mMaxX As Double
Private mMinY As Double
Private mMaxY As Double
Private mCentroidX As Double
Private mCentroidY As Double
Private mPrincipalAngle As Double
Private mEpsilon0 As Double
Private mKappaX As Double
Private mKappaY As Double
Private mLoadReferenceX As Double
Private mLoadReferenceY As Double
Private mAnnotationCount As Long
Private mAnnotationType() As String
Private mAnnotationStartX() As Double
Private mAnnotationStartY() As Double
Private mAnnotationEndX() As Double
Private mAnnotationEndY() As Double
Private mAnnotationNormalX() As Double
Private mAnnotationNormalY() As Double
Private mAnnotationOffset() As Double
Private mAnnotationText() As String

Public Sub LoadFromWorkbook(ByVal workbook As Object, ByVal settings As CSystemSettingsReader)
    If workbook Is Nothing Then Err.Raise vbObjectError + 4700, "CSectionPlotDataReader", "Книга Excel не передана."
    If settings Is Nothing Then Err.Raise vbObjectError + 4701, "CSectionPlotDataReader", "Настройки System не переданы."

    mResultType = NormalizeResultType(settings.GetString("Plot.ResultType", "Stress"))
    mLoadCase = ResolveLoadCase(workbook, settings.GetString("Plot.LoadCase", "Worst"))

    ReadSectionProperties workbook
    ReadGeometryAndResults workbook
    ReadAnnotations workbook
End Sub

Private Function ResolveLoadCase(ByVal workbook As Object, ByVal requested As String) As String
    Dim valueText As String
    valueText = Trim$(requested)
    If Len(valueText) = 0 Then valueText = "Worst"

    If StrComp(valueText, "Worst", vbTextCompare) <> 0 Then
        ResolveLoadCase = valueText
        Exit Function
    End If

    Dim data As Variant
    data = workbook.Names.Item("rngBatchSummary").RefersToRange.Value2
    If UBound(data, 1) >= 2 And UBound(data, 2) >= 2 Then
        ResolveLoadCase = Trim$(CStr(data(2, 2)))
    End If
    If Len(ResolveLoadCase) = 0 Then Err.Raise vbObjectError + 4702, "CSectionPlotDataReader", _
        "В rngBatchSummary не найдено определяющее сочетание для Plot.LoadCase = Worst."
End Function

Private Sub ReadSectionProperties(ByVal workbook As Object)
    Dim data As Variant
    data = workbook.Names.Item("rngNDMSectionProperties").RefersToRange.CurrentRegion.Value2
    If Not HasRows(data) Then Err.Raise vbObjectError + 4703, "CSectionPlotDataReader", _
        "На листе Results нет rngNDMSectionProperties. Сначала выполните расчет."

    Dim colLoadCase As Long: colLoadCase = ResultColumn(data, "LoadCase")
    Dim colParameter As Long: colParameter = ResultColumn(data, "Parameter")
    Dim colValue As Long: colValue = ResultColumn(data, "Value")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim parameter As String
        Dim loadCase As String
        parameter = Trim$(CStr(data(rowIndex, colParameter)))
        loadCase = Trim$(CStr(data(rowIndex, colLoadCase)))
        If Len(parameter) > 0 Then
            If StrComp(loadCase, "ALL", vbTextCompare) = 0 Or Len(loadCase) = 0 Then
                ReadCommonProperty parameter, data(rowIndex, colValue)
            ElseIf StrComp(loadCase, mLoadCase, vbTextCompare) = 0 Then
                ReadLoadCaseProperty parameter, data(rowIndex, colValue)
            End If
        End If
    Next rowIndex

    If Len(mLengthUnit) = 0 Then mLengthUnit = "mm"
    If Len(mAreaUnit) = 0 Then mAreaUnit = mLengthUnit & "2"
    If Len(mStressUnit) = 0 Then mStressUnit = "MPa"
    If Len(mCurvatureUnit) = 0 Then mCurvatureUnit = "1/mm"
    If StrComp(mResultType, "Strain", vbTextCompare) = 0 Then
        mResultUnit = "-"
        mResultCaption = "Strain"
    Else
        mResultUnit = mStressUnit
        mResultCaption = "Stress"
    End If
End Sub

Private Sub ReadCommonProperty(ByVal parameter As String, ByVal value As Variant)
    Select Case LCase$(parameter)
        Case "bounds.minx": mMinX = CDbl(value)
        Case "bounds.maxx": mMaxX = CDbl(value)
        Case "bounds.miny": mMinY = CDbl(value)
        Case "bounds.maxy": mMaxY = CDbl(value)
        Case "centroidx": mCentroidX = CDbl(value)
        Case "centroidy": mCentroidY = CDbl(value)
        Case "principalangle": mPrincipalAngle = CDbl(value)
        Case "output.lengthunit": mLengthUnit = CStr(value)
        Case "output.areaunit": mAreaUnit = CStr(value)
        Case "output.stressunit": mStressUnit = CStr(value)
        Case "output.curvatureunit": mCurvatureUnit = CStr(value)
    End Select
End Sub

Private Sub ReadLoadCaseProperty(ByVal parameter As String, ByVal value As Variant)
    Select Case LCase$(parameter)
        Case "epsilon0": mEpsilon0 = CDbl(value)
        Case "kappax": mKappaX = CDbl(value)
        Case "kappay": mKappaY = CDbl(value)
        Case "loadreferencex": mLoadReferenceX = CDbl(value)
        Case "loadreferencey": mLoadReferenceY = CDbl(value)
    End Select
End Sub

Private Sub ReadGeometryAndResults(ByVal workbook As Object)
    Dim geom As Variant
    geom = workbook.Names.Item("rngNDMSectionGeometry").RefersToRange.CurrentRegion.Value2
    If Not HasRows(geom) Then Err.Raise vbObjectError + 4704, "CSectionPlotDataReader", _
        "На листе Results нет расчетной геометрии."

    Dim results As Variant
    results = workbook.Names.Item("rngNDMElementResults").RefersToRange.CurrentRegion.Value2
    If Not HasRows(results) Then Err.Raise vbObjectError + 4705, "CSectionPlotDataReader", _
        "На листе Results нет результатов элементов."

    Dim resultMap As Object
    Set resultMap = CreateObject("Scripting.Dictionary")
    resultMap.CompareMode = vbTextCompare

    Dim colResLC As Long: colResLC = ResultColumn(results, "LoadCase")
    Dim colResID As Long: colResID = ResultColumn(results, "ElementID")
    Dim colStrain As Long: colStrain = ResultColumn(results, "Strain")
    Dim colStress As Long: colStress = ResultColumn(results, "Stress")
    Dim colState As Long: colState = ResultColumn(results, "PhysicalState")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(results, 1)
        If StrComp(CStr(results(rowIndex, colResLC)), mLoadCase, vbTextCompare) = 0 Then
            resultMap(CStr(results(rowIndex, colResID))) = Array(CDbl(results(rowIndex, colStrain)), _
                CDbl(results(rowIndex, colStress)), CStr(results(rowIndex, colState)))
        End If
    Next rowIndex
    If resultMap.Count = 0 Then Err.Raise vbObjectError + 4706, "CSectionPlotDataReader", _
        "В Results нет данных для Plot.LoadCase: " & mLoadCase

    Dim colID As Long: colID = ResultColumn(geom, "ElementID")
    Dim colType As Long: colType = ResultColumn(geom, "MaterialType")
    Dim colX As Long: colX = ResultColumn(geom, "X")
    Dim colY As Long: colY = ResultColumn(geom, "Y")
    Dim colArea As Long: colArea = ResultColumn(geom, "Area")
    Dim colDiameter As Long: colDiameter = ResultColumn(geom, "Diameter")
    Dim colWidth As Long: colWidth = ResultColumn(geom, "Width")
    Dim colHeight As Long: colHeight = ResultColumn(geom, "Height")

    For rowIndex = 2 To UBound(geom, 1)
        Dim elementID As String
        elementID = Trim$(CStr(geom(rowIndex, colID)))
        If Len(elementID) > 0 And resultMap.Exists(elementID) Then
            AppendElement elementID, CStr(geom(rowIndex, colType)), CDbl(geom(rowIndex, colX)), _
                CDbl(geom(rowIndex, colY)), CDbl(geom(rowIndex, colArea)), _
                SafeDouble(geom(rowIndex, colDiameter)), SafeDouble(geom(rowIndex, colWidth)), _
                SafeDouble(geom(rowIndex, colHeight)), resultMap(elementID)
        End If
    Next rowIndex
End Sub

Private Sub AppendElement(ByVal elementID As String, ByVal materialType As String, ByVal xCoord As Double, _
        ByVal yCoord As Double, ByVal area As Double, ByVal diameter As Double, _
        ByVal width As Double, ByVal height As Double, ByVal resultData As Variant)
    mCount = mCount + 1
    ReDim Preserve mElementID(1 To mCount)
    ReDim Preserve mMaterialType(1 To mCount)
    ReDim Preserve mX(1 To mCount)
    ReDim Preserve mY(1 To mCount)
    ReDim Preserve mArea(1 To mCount)
    ReDim Preserve mDiameter(1 To mCount)
    ReDim Preserve mWidth(1 To mCount)
    ReDim Preserve mHeight(1 To mCount)
    ReDim Preserve mResultValue(1 To mCount)
    ReDim Preserve mPhysicalState(1 To mCount)

    mElementID(mCount) = elementID
    mMaterialType(mCount) = materialType
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mArea(mCount) = area
    mDiameter(mCount) = diameter
    mWidth(mCount) = width
    mHeight(mCount) = height
    If StrComp(mResultType, "Strain", vbTextCompare) = 0 Then
        mResultValue(mCount) = CDbl(resultData(0))
    Else
        mResultValue(mCount) = CDbl(resultData(1))
    End If
    mPhysicalState(mCount) = CStr(resultData(2))
End Sub

Private Sub ReadAnnotations(ByVal workbook As Object)
    On Error GoTo NoAnnotations
    Dim data As Variant
    data = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange.CurrentRegion.Value2
    If Not HasRows(data) Then Exit Sub

    Dim colType As Long: colType = ResultColumn(data, "AnnotationType")
    Dim colStartX As Long: colStartX = ResultColumn(data, "StartX")
    Dim colStartY As Long: colStartY = ResultColumn(data, "StartY")
    Dim colEndX As Long: colEndX = ResultColumn(data, "EndX")
    Dim colEndY As Long: colEndY = ResultColumn(data, "EndY")
    Dim colNormalX As Long: colNormalX = ResultColumn(data, "OutsideNormalX")
    Dim colNormalY As Long: colNormalY = ResultColumn(data, "OutsideNormalY")
    Dim colOffset As Long: colOffset = ResultColumn(data, "Offset")
    Dim colText As Long: colText = ResultColumn(data, "Text")

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        If Len(Trim$(CStr(data(rowIndex, colType)))) > 0 Then
            mAnnotationCount = mAnnotationCount + 1
            ReDim Preserve mAnnotationType(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationStartY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationEndY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalX(1 To mAnnotationCount)
            ReDim Preserve mAnnotationNormalY(1 To mAnnotationCount)
            ReDim Preserve mAnnotationOffset(1 To mAnnotationCount)
            ReDim Preserve mAnnotationText(1 To mAnnotationCount)
            mAnnotationType(mAnnotationCount) = CStr(data(rowIndex, colType))
            mAnnotationStartX(mAnnotationCount) = SafeDouble(data(rowIndex, colStartX))
            mAnnotationStartY(mAnnotationCount) = SafeDouble(data(rowIndex, colStartY))
            mAnnotationEndX(mAnnotationCount) = SafeDouble(data(rowIndex, colEndX))
            mAnnotationEndY(mAnnotationCount) = SafeDouble(data(rowIndex, colEndY))
            mAnnotationNormalX(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalX))
            mAnnotationNormalY(mAnnotationCount) = SafeDouble(data(rowIndex, colNormalY))
            mAnnotationOffset(mAnnotationCount) = SafeDouble(data(rowIndex, colOffset))
            mAnnotationText(mAnnotationCount) = CStr(data(rowIndex, colText))
        End If
    Next rowIndex
NoAnnotations:
End Sub

Private Function NormalizeResultType(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case "stress": NormalizeResultType = "Stress"
        Case "strain": NormalizeResultType = "Strain"
        Case Else: Err.Raise vbObjectError + 4707, "CSectionPlotDataReader", "Plot.ResultType должен быть Stress или Strain."
    End Select
End Function

Private Function SafeDouble(ByVal value As Variant) As Double
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    If IsNumeric(value) Then SafeDouble = CDbl(value)
End Function

Private Function HasRows(ByVal data As Variant) As Boolean
    On Error GoTo Failed
    HasRows = (UBound(data, 1) >= 2 And UBound(data, 2) >= 1)
    Exit Function
Failed:
End Function

Private Function ResultColumn(ByRef data As Variant, ByVal headerName As String) As Long
    Dim colIndex As Long
    For colIndex = 1 To UBound(data, 2)
        If StrComp(ResultHeaderBase(CStr(data(1, colIndex))), headerName, vbTextCompare) = 0 Then
            ResultColumn = colIndex
            Exit Function
        End If
    Next colIndex
    Err.Raise vbObjectError + 4708, "CSectionPlotDataReader", "В Results не найден столбец: " & headerName
End Function

Private Function ResultHeaderBase(ByVal headerText As String) As String
    Dim commaPos As Long
    commaPos = InStr(1, headerText, ",", vbTextCompare)
    If commaPos > 0 Then
        ResultHeaderBase = Trim$(Left$(headerText, commaPos - 1))
    Else
        ResultHeaderBase = Trim$(headerText)
    End If
End Function

Public Property Get Count() As Long: Count = mCount: End Property
Public Property Get ElementID(ByVal index As Long) As String: ElementID = mElementID(index): End Property
Public Property Get MaterialType(ByVal index As Long) As String: MaterialType = mMaterialType(index): End Property
Public Property Get X(ByVal index As Long) As Double: X = mX(index): End Property
Public Property Get Y(ByVal index As Long) As Double: Y = mY(index): End Property
Public Property Get Area(ByVal index As Long) As Double: Area = mArea(index): End Property
Public Property Get Diameter(ByVal index As Long) As Double: Diameter = mDiameter(index): End Property
Public Property Get Width(ByVal index As Long) As Double: Width = mWidth(index): End Property
Public Property Get Height(ByVal index As Long) As Double: Height = mHeight(index): End Property
Public Property Get ResultValue(ByVal index As Long) As Double: ResultValue = mResultValue(index): End Property
Public Property Get PhysicalState(ByVal index As Long) As String: PhysicalState = mPhysicalState(index): End Property
Public Property Get LoadCase() As String: LoadCase = mLoadCase: End Property
Public Property Get ResultType() As String: ResultType = mResultType: End Property
Public Property Get ResultUnit() As String: ResultUnit = mResultUnit: End Property
Public Property Get ResultCaption() As String: ResultCaption = mResultCaption: End Property
Public Property Get LengthUnit() As String: LengthUnit = mLengthUnit: End Property
Public Property Get AreaUnit() As String: AreaUnit = mAreaUnit: End Property
Public Property Get MinX() As Double: MinX = mMinX: End Property
Public Property Get MaxX() As Double: MaxX = mMaxX: End Property
Public Property Get MinY() As Double: MinY = mMinY: End Property
Public Property Get MaxY() As Double: MaxY = mMaxY: End Property
Public Property Get CentroidX() As Double: CentroidX = mCentroidX: End Property
Public Property Get CentroidY() As Double: CentroidY = mCentroidY: End Property
Public Property Get PrincipalAngle() As Double: PrincipalAngle = mPrincipalAngle: End Property
Public Property Get Epsilon0() As Double: Epsilon0 = mEpsilon0: End Property
Public Property Get KappaX() As Double: KappaX = mKappaX: End Property
Public Property Get KappaY() As Double: KappaY = mKappaY: End Property
Public Property Get LoadReferenceX() As Double: LoadReferenceX = mLoadReferenceX: End Property
Public Property Get LoadReferenceY() As Double: LoadReferenceY = mLoadReferenceY: End Property
Public Property Get AnnotationCount() As Long: AnnotationCount = mAnnotationCount: End Property
Public Property Get AnnotationType(ByVal index As Long) As String: AnnotationType = mAnnotationType(index): End Property
Public Property Get AnnotationStartX(ByVal index As Long) As Double: AnnotationStartX = mAnnotationStartX(index): End Property
Public Property Get AnnotationStartY(ByVal index As Long) As Double: AnnotationStartY = mAnnotationStartY(index): End Property
Public Property Get AnnotationEndX(ByVal index As Long) As Double: AnnotationEndX = mAnnotationEndX(index): End Property
Public Property Get AnnotationEndY(ByVal index As Long) As Double: AnnotationEndY = mAnnotationEndY(index): End Property
Public Property Get AnnotationNormalX(ByVal index As Long) As Double: AnnotationNormalX = mAnnotationNormalX(index): End Property
Public Property Get AnnotationNormalY(ByVal index As Long) As Double: AnnotationNormalY = mAnnotationNormalY(index): End Property
Public Property Get AnnotationOffset(ByVal index As Long) As Double: AnnotationOffset = mAnnotationOffset(index): End Property
Public Property Get AnnotationText(ByVal index As Long) As String: AnnotationText = mAnnotationText(index): End Property
