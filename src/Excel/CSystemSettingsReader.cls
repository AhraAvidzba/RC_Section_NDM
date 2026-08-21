VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSystemSettingsReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKeys() As String
Private mValues() As String
Private mDefaults() As String
Private mSources() As String
Private mCount As Long
Private mConcreteStrains() As Double
Private mConcreteStresses() As Double
Private mConcretePointCount As Long
Private mSteelStrains() As Double
Private mSteelStresses() As Double
Private mSteelPointCount As Long

Public Sub LoadFromWorkbook(ByVal workbook As Object)
    Clear
    AppendSettingsRange workbook.Names.Item("rngUnitSettings").RefersToRange
    AppendSettingsRange workbook.Names.Item("rngSignConventionSettings").RefersToRange
    AppendSettingsRange workbook.Names.Item("rngSystemSettings").RefersToRange
    AppendSettingsRange workbook.Names.Item("rngPlotAnnotationSettings").RefersToRange

    Dim geometryType As String
    geometryType = GetString("Geometry.Type", vbNullString)
    Select Case LCase$(Trim$(geometryType))
        Case "circle"
            AppendSettingsRange workbook.Names.Item("rngCircleGeometry").RefersToRange
        Case "roundedrectangle"
            AppendSettingsRange workbook.Names.Item("rngRoundedRectangleGeometry").RefersToRange
        Case "lshape"
            AppendSettingsRange workbook.Names.Item("rngLShapeGeometry").RefersToRange
    End Select

    LoadDiagramPoints workbook.Names.Item("rngConcreteDiagramPoints").RefersToRange, _
        mConcreteStrains, mConcreteStresses, mConcretePointCount, "rngConcreteDiagramPoints"
    LoadDiagramPoints workbook.Names.Item("rngSteelDiagramPoints").RefersToRange, _
        mSteelStrains, mSteelStresses, mSteelPointCount, "rngSteelDiagramPoints"
End Sub

Public Sub LoadFromRange(ByVal settingsRange As Object)
    Clear
    AppendSettingsRange settingsRange
End Sub

Private Sub AppendSettingsRange(ByVal settingsRange As Object)
    Dim data As Variant
    data = settingsRange.Value2
    Dim columnCount As Long
    columnCount = UBound(data, 2)

    If IsLShapeFaceSettingsRange(data) Then
        AppendLShapeFaceSettings data
        Exit Sub
    End If
    If IsUnitSettingsRange(data) Then
        AppendUnitSettings data
        Exit Sub
    End If
    If IsSignConventionSettingsRange(data) Then
        AppendSignConventionSettings data
        Exit Sub
    End If
    If IsPlotAnnotationSettingsRange(data) Then
        AppendPlotAnnotationSettings data
        Exit Sub
    End If

    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim keyValue As String
        keyValue = Trim$(CStr(data(rowIndex, 1)))
        If Len(keyValue) > 0 And Left$(keyValue, 1) <> "[" Then
            Dim sourceValue As String
            If columnCount >= 6 Then
                sourceValue = Trim$(CStr(data(rowIndex, 6)))
            ElseIf columnCount >= 4 Then
                sourceValue = Trim$(CStr(data(rowIndex, 4)))
            Else
                sourceValue = vbNullString
            End If
            AppendSetting keyValue, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 3))), sourceValue
        End If
    Next rowIndex
End Sub

Private Function IsUnitSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 4 Then Exit Function
    IsUnitSettingsRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Quantity", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "INPUT", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "INTERNAL", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 4))), "OUTPUT", vbTextCompare) = 0
End Function

Private Function IsSignConventionSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 2 Or UBound(data, 2) < 3 Then Exit Function
    IsSignConventionSettingsRange = _
        StrComp(Trim$(CStr(data(1, 1))), "Quantity", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 2))), "USER", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "INTERNAL", vbTextCompare) = 0
End Function
Private Function IsPlotAnnotationSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) < 3 Or UBound(data, 2) < 5 Then Exit Function
    IsPlotAnnotationSettingsRange = _
        StrComp(Trim$(CStr(data(2, 1))), "Enabled", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(3, 1))), "Placement", vbTextCompare) = 0
End Function
Private Sub AppendPlotAnnotationSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim settingName As String
        settingName = Trim$(CStr(data(rowIndex, 1)))
        If Len(settingName) > 0 And Left$(settingName, 1) <> "[" Then
            AppendPlotAnnotationValue "Plot.RebarLabels." & settingName, Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 5)))
            AppendPlotAnnotationValue "Plot.Dimensions." & settingName, Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 5)))
        End If
    Next rowIndex
End Sub

Private Sub AppendPlotAnnotationValue(ByVal keyValue As String, ByVal settingValue As String, ByVal commentValue As String)
    If Len(settingValue) = 0 Or settingValue = "-" Then Exit Sub
    AppendSetting keyValue, settingValue, settingValue, commentValue
End Sub

Private Sub AppendUnitSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = UnitQuantityKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(quantity) > 0 Then
            AppendSetting "Units." & quantity & ".Input", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), "Пользовательская единица ввода"
            AppendSetting "Units." & quantity & ".Internal", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), "Фиксированная внутренняя единица"
            AppendSetting "Units." & quantity & ".Output", Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), "Пользовательская единица вывода"
        End If
    Next rowIndex
End Sub

Private Sub AppendSignConventionSettings(ByRef data As Variant)
    Dim rowIndex As Long
    For rowIndex = 2 To UBound(data, 1)
        Dim quantity As String
        quantity = SignQuantityKey(Trim$(CStr(data(rowIndex, 1))))
        If Len(quantity) > 0 Then
            AppendSetting "Sign." & quantity & ".User", Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), "Пользовательское правило знаков"
            AppendSetting "Sign." & quantity & ".Internal", Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), "Фиксированное внутреннее правило знаков"
        End If
    Next rowIndex
End Sub

Private Function UnitQuantityKey(ByVal caption As String) As String
    Select Case LCase$(Trim$(caption))
        Case "length": UnitQuantityKey = "Length"
        Case "area": UnitQuantityKey = "Area"
        Case "force": UnitQuantityKey = "Force"
        Case "moment": UnitQuantityKey = "Moment"
        Case "stress": UnitQuantityKey = "Stress"
        Case "curvature": UnitQuantityKey = "Curvature"
    End Select
End Function

Private Function SignQuantityKey(ByVal caption As String) As String
    Select Case LCase$(Trim$(caption))
        Case "+n": SignQuantityKey = "N"
        Case "+mx": SignQuantityKey = "Mx"
        Case "+my": SignQuantityKey = "My"
    End Select
End Function

Private Function IsLShapeFaceSettingsRange(ByRef data As Variant) As Boolean
    If UBound(data, 1) >= 3 And UBound(data, 2) >= 4 Then
        If StrComp(Trim$(CStr(data(2, 1))), "H1", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(2, 2))), "B1", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(2, 3))), "H2", vbTextCompare) = 0 And _
                StrComp(Trim$(CStr(data(2, 4))), "B2", vbTextCompare) = 0 Then
            IsLShapeFaceSettingsRange = True
            Exit Function
        End If
    End If

    If UBound(data, 1) < 12 Or UBound(data, 2) < 7 Then Exit Function
    IsLShapeFaceSettingsRange = _
        StrComp(Trim$(CStr(data(1, 2))), "H1", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 3))), "B1", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 4))), "H2", vbTextCompare) = 0 And _
        StrComp(Trim$(CStr(data(1, 5))), "B2", vbTextCompare) = 0
End Function

Private Sub AppendLShapeFaceSettings(ByRef data As Variant)
    If UBound(data, 1) >= 25 And UBound(data, 2) >= 9 And _
            StrComp(Trim$(CStr(data(2, 1))), "H1", vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(data(2, 2))), "B1", vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(data(2, 3))), "H2", vbTextCompare) = 0 And _
            StrComp(Trim$(CStr(data(2, 4))), "B2", vbTextCompare) = 0 Then
        AppendLShapeSectionedSettings data
        Exit Sub
    End If

    Dim faceColumns(1 To 4) As Long
    Dim faceNames(1 To 4) As String
    faceColumns(1) = 2: faceNames(1) = "H1"
    faceColumns(2) = 3: faceNames(2) = "B1"
    faceColumns(3) = 4: faceNames(3) = "H2"
    faceColumns(4) = 5: faceNames(4) = "B2"

    Dim i As Long
    For i = 1 To 4
        Dim col As Long
        Dim faceName As String
        col = faceColumns(i)
        faceName = faceNames(i)
        AppendSetting "LShape." & faceName, Trim$(CStr(data(2, col))), Trim$(CStr(data(2, col))), Trim$(CStr(data(2, 7)))
        AppendSetting "LShape." & faceName & ".as_1", Trim$(CStr(data(3, col))), Trim$(CStr(data(3, col))), Trim$(CStr(data(3, 7)))
        AppendSetting "LShape." & faceName & ".as_2", Trim$(CStr(data(4, col))), Trim$(CStr(data(4, col))), Trim$(CStr(data(4, 7)))
        AppendSetting "LShape." & faceName & ".d_1", Trim$(CStr(data(5, col))), Trim$(CStr(data(5, col))), Trim$(CStr(data(5, 7)))
        AppendSetting "LShape." & faceName & ".d_2", Trim$(CStr(data(6, col))), Trim$(CStr(data(6, col))), Trim$(CStr(data(6, 7)))
        AppendSetting "LShape." & faceName & ".n_1", Trim$(CStr(data(7, col))), Trim$(CStr(data(7, col))), Trim$(CStr(data(7, 7)))
        AppendSetting "LShape." & faceName & ".n_2", Trim$(CStr(data(8, col))), Trim$(CStr(data(8, col))), Trim$(CStr(data(8, 7)))
        AppendSetting "LShape." & faceName & ".StartOffset1", Trim$(CStr(data(9, col))), Trim$(CStr(data(9, col))), Trim$(CStr(data(9, 7)))
        AppendSetting "LShape." & faceName & ".EndOffset1", Trim$(CStr(data(10, col))), Trim$(CStr(data(10, col))), Trim$(CStr(data(10, 7)))
        AppendSetting "LShape." & faceName & ".StartOffset2", Trim$(CStr(data(11, col))), Trim$(CStr(data(11, col))), Trim$(CStr(data(11, 7)))
        AppendSetting "LShape." & faceName & ".EndOffset2", Trim$(CStr(data(12, col))), Trim$(CStr(data(12, col))), Trim$(CStr(data(12, 7)))
        If UBound(data, 1) >= 19 Then
            AppendSetting "LShape." & faceName & ".d_2row_1", Trim$(CStr(data(14, col))), Trim$(CStr(data(14, col))), Trim$(CStr(data(14, 7)))
            AppendSetting "LShape." & faceName & ".d_2row_2", Trim$(CStr(data(15, col))), Trim$(CStr(data(15, col))), Trim$(CStr(data(15, 7)))
            AppendSetting "LShape." & faceName & ".d_3row_1", Trim$(CStr(data(16, col))), Trim$(CStr(data(16, col))), Trim$(CStr(data(16, 7)))
            AppendSetting "LShape." & faceName & ".d_3row_2", Trim$(CStr(data(17, col))), Trim$(CStr(data(17, col))), Trim$(CStr(data(17, 7)))
            AppendSetting "LShape." & faceName & ".loc_2row", Trim$(CStr(data(18, col))), Trim$(CStr(data(18, col))), Trim$(CStr(data(18, 7)))
            AppendSetting "LShape." & faceName & ".loc_3row", Trim$(CStr(data(19, col))), Trim$(CStr(data(19, col))), Trim$(CStr(data(19, 7)))
            If UBound(data, 1) >= 21 Then
                AppendSetting "LShape." & faceName & ".bind_2row", Trim$(CStr(data(20, col))), Trim$(CStr(data(20, col))), Trim$(CStr(data(20, 7)))
                AppendSetting "LShape." & faceName & ".bind_3row", Trim$(CStr(data(21, col))), Trim$(CStr(data(21, col))), Trim$(CStr(data(21, 7)))
            End If
        End If
    Next i
End Sub

Private Sub AppendLShapeSectionedSettings(ByRef data As Variant)
    AppendSetting "LShape.H1", Trim$(CStr(data(3, 1))), Trim$(CStr(data(3, 1))), "LShape geometry"
    AppendSetting "LShape.B1", Trim$(CStr(data(3, 2))), Trim$(CStr(data(3, 2))), "LShape geometry"
    AppendSetting "LShape.H2", Trim$(CStr(data(3, 3))), Trim$(CStr(data(3, 3))), "LShape geometry"
    AppendSetting "LShape.B2", Trim$(CStr(data(3, 4))), Trim$(CStr(data(3, 4))), "LShape geometry"

    Dim faceNames(1 To 8) As String
    Dim sideIndexes(1 To 8) As Long
    FillLShapeSectionedFaceOrder faceNames, sideIndexes

    Dim i As Long
    For i = 1 To 8
        Dim rowIndex As Long
        rowIndex = 6 + i
        AppendSetting "LShape." & faceNames(i) & ".as_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 2))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "LShape." & faceNames(i) & ".d_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 3))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "LShape." & faceNames(i) & ".n_" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 4))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "LShape." & faceNames(i) & ".StartOffset" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 5))), Trim$(CStr(data(rowIndex, 8)))
        AppendSetting "LShape." & faceNames(i) & ".EndOffset" & CStr(sideIndexes(i)), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 6))), Trim$(CStr(data(rowIndex, 8)))

        Dim extraRow As Long
        extraRow = 17 + i
        AppendSetting "LShape." & faceNames(i) & ".d_2row_" & CStr(sideIndexes(i)), Trim$(CStr(data(extraRow, 2))), Trim$(CStr(data(extraRow, 2))), Trim$(CStr(data(extraRow, 9)))
        AppendSetting "LShape." & faceNames(i) & ".d_3row_" & CStr(sideIndexes(i)), Trim$(CStr(data(extraRow, 5))), Trim$(CStr(data(extraRow, 5))), Trim$(CStr(data(extraRow, 9)))
        If sideIndexes(i) = 1 Then
            AppendSetting "LShape." & faceNames(i) & ".loc_2row", Trim$(CStr(data(extraRow, 3))), Trim$(CStr(data(extraRow, 3))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "LShape." & faceNames(i) & ".bind_2row", Trim$(CStr(data(extraRow, 4))), Trim$(CStr(data(extraRow, 4))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "LShape." & faceNames(i) & ".loc_3row", Trim$(CStr(data(extraRow, 6))), Trim$(CStr(data(extraRow, 6))), Trim$(CStr(data(extraRow, 9)))
            AppendSetting "LShape." & faceNames(i) & ".bind_3row", Trim$(CStr(data(extraRow, 7))), Trim$(CStr(data(extraRow, 7))), Trim$(CStr(data(extraRow, 9)))
        End If
    Next i
End Sub

Private Sub FillLShapeSectionedFaceOrder(ByRef faceNames() As String, ByRef sideIndexes() As Long)
    faceNames(1) = "H1": sideIndexes(1) = 1
    faceNames(2) = "H1": sideIndexes(2) = 2
    faceNames(3) = "B1": sideIndexes(3) = 1
    faceNames(4) = "B1": sideIndexes(4) = 2
    faceNames(5) = "H2": sideIndexes(5) = 1
    faceNames(6) = "H2": sideIndexes(6) = 2
    faceNames(7) = "B2": sideIndexes(7) = 1
    faceNames(8) = "B2": sideIndexes(8) = 2
End Sub

Private Sub AppendSetting(ByVal keyValue As String, ByVal settingValue As String, _
        ByVal defaultValue As String, ByVal sourceValue As String)
    mCount = mCount + 1
    ReDim Preserve mKeys(1 To mCount)
    ReDim Preserve mValues(1 To mCount)
    ReDim Preserve mDefaults(1 To mCount)
    ReDim Preserve mSources(1 To mCount)
    mKeys(mCount) = keyValue
    mValues(mCount) = settingValue
    mDefaults(mCount) = defaultValue
    mSources(mCount) = sourceValue
End Sub

Public Sub Clear()
    Erase mKeys
    Erase mValues
    Erase mDefaults
    Erase mSources
    Erase mConcreteStrains
    Erase mConcreteStresses
    Erase mSteelStrains
    Erase mSteelStresses
    mCount = 0
    mConcretePointCount = 0
    mSteelPointCount = 0
End Sub

Public Function GetString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetString = defaultValue
    ElseIf Len(mValues(index)) = 0 Or UCase$(mValues(index)) = "TODO" Then
        GetString = defaultValue
    Else
        GetString = mValues(index)
    End If
End Function

Public Function GetRawString(ByVal key As String, Optional ByVal defaultValue As String = vbNullString) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetRawString = defaultValue
    Else
        GetRawString = mValues(index)
    End If
End Function

Public Function GetDouble(ByVal key As String, Optional ByVal defaultValue As Double = 0#) As Double
    Dim textValue As String
    textValue = GetString(key, vbNullString)
    If Len(textValue) = 0 Then
        GetDouble = defaultValue
    Else
        GetDouble = CDbl(Replace$(textValue, ".", ApplicationDecimalSeparator()))
    End If
End Function

Public Function GetLong(ByVal key As String, Optional ByVal defaultValue As Long = 0) As Long
    GetLong = CLng(GetDouble(key, CDbl(defaultValue)))
End Function

Public Function GetBoolean(ByVal key As String, Optional ByVal defaultValue As Boolean = False) As Boolean
    Dim textValue As String
    textValue = UCase$(GetString(key, IIf(defaultValue, "YES", "NO")))
    GetBoolean = (textValue = "YES" Or textValue = "TRUE" Or textValue = "1" Or textValue = "ДА")
End Function

Public Function GetSource(ByVal key As String) As String
    Dim index As Long
    index = FindKey(key)
    If index = 0 Then
        GetSource = vbNullString
    Else
        GetSource = mSources(index)
    End If
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = (FindKey(key) > 0)
End Function

Public Function KeyCount() As Long
    KeyCount = mCount
End Function

Public Function KeyAt(ByVal index As Long) As String
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 4300, "CSystemSettingsReader", "Setting index is out of range."
    KeyAt = mKeys(index)
End Function

Public Function DuplicateCount() As Long
    Dim duplicates As Long
    Dim i As Long
    Dim j As Long
    For i = 1 To mCount
        For j = i + 1 To mCount
            If StrComp(mKeys(i), mKeys(j), vbTextCompare) = 0 Then
                duplicates = duplicates + 1
                Exit For
            End If
        Next j
    Next i
    DuplicateCount = duplicates
End Function

Public Function ConcreteDiagramPointCount() As Long
    ConcreteDiagramPointCount = mConcretePointCount
End Function

Public Function ConcreteDiagramStrain(ByVal index As Long) As Double
    If index < 1 Or index > mConcretePointCount Then Err.Raise vbObjectError + 4310, "CSystemSettingsReader", "Concrete diagram point index is out of range."
    ConcreteDiagramStrain = mConcreteStrains(index)
End Function

Public Function ConcreteDiagramStress(ByVal index As Long) As Double
    If index < 1 Or index > mConcretePointCount Then Err.Raise vbObjectError + 4311, "CSystemSettingsReader", "Concrete diagram point index is out of range."
    ConcreteDiagramStress = mConcreteStresses(index)
End Function

Public Function SteelDiagramPointCount() As Long
    SteelDiagramPointCount = mSteelPointCount
End Function

Public Function SteelDiagramStrain(ByVal index As Long) As Double
    If index < 1 Or index > mSteelPointCount Then Err.Raise vbObjectError + 4312, "CSystemSettingsReader", "Steel diagram point index is out of range."
    SteelDiagramStrain = mSteelStrains(index)
End Function

Public Function SteelDiagramStress(ByVal index As Long) As Double
    If index < 1 Or index > mSteelPointCount Then Err.Raise vbObjectError + 4313, "CSystemSettingsReader", "Steel diagram point index is out of range."
    SteelDiagramStress = mSteelStresses(index)
End Function

Private Function FindKey(ByVal key As String) As Long
    Dim i As Long
    For i = 1 To mCount
        If StrComp(mKeys(i), key, vbTextCompare) = 0 Then
            FindKey = i
            Exit Function
        End If
    Next i
    FindKey = 0
End Function

Private Sub LoadDiagramPoints(ByVal pointRange As Object, ByRef strains() As Double, ByRef stresses() As Double, _
        ByRef pointCount As Long, ByVal rangeName As String)
    pointCount = 0
    Erase strains
    Erase stresses

    Dim data As Variant
    data = pointRange.Value2

    Dim rowIndex As Long
    For rowIndex = 1 To UBound(data, 1)
        Dim pointText As String
        Dim strainText As String
        Dim stressText As String
        pointText = Trim$(CStr(data(rowIndex, 1)))
        strainText = Trim$(CStr(data(rowIndex, 2)))
        stressText = Trim$(CStr(data(rowIndex, 3)))

        If Len(pointText) > 0 Or Len(strainText) > 0 Or Len(stressText) > 0 Then
            If Len(pointText) = 0 Or Len(strainText) = 0 Or Len(stressText) = 0 Then
                Err.Raise vbObjectError + 4314, "CSystemSettingsReader", rangeName & " has partially filled point row."
            End If
            pointCount = pointCount + 1
            ReDim Preserve strains(1 To pointCount)
            ReDim Preserve stresses(1 To pointCount)
            strains(pointCount) = CDbl(Replace$(strainText, ".", ApplicationDecimalSeparator()))
            stresses(pointCount) = CDbl(Replace$(stressText, ".", ApplicationDecimalSeparator()))
        End If
    Next rowIndex

    If pointCount < 2 Then Err.Raise vbObjectError + 4315, "CSystemSettingsReader", rangeName & " must contain at least two points."
End Sub

Private Function ApplicationDecimalSeparator() As String
    ApplicationDecimalSeparator = Mid$(CStr(1.1), 2, 1)
End Function





