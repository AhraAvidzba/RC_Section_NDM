VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRebarLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mBarID() As String
Private mX() As Double
Private mY() As Double
Private mDiameter() As Double
Private mArea() As Double
Private mSteelClass() As String
Private mComment() As String
Private mCount As Long

Public Sub Clear()
    Erase mBarID
    Erase mX
    Erase mY
    Erase mDiameter
    Erase mArea
    Erase mSteelClass
    Erase mComment
    mCount = 0
End Sub

Public Sub AddBar(ByVal barName As String, ByVal xCoord As Double, ByVal yCoord As Double, _
        ByVal diameter As Double, ByVal area As Double, ByVal steelClass As String, _
        Optional ByVal comment As String = vbNullString, Optional ByVal geometry As ISectionGeometry = Nothing)

    If Len(Trim$(barName)) = 0 Then Err.Raise vbObjectError + 3100, "CRebarLayout", "Rebar ID is empty."
    If diameter <= 0# Then Err.Raise vbObjectError + 3101, "CRebarLayout", "Rebar diameter must be positive."

    Dim actualArea As Double
    If area <= 0# Then
        actualArea = GEOM_PI * diameter * diameter / 4#
    Else
        actualArea = area
    End If
    If actualArea <= 0# Then Err.Raise vbObjectError + 3102, "CRebarLayout", "Rebar area must be positive."

    Dim i As Long
    For i = 1 To mCount
        If StrComp(mBarID(i), barName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 3103, "CRebarLayout", "Duplicate rebar ID: " & barName
        If Abs(mX(i) - xCoord) <= GEOM_TOLERANCE And Abs(mY(i) - yCoord) <= GEOM_TOLERANCE Then
            Err.Raise vbObjectError + 3104, "CRebarLayout", "Duplicate rebar coordinates."
        End If
    Next i

    If Not geometry Is Nothing Then
        If Not geometry.ContainsPoint(xCoord, yCoord) Then Err.Raise vbObjectError + 3105, "CRebarLayout", "Rebar center is outside section."
    End If

    mCount = mCount + 1
    ReDim Preserve mBarID(1 To mCount)
    ReDim Preserve mX(1 To mCount)
    ReDim Preserve mY(1 To mCount)
    ReDim Preserve mDiameter(1 To mCount)
    ReDim Preserve mArea(1 To mCount)
    ReDim Preserve mSteelClass(1 To mCount)
    ReDim Preserve mComment(1 To mCount)

    mBarID(mCount) = barName
    mX(mCount) = xCoord
    mY(mCount) = yCoord
    mDiameter(mCount) = diameter
    mArea(mCount) = actualArea
    mSteelClass(mCount) = steelClass
    mComment(mCount) = comment
End Sub

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get BarID(ByVal index As Long) As String
    CheckIndex index
    BarID = mBarID(index)
End Property

Public Property Get X(ByVal index As Long) As Double
    CheckIndex index
    X = mX(index)
End Property

Public Property Get Y(ByVal index As Long) As Double
    CheckIndex index
    Y = mY(index)
End Property

Public Property Get Diameter(ByVal index As Long) As Double
    CheckIndex index
    Diameter = mDiameter(index)
End Property

Public Property Get Area(ByVal index As Long) As Double
    CheckIndex index
    Area = mArea(index)
End Property

Public Property Get SteelClass(ByVal index As Long) As String
    CheckIndex index
    SteelClass = mSteelClass(index)
End Property

Public Property Get Comment(ByVal index As Long) As String
    CheckIndex index
    Comment = mComment(index)
End Property

Private Sub CheckIndex(ByVal index As Long)
    If index < 1 Or index > mCount Then Err.Raise vbObjectError + 3106, "CRebarLayout", "Rebar index out of range."
End Sub







