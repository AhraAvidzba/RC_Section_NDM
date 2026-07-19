VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFiberMeshBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mFibers() As TFiber
Private mFiberCount As Long
Private mBuildSeconds As Double

Public Sub BuildMesh(ByVal geometry As ISectionGeometry, ByVal stepX As Double, ByVal stepY As Double, _
        Optional ByVal materialID As Long = 1, Optional ByVal boundarySubdivisions As Long = 1)

    Dim message As String
    If Not geometry.IsValid(message) Then Err.Raise vbObjectError + 2100, "CFiberMeshBuilder", message
    If stepX <= 0# Then Err.Raise vbObjectError + 2101, "CFiberMeshBuilder", "Mesh step X must be positive."
    If stepY <= 0# Then Err.Raise vbObjectError + 2102, "CFiberMeshBuilder", "Mesh step Y must be positive."
    If boundarySubdivisions < 1 Then boundarySubdivisions = 1

    Dim t0 As Double
    t0 = Timer

    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    minX = geometry.MinX
    maxX = geometry.MaxX
    minY = geometry.MinY
    maxY = geometry.MaxY

    If maxX <= minX + GEOM_TOLERANCE Or maxY <= minY + GEOM_TOLERANCE Then
        Err.Raise vbObjectError + 2103, "CFiberMeshBuilder", _
            "Degenerate geometry bounds: minX=" & CStr(minX) & _
            "; maxX=" & CStr(maxX) & "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY)
    End If

    Dim nx As Long
    Dim ny As Long
    nx = CLng(Fix((maxX - minX) / stepX)) + 1
    ny = CLng(Fix((maxY - minY) / stepY)) + 1
    If nx < 1 Or ny < 1 Then Err.Raise vbObjectError + 2104, "CFiberMeshBuilder", "Empty mesh bounding rectangle."

    ReDim mFibers(1 To nx * ny)
    mFiberCount = 0

    Dim ix As Long
    Dim iy As Long
    Dim x As Double
    Dim y As Double
    Dim cellMinX As Double
    Dim cellMinY As Double
    Dim allCornersInside As Boolean
    Dim anyCornerInside As Boolean

    For iy = 0 To ny - 1
        cellMinY = minY + iy * stepY
        y = cellMinY + stepY / 2#
        If y <= maxY + GEOM_TOLERANCE Then
            For ix = 0 To nx - 1
                cellMinX = minX + ix * stepX
                x = cellMinX + stepX / 2#
                If x <= maxX + GEOM_TOLERANCE Then
                    If boundarySubdivisions = 1 Then
                        If geometry.ContainsPoint(x, y) Then AddFiber x, y, stepX, stepY, materialID
                    Else
                        GetCornerStatus geometry, cellMinX, cellMinY, stepX, stepY, allCornersInside, anyCornerInside
                        If allCornersInside Then
                            AddFiber x, y, stepX, stepY, materialID
                        Else
                            AddSubcellFibers geometry, cellMinX, cellMinY, stepX, stepY, boundarySubdivisions, materialID
                        End If
                    End If
                End If
            Next ix
        End If
    Next iy

    If mFiberCount <= 0 Then
        Err.Raise vbObjectError + 2105, "CFiberMeshBuilder", _
            "Mesh is empty: minX=" & CStr(minX) & "; maxX=" & CStr(maxX) & _
            "; minY=" & CStr(minY) & "; maxY=" & CStr(maxY) & _
            "; stepX=" & CStr(stepX) & "; stepY=" & CStr(stepY)
    End If

    mBuildSeconds = Timer - t0
End Sub

Private Sub AddSubcellFibers(ByVal geometry As ISectionGeometry, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal stepX As Double, ByVal stepY As Double, ByVal boundarySubdivisions As Long, ByVal materialID As Long)
    Dim subX As Double
    Dim subY As Double
    subX = stepX / CDbl(boundarySubdivisions)
    subY = stepY / CDbl(boundarySubdivisions)

    Dim ix As Long
    Dim iy As Long
    Dim x As Double
    Dim y As Double
    For iy = 0 To boundarySubdivisions - 1
        y = cellMinY + subY / 2# + iy * subY
        For ix = 0 To boundarySubdivisions - 1
            x = cellMinX + subX / 2# + ix * subX
            If geometry.ContainsPoint(x, y) Then AddFiber x, y, subX, subY, materialID
        Next ix
    Next iy
End Sub

Private Sub GetCornerStatus(ByVal geometry As ISectionGeometry, ByVal cellMinX As Double, ByVal cellMinY As Double, _
        ByVal width As Double, ByVal height As Double, ByRef allInside As Boolean, ByRef anyInside As Boolean)
    Dim p1 As Boolean
    Dim p2 As Boolean
    Dim p3 As Boolean
    Dim p4 As Boolean
    p1 = geometry.ContainsPoint(cellMinX, cellMinY)
    p2 = geometry.ContainsPoint(cellMinX + width, cellMinY)
    p3 = geometry.ContainsPoint(cellMinX + width, cellMinY + height)
    p4 = geometry.ContainsPoint(cellMinX, cellMinY + height)
    allInside = p1 And p2 And p3 And p4
    anyInside = p1 Or p2 Or p3 Or p4
End Sub

Private Sub AddFiber(ByVal x As Double, ByVal y As Double, ByVal width As Double, ByVal height As Double, ByVal materialID As Long)
    mFiberCount = mFiberCount + 1
    If mFiberCount > UBound(mFibers) Then ReDim Preserve mFibers(1 To UBound(mFibers) * 2)
    mFibers(mFiberCount).X = x
    mFibers(mFiberCount).Y = y
    mFibers(mFiberCount).Width = width
    mFibers(mFiberCount).Height = height
    mFibers(mFiberCount).Area = width * height
    mFibers(mFiberCount).FillFactor = 1#
    mFibers(mFiberCount).MaterialID = materialID
End Sub

Public Property Get FiberCount() As Long
    FiberCount = mFiberCount
End Property

Public Property Get BuildSeconds() As Double
    BuildSeconds = mBuildSeconds
End Property

Public Property Get FiberX(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberX = mFibers(index).X
End Property

Public Property Get FiberY(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberY = mFibers(index).Y
End Property

Public Property Get FiberWidth(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberWidth = mFibers(index).Width
End Property

Public Property Get FiberHeight(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberHeight = mFibers(index).Height
End Property

Public Property Get FiberArea(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberArea = mFibers(index).Area
End Property

Public Property Get FiberFillFactor(ByVal index As Long) As Double
    CheckFiberIndex index
    FiberFillFactor = mFibers(index).FillFactor
End Property

Public Property Get FiberMaterialID(ByVal index As Long) As Long
    CheckFiberIndex index
    FiberMaterialID = mFibers(index).MaterialID
End Property

Private Sub CheckFiberIndex(ByVal index As Long)
    If index < 1 Or index > mFiberCount Then
        Err.Raise vbObjectError + 2106, "CFiberMeshBuilder", "Fiber index out of range."
    End If
End Sub







