Attribute VB_Name = "modMeshConfiguration"
Option Explicit

' ==========================================================================
' Общие настройки дискретизации бетона для Generated и точных CAD-контуров
' ==========================================================================
' Excel-адаптер нормализует обязательные INPUT-параметры и передает их
' обычному CFiberMeshBuilder. Математическая геометрия Config не читает.

Public Function BuildConfiguredConcreteMesh(ByVal geometry As CGeometryRegion, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CFiberMeshBuilder
    If settings Is Nothing Or units Is Nothing Then Err.Raise vbObjectError + 4139, _
        "modMeshConfiguration", "Настройки и единицы бетонной сетки не переданы."
    Dim mesh As CFiberMeshBuilder, stepX As Double, stepY As Double
    stepX = ReadMeshStep(settings, units, "Mesh.StepX")
    stepY = ReadMeshStep(settings, units, "Mesh.StepY")
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, stepX, stepY, 1, ReadMeshBoundarySubdivisions(settings)
    Set BuildConfiguredConcreteMesh = mesh
End Function

Private Function ReadMeshStep(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, ByVal key As String) As Double
    Dim value As Double
    value = settings.GetRequiredDouble(key)
    If value <= 0# Then Err.Raise vbObjectError + 4139, "modMeshConfiguration", _
        settings.InputErrorMessage(key, "Шаг бетонной сетки должен быть положительным.", _
            "Введите число больше нуля в текущей INPUT-единице длины.")
    On Error GoTo InvalidConversion
    ReadMeshStep = units.InputLengthToInternal(value)
    Exit Function
InvalidConversion:
    If Err.Number <> 6 Then Err.Raise Err.Number, Err.Source, Err.Description
    Err.Raise vbObjectError + 4139, "modMeshConfiguration", _
        settings.InputErrorMessage(key, "Шаг слишком велик для пересчета в мм.", _
            "Уменьшите число и проверьте INPUT-единицу длины.")
End Function

Public Function ReadMeshBoundarySubdivisions(ByVal settings As CSystemSettingsReader) As Long
    If settings Is Nothing Then Err.Raise vbObjectError + 4139, "modMeshConfiguration", "Настройки разбиения бетонной сетки не переданы."
    If settings.HasKey("Mesh.BoundarySubdivisions") Then
        Dim rawValue As String
        rawValue = Trim$(settings.GetRawString("Mesh.BoundarySubdivisions"))
        If Len(rawValue) = 0 Or UCase$(rawValue) = "TODO" Then Err.Raise vbObjectError + 4139, "modMeshConfiguration", _
            settings.InputErrorMessage("Mesh.BoundarySubdivisions", "Число подъячеек не задано.", "Введите целое число не меньше 1.")
    End If
    ReadMeshBoundarySubdivisions = settings.GetRequiredLong("Mesh.BoundarySubdivisions")
    If ReadMeshBoundarySubdivisions < 1 Then Err.Raise vbObjectError + 4139, "modMeshConfiguration", _
        settings.InputErrorMessage("Mesh.BoundarySubdivisions", "Число подъячеек меньше 1.", "Введите целое число не меньше 1.")
End Function
