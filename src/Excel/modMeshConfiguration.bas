Attribute VB_Name = "modMeshConfiguration"
Option Explicit

' ==========================================================================
' Общие настройки построения бетонной сетки
' ==========================================================================
' Читает обязательные параметры сетки из Excel, переводит входные единицы
' и передаёт их CFiberMeshBuilder. Одни настройки используются
' для параметрических сечений и точных контуров AutoCAD.

' Строит бетонную сетку по готовой области с едиными настройками Mesh.
' Шаги переводятся в мм один раз до передачи общему построителю сетки.
Public Function BuildConfiguredConcreteMesh(ByVal geometry As CGeometryRegion, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As CFiberMeshBuilder
    Dim options As TSectionMeshOptions, mesh As CFiberMeshBuilder
    options = ReadSectionMeshOptions(settings, units)
    Set mesh = New CFiberMeshBuilder
    mesh.BuildMesh geometry, options.StepX, options.StepY, options.MaterialID, options.BoundarySubdivisions, options.SourceReference, options.AutoStepFactor, options.CellGeometry
    Set BuildConfiguredConcreteMesh = mesh
End Function

' Читает режим шага и целое разбиение границы. В User переводит оба шага в мм;
' в Auto передаёт коэффициент общему построителю, не читая неиспользуемые ручные поля.
Public Function ReadSectionMeshOptions(ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem) As TSectionMeshOptions
    If settings Is Nothing Or units Is Nothing Then Err.Raise vbObjectError + 4139, _
        "modMeshConfiguration", "Настройки и единицы бетонной сетки не переданы."
    Dim options As TSectionMeshOptions, mode As String
    options.CellGeometry = ReadMeshCellGeometry(settings)
    mode = settings.GetRequiredChoice("Mesh.StepMode", Array("Auto 0.05", "Auto 0.1", "Auto 0.2", "User"))
    Select Case mode
        Case "Auto 0.05": options.AutoStepFactor = 0.05
        Case "Auto 0.1": options.AutoStepFactor = 0.1
        Case "Auto 0.2": options.AutoStepFactor = 0.2
        Case "User"
            options.StepX = ReadMeshStep(settings, units, "Mesh.StepX")
            options.StepY = ReadMeshStep(settings, units, "Mesh.StepY")
    End Select
    options.BoundarySubdivisions = 1
    If options.CellGeometry = "Adaptive" Then options.BoundarySubdivisions = ReadMeshBoundarySubdivisions(settings)
    options.MaterialID = 1
    options.SourceReference = "Mesh.StepMode: " & settings.InputLocation("Mesh.StepMode")
    If options.AutoStepFactor = 0# Then options.SourceReference = options.SourceReference & _
        "; Mesh.StepX: " & settings.InputLocation("Mesh.StepX") & "; Mesh.StepY: " & settings.InputLocation("Mesh.StepY")
    If settings.HasKey("Mesh.CellGeometry") Then options.SourceReference = options.SourceReference & "; Mesh.CellGeometry: " & settings.InputLocation("Mesh.CellGeometry")
    If options.CellGeometry = "Adaptive" Then options.SourceReference = options.SourceReference & "; Mesh.BoundarySubdivisions: " & settings.InputLocation("Mesh.BoundarySubdivisions")
    ReadSectionMeshOptions = options
End Function

' Читает единый способ получения фактической геометрии для генератора и AutoCAD.
' Отсутствующий новый ключ в прежнем внешнем снимке сохраняет старый режим;
' присутствующее пустое или неизвестное значение отклоняется адресно.
Public Function ReadMeshCellGeometry(ByVal settings As CSystemSettingsReader) As String
    If settings Is Nothing Then Err.Raise vbObjectError + 4139, "modMeshConfiguration", "Настройки способа бетонной сетки не переданы."
    ReadMeshCellGeometry = "Adaptive"
    If settings.HasKey("Mesh.CellGeometry") Then ReadMeshCellGeometry = settings.GetRequiredChoice("Mesh.CellGeometry", Array("Adaptive", "Contour"))
End Function

' Читает обязательный положительный шаг и переводит INPUT-длину в мм.
' При переполнении преобразования сообщает исходную настройку и способ исправления.
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

' Читает обязательное целое число подъячеек по каждой оси граничной ячейки.
' Пустое значение и число меньше 1 отклоняются с адресом настройки Mesh.
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
