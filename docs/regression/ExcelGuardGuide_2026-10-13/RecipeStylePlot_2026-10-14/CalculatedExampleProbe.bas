' ДЛЯ ТЕСТОВ: сведения сохранённого расчёта; добавляется только в отдельные копии.
Public Function GoalCalculatedExampleSnapshot() As String
    Dim settings As CSystemSettingsReader, data As CSectionPlotDataReader, i As Long, text As String
    Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
    Set data = New CSectionPlotDataReader: data.LoadFromWorkbook ThisWorkbook, settings
    text = "STATE" & vbTab & data.DirectStateStatus & vbTab & CStr(data.GeometryOnly) & vbTab & data.LoadCase & vbCrLf
    text = text & "CENTERS" & vbTab & GoalExampleNumber(data.PrincipalAxesCentroidX("Concrete")) & vbTab & GoalExampleNumber(data.PrincipalAxesCentroidY("Concrete"))
    text = text & vbTab & GoalExampleNumber(data.CentroidX) & vbTab & GoalExampleNumber(data.CentroidY)
    text = text & vbTab & GoalExampleNumber(data.PrincipalAxesAngle("Concrete")) & vbTab & GoalExampleNumber(data.PrincipalAngle)
    text = text & vbTab & GoalExampleNumber(data.LoadReferenceX) & vbTab & GoalExampleNumber(data.LoadReferenceY) & vbCrLf
    For i = 1 To data.Count
        text = text & "ELEMENT" & vbTab & data.ElementID(i) & vbTab & data.MaterialType(i) & vbTab & GoalExampleNumber(data.X(i)) & vbTab & GoalExampleNumber(data.Y(i)) & vbTab & GoalExampleNumber(data.Area(i)) & vbTab & GoalExampleNumber(data.ResultValue(i)) & vbCrLf
    Next i
    For i = 1 To data.AnnotationCount
        text = text & "ANNOTATION" & vbTab & data.AnnotationType(i) & vbTab & GoalExampleNumber(data.AnnotationStartX(i)) & vbTab & GoalExampleNumber(data.AnnotationStartY(i)) & vbTab & GoalExampleNumber(data.AnnotationEndX(i)) & vbTab & GoalExampleNumber(data.AnnotationEndY(i)) & vbCrLf
    Next i
    GoalCalculatedExampleSnapshot = text
End Function

Private Function GoalExampleNumber(ByVal value As Double) As String
    GoalExampleNumber = Replace(CStr(value), ",", ".")
End Function
