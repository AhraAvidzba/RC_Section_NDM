Attribute VB_Name = "modTestPolygonSpec"
Option Explicit

' ДЛЯ ТЕСТОВ: структура общего описания вершин многоугольника.
Private mPassed As Long, mFailed As Long, mReport As String
Private Type TDoubleValue
    Value As Double
End Type
Private Type TDoubleWords
    Low As Long
    High As Long
End Type

' ДЛЯ ТЕСТОВ: порядок/источники, более шести вершин, ошибки и повторное чтение снимка.
Public Function RunPolygonSpecTests() As String
    mPassed = 0: mFailed = 0: mReport = ""
    On Error GoTo Failed
    Dim polygon As CPolygonSpec, n As Variant, i As Long, errorCode As Long, reason As String
    Dim vertices As Variant, radii As Variant, names As Variant, number As TDoubleValue, words As TDoubleWords
    For Each n In Array(4, 5, 6, 9)
        Set polygon = New CPolygonSpec
        Check "new.open." & CStr(n), Not polygon.IsSealed And polygon.Count = 0
        For i = 1 To CLng(n)
            polygon.AddVertex "V" & CStr(i), i * 10#, -i * 20#, i - 1#, "Config!C" & CStr(i)
        Next i
        Check "count.before.seal." & CStr(n), polygon.Count = CLng(n)
        polygon.Seal: polygon.Seal
        Check "sealed." & CStr(n), polygon.IsSealed
        For i = 1 To CLng(n)
            Check "record." & CStr(n) & "." & CStr(i), polygon.VertexID(i) = "V" & CStr(i) And _
                polygon.X(i) = i * 10# And polygon.Y(i) = -i * 20# And polygon.FilletRadius(i) = i - 1# And _
                polygon.SourceReference(i) = "Config!C" & CStr(i)
        Next i
        polygon.CopyKernelData vertices, radii, names
        vertices(0) = Array(-1#, -2#): radii(0) = 999#: names(0) = "Changed"
        Check "kernel.copy.independent." & CStr(n), polygon.X(1) = 10# And polygon.FilletRadius(1) = 0# And polygon.VertexID(1) = "V1"
        On Error Resume Next
        polygon.AddVertex "After", 0#, 0#, 0#, "Config!After"
        errorCode = Err.Number: reason = Err.Description: Err.Clear
        On Error GoTo Failed
        Check "reject.after.seal." & CStr(n), errorCode <> 0 And polygon.Count = CLng(n) And InStr(reason, "Config!After") > 0
    Next n
    Set polygon = New CPolygonSpec
    polygon.AddVertex "A", 0#, 0#
    polygon.AddVertex "a", 1#, 0#
    Check "ids.binary.case", polygon.Count = 2
    On Error Resume Next
    polygon.AddVertex "A", 0#, 1#, 0#, "Config!Duplicate"
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "reject.duplicate", errorCode <> 0 And polygon.Count = 2 And InStr(reason, "Config!Duplicate") > 0
    On Error Resume Next
    polygon.AddVertex " ", 0#, 1#, 0#, "Config!Empty"
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.empty.id", errorCode <> 0 And polygon.Count = 2
    On Error Resume Next
    polygon.AddVertex "Negative", 0#, 1#, -1#, "Config!Radius"
    errorCode = Err.Number: reason = Err.Description: Err.Clear
    On Error GoTo Failed
    Check "reject.negative.radius", errorCode <> 0 And polygon.Count = 2 And InStr(reason, "Config!Radius") > 0
    On Error Resume Next
    polygon.Seal
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.incomplete", errorCode <> 0 And Not polygon.IsSealed
    words.High = &H7FF00000: LSet number = words
    On Error Resume Next
    polygon.AddVertex "Inf", number.Value, 1#
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.infinity", errorCode <> 0 And polygon.Count = 2
    words.Low = 1: LSet number = words
    On Error Resume Next
    polygon.AddVertex "NaN", 1#, 1#, number.Value
    errorCode = Err.Number: Err.Clear
    On Error GoTo Failed
    Check "reject.nan", errorCode <> 0 And polygon.Count = 2
    polygon.AddVertex "B", 0#, 1#: polygon.Seal
    Check "recover.after.failed.seal", polygon.Count = 3 And polygon.IsSealed
    GoTo Finished
Failed:
    Check "runtime: " & CStr(Err.Number) & "; " & Err.Description, False
Finished:
    RunPolygonSpecTests = mReport & "TOTAL POLYGON SPEC: passed=" & CStr(mPassed) & "; failed=" & CStr(mFailed) & vbCrLf
End Function
Private Sub Check(ByVal name As String, ByVal condition As Boolean)
    If condition Then
        mPassed = mPassed + 1: mReport = mReport & "PASS: " & name & vbCrLf
    Else
        mFailed = mFailed + 1: mReport = mReport & "FAIL: " & name & vbCrLf
    End If
End Sub
