# ======================================================================
# ДЛЯ ТЕСТОВ: чистые запись/чтение снимка, без решателя и построения схемы.
# ======================================================================
# Синтетические данные проверяют сериализацию, не изображают инженерный расчёт.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [int[]]$Elements=@(50,5000,50000),
    [int[]]$States=@(1,5),
    [ValidateRange(1,20)][int]$Repeats=5,
    [ValidateSet('All','File')][string]$Scope='All'
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Private evidence directory required.'}
if(Test-Path -LiteralPath (Join-Path $directory 'Measurements.json')){throw 'A new evidence directory is required.'}
[void][IO.Directory]::CreateDirectory($directory)
$hash=(Get-FileHash -LiteralPath $source).Hash
$copy=Join-Path $directory 'Codec.xlsm';Copy-Item -LiteralPath $source -Destination $copy
Add-Type 'using System;using System.Runtime.InteropServices;public static class SnapshotCodecOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$records=New-Object 'System.Collections.Generic.List[object]'

# Проверяет точную идентичность единственного тестового Excel перед закрытием.
function Test-OwnedSnapshotCodecExcel{
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SnapshotCodecOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedSnapshotCodecExcel)){throw 'Excel ownership not established.'}
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1;$app.ScreenUpdating=$false
    $book=$app.Workbooks.Open($copy,0,$false)
    [ordered]@{Source=$SourceWorkbook;SourceSHA256=$hash;PID=$owned;StartTicks=[string]$ticks;Synthetic=$true;SolverCalls=0;Elements=$Elements;States=$States;Repeats=$Repeats}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'Environment.json') -Encoding UTF8
    $helper=@'
Option Explicit
' ==================== ДЛЯ ТЕСТОВ ====================
' Одинаковые синтетические матрицы измеряют только хранение, без инженерных заявлений.
Private Declare PtrSafe Function TestQpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function TestQpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
Private Type TTestNumber
    Value As Double
End Type
Private Type TTestBits
    Low As Long
    High As Long
End Type
Private mTestSample As CCalculationSnapshot
Private mTestError As String
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает QPC, исключая создание данных и проверки равенства из таймера IO.
Private Function TestNow() As Double
    Dim ticks As Currency, frequency As Currency
    TestQpc ticks: TestQpf frequency: TestNow = ticks / frequency
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт полный контракт из 16 матриц с фиксированными числами и Unicode.
Public Sub TestPrepare(ByVal elementCount As Long, ByVal stateCount As Long)
    Set mTestSample = New CCalculationSnapshot: mTestSample.RunID = 45000.125: mTestSample.ElapsedSeconds = 1.25
    Dim keys As Variant, columns As Variant, rows As Long, key As String, index As Long, row As Long, column As Long, element As Long
    keys = Array("rngNDMElementResults", "rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionAnnotations", "Summary", "Strength.Headers", "Strength.Data", "Cracks.Headers", "Cracks.Data", "Stability.Labels", "Stability.Data", "LoadCase.Properties", "LoadCase.Annotations", "Calculation.Profiles")
    columns = Array(9, 16, 17, 6, 11, 24, 26, 49, 49, 72, 72, 84, 84, 6, 24, 3)
    Dim data() As Variant, names As Variant
    For index = 0 To UBound(keys)
        key = CStr(keys(index)): rows = 1
        Select Case key
            Case "rngNDMElementResults": rows = 1 + elementCount * stateCount
            Case "rngNDMSectionGeometry": rows = 1 + elementCount
            Case "Summary", "Strength.Headers", "Cracks.Headers": rows = 4
        End Select
        ReDim data(1 To rows, 1 To CLng(columns(index)))
        For column = 1 To UBound(data, 2): data(1, column) = "Header" & CStr(column): Next column
        If Left$(key, 6) = "rngNDM" Or Left$(key, 9) = "LoadCase." Then
            data(1, 1) = "RunID"
            For row = 2 To rows: data(row, 1) = mTestSample.RunID: Next row
        End If
        If key = "rngNDMElementResults" Then
            names = Array("RunID", "LoadCase", "ProfileId", "StateType", "ElementID", "MaterialType", "Strain", "Stress, MPa", "PhysicalState")
            For column = 1 To 9: data(1, column) = names(column - 1): Next column
            For row = 2 To rows
                element = (row - 2) Mod elementCount + 1
                data(row, 2) = "LC_01": data(row, 3) = "PR1": data(row, 4) = "State_" & CStr((row - 2) \ elementCount + 1)
                data(row, 5) = "Element_" & CStr(element) & ChrW$(&H3B1)
                data(row, 6) = "Concrete": If element Mod 20 = 0 Then data(row, 6) = "Rebar"
                data(row, 7) = (CDbl(row Mod 2001) - 1000#) / 10000000#: data(row, 8) = CDbl(data(row, 7)) * 30000#
                data(row, 9) = "Tension": If CDbl(data(row, 8)) < 0# Then data(row, 9) = "Compression"
            Next row
        ElseIf key = "rngNDMSectionGeometry" Then
            names = Array("RunID", "ElementID", "MaterialType", "X, mm", "Y, mm", "Area, mm2", "GeometryInterpretationStatus", "Width, mm", "Height, mm", "Diameter, mm", "Rotation", "LocalIx, mm4", "LocalIy, mm4", "LocalIxy, mm4", "Comment", "ExtraData")
            For column = 1 To 16: data(1, column) = names(column - 1): Next column
            For row = 2 To rows
                element = row - 1: data(row, 2) = "Element_" & CStr(element) & ChrW$(&H3B1)
                data(row, 3) = "Concrete": If element Mod 20 = 0 Then data(row, 3) = "Rebar"
                data(row, 4) = CDbl(element Mod 200): data(row, 5) = CDbl(element \ 200): data(row, 6) = 1#
                data(row, 7) = "Known": data(row, 8) = 1#: data(row, 9) = 1#: data(row, 10) = 0#: data(row, 11) = 0#
                data(row, 12) = 1# / 12#: data(row, 13) = 1# / 12#: data(row, 14) = 0#: data(row, 15) = "IO_" & ChrW$(&H3B1): data(row, 16) = "{}"
            Next row
        End If
        mTestSample.AddTable key, data
    Next index
    mTestSample.Validate
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает каждую ячейку и каждый бит Double независимо от двоичного кодека.
Private Sub TestEqual(ByVal actual As CCalculationSnapshot)
    Dim key As Variant, a As Variant, b As Variant, row As Long, column As Long, x As TTestNumber, y As TTestNumber, bitsX As TTestBits, bitsY As TTestBits
    For Each key In mTestSample.Keys()
        a = mTestSample.Table(CStr(key)): b = actual.Table(CStr(key))
        If UBound(a, 1) <> UBound(b, 1) Or UBound(a, 2) <> UBound(b, 2) Then Err.Raise 5, , "Matrix dimensions changed: " & CStr(key)
        For row = 1 To UBound(a, 1)
            For column = 1 To UBound(a, 2)
                If VarType(a(row, column)) <> VarType(b(row, column)) Then Err.Raise 5, , "Value type changed: " & CStr(key)
                If VarType(a(row, column)) = vbDouble Then
                    x.Value = a(row, column): y.Value = b(row, column): LSet bitsX = x: LSet bitsY = y
                    If bitsX.Low <> bitsY.Low Or bitsX.High <> bitsY.High Then Err.Raise 5, , "Double bits changed: " & CStr(key)
                ElseIf CStr(a(row, column)) <> CStr(b(row, column)) Then
                    Err.Raise 5, , "Value changed: " & CStr(key)
                End If
            Next column
        Next row
    Next key
End Sub
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Таймер включает ровно указанную операцию, полное сравнение выполняется после него.
Public Function TestMeasure(ByVal action As String) As Double
    On Error GoTo Failed
    mTestError = vbNullString
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim loaded As CCalculationSnapshot, writer As CNDMResultsWriter, started As Double
    started = TestNow()
    Select Case action
        Case "SaveFile": store.SaveFile ThisWorkbook, mTestSample
        Case "LoadFile": Set loaded = store.LoadFile(ThisWorkbook)
        Case "WriteSheet"
            Set writer = New CNDMResultsWriter
            writer.RestoreSnapshotBase ThisWorkbook, mTestSample, True, True
            store.SaveSheetExtra ThisWorkbook, mTestSample
        Case "ReadSheet": Set loaded = store.LoadSheet(ThisWorkbook)
        Case Else: Err.Raise 5, , "Unknown codec operation."
    End Select
    TestMeasure = TestNow() - started
    If Not loaded Is Nothing Then TestEqual loaded
    Exit Function
Failed:
    mTestError = Err.Description
    TestMeasure = -1#
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт точную причину неудачи в runner без неуправляемого окна VBE.
Public Function TestLastError() As String
    TestLastError = mTestError
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Освобождает большой тестовый набор после законченной серии измерений.
Public Sub TestRelease()
    Set mTestSample = Nothing
End Sub
'@
    $component=$book.VBProject.VBComponents.Add(1);$component.Name='modTestSnapshotCodec';$component.CodeModule.AddFromString($helper)
    $settings=$book.Names.Item('rngSystemSettings').RefersToRange;$values=$settings.Value2
    for($row=1;$row -le $values.GetLength(0);$row++){if([string]$values[$row,1] -ceq 'Results.SnapshotName'){$settings.Cells.Item($row,2).Value2='TestSnapshotCodec.ndm'}}
    $app.VBE.ActiveVBProject=$book.VBProject;$compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()};if($compile.Enabled){throw 'Codec measurement did not compile.'}
    $actions=@('SaveFile','LoadFile','WriteSheet','ReadSheet');if($Scope -eq 'File'){$actions=@('SaveFile','LoadFile')}
    foreach($count in $Elements){foreach($state in $States){
        [void]$app.Run("'$($book.Name)'!modTestSnapshotCodec.TestPrepare",[int]$count,[int]$state)
        for($repeat=0;$repeat -le $Repeats;$repeat++){
            foreach($action in $actions){
                $seconds=[double]$app.Run("'$($book.Name)'!modTestSnapshotCodec.TestMeasure",$action)
                if($seconds -lt 0){throw [string]$app.Run("'$($book.Name)'!modTestSnapshotCodec.TestLastError")}
                $record=[ordered]@{Elements=$count;States=$state;Rows=1+$count*$state;Action=$action;Repeat=$repeat;Warmup=($repeat -eq 0);Seconds=$seconds;ExactAllCells=$true;Synthetic=$true}
                $records.Add($record);$records|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'Measurements.json') -Encoding UTF8
                Write-Output ($record|ConvertTo-Json -Compress)
            }
        }
        $path=Join-Path $directory ('Payload_'+$count+'_'+$state+'.ndm');Copy-Item -LiteralPath (Join-Path $directory 'TestSnapshotCodec.ndm') -Destination $path
        [ordered]@{Elements=$count;States=$state;SHA256=(Get-FileHash -LiteralPath $path).Hash;Bytes=(Get-Item -LiteralPath $path).Length}|ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $directory 'Payloads.jsonl') -Encoding UTF8
        [void]$app.Run("'$($book.Name)'!modTestSnapshotCodec.TestRelease")
    }}
}finally{
    if(Test-OwnedSnapshotCodecExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
if((Get-FileHash -LiteralPath $source).Hash -cne $hash){throw 'Source workbook changed.'}
Write-Output 'SNAPSHOT_CODEC_COMPLETED'
