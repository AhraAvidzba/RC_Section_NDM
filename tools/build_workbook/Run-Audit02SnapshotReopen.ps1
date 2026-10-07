# Проверяет сериализованный Results после изменения Config и повторного открытия.
# Работает только с временной копией; исходная книга и пользовательские входы не меняются.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
$source = (Resolve-Path -LiteralPath $WorkbookPath).Path
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$folder = Join-Path ([IO.Path]::GetTempPath()) ('RC_NDM_SnapshotReopen_' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $folder | Out-Null
$fixture = Join-Path $folder 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $fixture
$report = [IO.Path]::GetFullPath($ReportPath)
$lines = New-Object System.Collections.Generic.List[string]
$excel = $null
$book = $null
$failed = $false

# Читает существующее COM-свойство напрямую, не скрывая потерю объекта.
function Get-RequiredComProperty([object]$Target,[string]$Property) {
    if ($null -eq $Target) {throw "COM target unavailable: $Property"}
    $value=$Target.GetType().InvokeMember($Property,[Reflection.BindingFlags]::GetProperty,$null,$Target,$null)
    if ($null -eq $value) {throw "COM property unavailable: $Property"}
    return ,$value
}

try {
    $lines.Add("SOURCE: $source; SHA256=$sourceHash")
    $lines.Add("FIXTURE: $fixture")
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    $books=Get-RequiredComProperty $excel 'Workbooks'
    $book = $books.Open($fixture)
    $project=Get-RequiredComProperty $book 'VBProject'
    $components=Get-RequiredComProperty $project 'VBComponents'
    $module=Get-RequiredComProperty $components.Item('modTestWorkbookInterface') 'CodeModule'
    $module.AddFromString(@'
' ДЛЯ ТЕСТОВ: временный entrypoint подготовки сохраненной книги.
Public Function Audit02PrepareReopenFixture() As String
    PrepareCircleInput
    SetSystemSetting "Plot.AutoUpdateAfterCalculation", "No"
    SetSystemSetting "AutoCAD.Export.CombinationID", "LC1"
    Audit02PrepareReopenFixture = RunSectionCalculationForWorkbook(ThisWorkbook, False)
End Function

' ДЛЯ ТЕСТОВ: меняет настройки через общий reader/helper всех таблиц Config.
Public Function Audit02ChangeReopenMaterials() As String
    Dim keys As Variant, key As Variant, value As String
    keys = Array("General.DiagramExtension", "Concrete.E", "Concrete.R.ULS(I)", _
        "Concrete.R.SLS(II)", "Steel.E", "Steel.R.ULS(I)", "Steel.R.SLS(II)")
    For Each key In keys
        value = CStr(GetSystemSetting(CStr(key)))
        If CStr(key) = "General.DiagramExtension" Then
            If value = "Yes" Then value = "No" Else value = "Yes"
        Else
            value = CStr(CDbl(value) * 0.5)
        End If
        SetSystemSetting CStr(key), value
        Audit02ChangeReopenMaterials = Audit02ChangeReopenMaterials & CStr(key) & "=" & value & vbCrLf
    Next key
End Function

' ДЛЯ ТЕСТОВ: сериализует все таблицы через штатный reader, без нового НДС.
Public Function Audit02ReadReopenSnapshot() As String
    Dim beforeSolves As Long
    beforeSolves = SectionEquilibriumSolveCount()
    Dim result As String
    Dim quantity As Variant
    For Each quantity In Array("Stress", "Strain")
        SetProfileSetting "PR2", "Visualization.Quantity", CStr(quantity)
        result = result & vbCrLf & CStr(quantity) & ":" & Audit02ReadExportSnapshotForTests(ThisWorkbook)
    Next quantity
    Dim names As Variant
    names = Array("rngNDMSectionGeometry", "rngNDMSectionContours", "rngNDMElementResults", "rngNDMSectionProperties", _
        "rngNDMSectionAnnotations", "rngNDMMaterialDiagrams", "rngBatchSummary")
    Dim name As Variant, table As Variant
    Dim r As Long, c As Long
    For Each name In names
        table = ResultTable(CStr(name))
        result = result & vbCrLf & CStr(name) & ":" & CStr(UBound(table, 1)) & "x" & CStr(UBound(table, 2))
        For r = 1 To UBound(table, 1)
            For c = 1 To UBound(table, 2)
                result = result & "|" & CStr(table(r, c))
            Next c
        Next r
    Next name
    UpdateSectionPlotForWorkbook ThisWorkbook
    If SectionEquilibriumSolveCount() <> beforeSolves Then Err.Raise 5, "Audit02Reopen", "Reader/plot запустил новый solve"
    Audit02ReadReopenSnapshot = result
End Function
'@
    )
    $lines.Add('STAGE: prepare calculation fixture')
    $message = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.Audit02PrepareReopenFixture")
    if ($message -notmatch 'Расчет завершен') { throw $message }
    $lines.Add('STAGE: read prepared snapshot')
    $before = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.Audit02ReadReopenSnapshot")
    $changed = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.Audit02ChangeReopenMaterials")
    $lines.Add("CHANGED: $changed")
    $book.Save()
    $book.Close($false)
    [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    $book = $null
    $books=Get-RequiredComProperty $excel 'Workbooks'
    $book = $books.Open($fixture, $null, $true)
    $solvesBefore = [long]$excel.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
    $after = [string]$excel.Run("'$($book.Name)'!modTestWorkbookInterface.Audit02ReadReopenSnapshot")
    $solvesAfter = [long]$excel.Run("'$($book.Name)'!modSolverWorkStats.SectionEquilibriumSolveCount")
    if ($before -cne $after) { throw 'Сохраненный snapshot изменился после Config/save/reopen.' }
    if ($solvesBefore -ne 0 -or $solvesAfter -ne 0) { throw 'После повторного открытия выполнялся solve.' }
    $lines.Add("SNAPSHOT_EXACT_MATCH: True; characters=$($after.Length)")
    $lines.Add('REOPEN_READ_EXPORT_PLOT_SOLVES: 0')
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Исходная книга изменена.' }
    $lines.Add('SOURCE_UNCHANGED: True')
} catch {
    $failed = $true
    $lines.Add("SCRIPT ERROR: $($_.Exception.Message); $($_.ScriptStackTrace)")
} finally {
    if ($book) { try { $book.Close($false) } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null }
    if ($excel) { try { $excel.Quit() } catch {}; [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $lines | Set-Content -LiteralPath $report -Encoding UTF8
    $lines
}
if ($failed) { exit 1 }
