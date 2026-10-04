# Создает только изолированную диагностическую книгу: production source не меняется.
# Отметки перед стадиями writer-а локализуют расход памяти без изменения данных,
# расчетных критериев, форматирования или общих настроек Excel.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$TargetWorkbook,
    [Parameter(Mandatory=$true)][string]$TracePath
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$source = (Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$target = [IO.Path]::GetFullPath((Join-Path $root $TargetWorkbook))
$trace = [IO.Path]::GetFullPath((Join-Path $root $TracePath))
foreach ($path in @($source, $target, $trace)) {
    if (-not $path.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Diagnostic paths must stay in docs/regression/Audit03.'
    }
}
if (Test-Path -LiteralPath $target) { throw 'Diagnostic target already exists.' }
if (Test-Path -LiteralPath $trace) { throw 'Diagnostic trace already exists.' }
Copy-Item -LiteralPath $source -Destination $target
$printAreas = @(Get-WorkbookPrintAreas $target)
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($target)
    $module = $book.VBProject.VBComponents.Item('CBatchResultWriter').CodeModule
    $body = $module.Lines(1, $module.CountOfLines)
    if ($body -match 'Audit03TraceWriterMemory') { throw 'Fixture already instrumented.' }
    $body = [regex]::Replace($body, '(?m)^([ \t]*)stage = ([^\r\n]+)(?=\r?$)', {
        param($match)
        $match.Value + "`r`n" + $match.Groups[1].Value + 'Audit03TraceWriterMemory stage'
    })
    if ([regex]::Matches($body, 'Audit03TraceWriterMemory stage').Count -lt 20) {
        throw 'Writer stage instrumentation is incomplete.'
    }
    $body = $body.Replace('WriteStabilitySummary workbook, batch, units',
        "WriteStabilitySummary workbook, batch, units`r`n    Audit03TraceWriterMemory `"Complete`"")
    $escapedTrace = $trace.Replace('"', '""')
    $body += @"

' ДЛЯ ТЕСТОВ: только временная fixture, не production source.
' Записывает память до очередной стадии, чтобы отделить writer от solve.
Private Sub Audit03TraceWriterMemory(ByVal stage As String)
    Dim process As Object, channel As Integer
    channel = FreeFile
    Open "$escapedTrace" For Append As #channel
    For Each process In GetObject("winmgmts:").ExecQuery( _
            "SELECT ProcessId, WorkingSetSize, PrivatePageCount FROM Win32_Process WHERE Name='EXCEL.EXE'")
        Print #channel, stage & "; pid=" & CStr(process.ProcessId) & _
            "; workingSet=" & CStr(process.WorkingSetSize) & _
            "; privateBytes=" & CStr(process.PrivatePageCount)
    Next process
    Close #channel
End Sub
"@
    $module.DeleteLines(1, $module.CountOfLines)
    $module.AddFromString($body)
    $book.Save()
    $book.Close($false)
    $book = $null
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
Restore-WorkbookPrintAreas $target $printAreas
Write-Output "WRITER_TRACE_FIXTURE: $target; source unchanged; trace=$trace"
