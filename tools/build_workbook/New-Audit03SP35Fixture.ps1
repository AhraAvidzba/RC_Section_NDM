# Добавляет воспроизводящий тест таблицы 7.21 в существующий test module
# отдельной копии. Production VBA и исходная книга остаются неизменными.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$OutputWorkbook,
    [Parameter(Mandatory=$true)][string]$FragmentPath
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$destination = [IO.Path]::GetFullPath((Join-Path $root $OutputWorkbook))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $destination.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Test fixture must stay in docs/regression/Audit03.' }
if (Test-Path -LiteralPath $destination) { throw 'Existing fixture must not be overwritten.' }
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$fragment = Get-Content -LiteralPath (Join-Path $root $FragmentPath) -Raw -Encoding UTF8
if (-not $fragment.Contains('Public Function RunAudit03SP35TableInvalidReproducer() As String')) { throw 'Unexpected test fragment.' }
Copy-Item -LiteralPath $source -Destination $destination
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($destination,0,$false)
    $before=@{}
    foreach ($component in $book.VBProject.VBComponents) {
        if ($component.Name -eq 'modTestBatchCalculation') { continue }
        $before[$component.Name]=if ($component.CodeModule.CountOfLines) { $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines) } else { '' }
    }
    $module=$book.VBProject.VBComponents.Item('modTestBatchCalculation').CodeModule
    $body=$module.Lines(1,$module.CountOfLines)
    if ($body.Contains('RunAudit03SP35TableInvalidReproducer')) { throw 'Test fragment is already present.' }
    $fragment=($fragment -split '\r?\n') -join "`r`n"
    $module.AddFromString("`r`n" + $fragment)
    foreach ($component in $book.VBProject.VBComponents) {
        if ($component.Name -eq 'modTestBatchCalculation') { continue }
        $after=if ($component.CodeModule.CountOfLines) { $component.CodeModule.Lines(1,$component.CodeModule.CountOfLines) } else { '' }
        if ($after -cne $before[$component.Name]) { throw ('Unexpected code change: ' + $component.Name) }
    }
    $book.Save()
    $book.Close($false); $book=$null
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Source workbook changed.' }
    Write-Output "SP35_REPRODUCER_FIXTURE: sourceSHA=$sourceHash; sourceUnchanged=True; productionCodeUnchanged=True; fixture=$OutputWorkbook"
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null }
}
