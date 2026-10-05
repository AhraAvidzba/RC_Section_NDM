# Обновляет тестовый модуль в изолированной Audit03-книге, не меняя production.
# Для существующего CTestLimitSearchProblem разрешена замена тела, но не
# создание нового класса: так усиленный reproducer проверяет прежний Search.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [string]$TestModule = 'modTestBatchCalculation',
    [switch]$RestoreGeometryConstants
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$bookPath = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $bookPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Only isolated Audit03 fixtures may be modified.' }
$isTestClass = $TestModule -eq 'CTestLimitSearchProblem'
if (-not $isTestClass -and $TestModule -notmatch '^modTest[A-Za-z0-9]+$') { throw 'Expected an approved test module.' }
$extension = if ($isTestClass) { 'cls' } else { 'bas' }
$source = Join-Path $root ("tests/$TestModule.$extension")
$sourceLines = [IO.File]::ReadAllText($source, [Text.Encoding]::UTF8) -split '\r?\n'
if ($isTestClass) {
    $start = [Array]::IndexOf($sourceLines, 'Option Explicit')
    if ($start -lt 0) { throw 'Test class must have Option Explicit.' }
    $body = ($sourceLines | Select-Object -Skip $start) -join "`r`n"
} else {
    $body = ($sourceLines | Where-Object { $_ -notmatch '^Attribute VB_' }) -join "`r`n"
}
$printAreas = @(Get-WorkbookPrintAreas $bookPath)
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($bookPath)
    $component = $null
    foreach ($candidate in $book.VBProject.VBComponents) {
        if ($candidate.Name -eq $TestModule) { $component = $candidate; break }
    }
    if ($null -eq $component) {
        if ($isTestClass) { throw 'The approved test class must already exist in the fixture.' }
        $component = $book.VBProject.VBComponents.Add(1)
        $component.Name = $TestModule
    }
    $expectedType = if ($isTestClass) { 2 } else { 1 }
    if ($component.Type -ne $expectedType) { throw 'Unexpected test component type.' }
    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) { $module.DeleteLines(1, $module.CountOfLines) }
    $module.AddFromString($body)
    if ($RestoreGeometryConstants) {
        # VBE rounds visible long literals. Restore the unchanged source last,
        # so both geometry fixtures compile the same exact GEOM_PI literal.
        $constantSource = Join-Path $root 'src/Common/modGeometryTypes.bas'
        $constantBody = ([IO.File]::ReadAllText($constantSource, [Text.Encoding]::UTF8) -split '\r?\n' | Where-Object { $_ -notmatch '^Attribute VB_' }) -join "`r`n"
        $constantModule = $book.VBProject.VBComponents.Item('modGeometryTypes').CodeModule
        $constantModule.DeleteLines(1, $constantModule.CountOfLines)
        $constantModule.AddFromString($constantBody)
    }
    $book.Save()
    $book.Close($false)
    $book = $null
    Write-Output "TEST_MODULE_UPDATED: $TestModule; production unchanged; $WorkbookPath"
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
# Excel может записать локализованный alias имени при сохранении VBA-проекта.
# Возвращаем сохраненную область печати только после закрытия COM-книги.
Restore-WorkbookPrintAreas $bookPath $printAreas
