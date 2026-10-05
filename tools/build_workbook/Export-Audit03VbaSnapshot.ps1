# Читает фактический VBA из отдельной проверяемой книги без импорта исходников,
# запуска макросов или сохранения книги. Экспорт служит доказательством
# source/book equality; операция не исправляет проверяемый артефакт.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$source = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputPath))
foreach ($path in @($source, $output)) {
    if (-not $path.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Read-only export paths must stay in docs/regression/Audit03.'
    }
}
if (Test-Path -LiteralPath $output) { throw 'Export already exists; use a new evidence path.' }
$beforeHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$excel = $null
$book = $null
$text = New-Object Text.StringBuilder
$componentCount = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($source, 0, $true)
    if (-not $book.ReadOnly) { throw 'Verification workbook must be read-only.' }
    [void]$text.AppendLine('VBA PROJECT EXPORT')
    [void]$text.AppendLine('Workbook: ' + $book.Name)
    [void]$text.AppendLine('Date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach ($component in $book.VBProject.VBComponents) {
        $componentCount++
        [void]$text.AppendLine('=' * 100)
        [void]$text.AppendLine('COMPONENT: ' + $component.Name)
        [void]$text.AppendLine('TYPE: ' + $component.Type)
        $lineCount = $component.CodeModule.CountOfLines
        [void]$text.AppendLine('LINES: ' + $lineCount)
        [void]$text.AppendLine('=' * 100)
        if ($lineCount -gt 0) { [void]$text.AppendLine($component.CodeModule.Lines(1, $lineCount)) }
        [void]$text.AppendLine()
    }
    $book.Close($false)
    $book = $null
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
$afterHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if ($beforeHash -ne $afterHash) { throw 'Verification source changed; export is not accepted.' }
[IO.File]::WriteAllText($output, $text.ToString(), (New-Object Text.UTF8Encoding($true)))
Write-Output "VBA_READ_ONLY_EXPORT: components=$componentCount; sourceSHA256=$afterHash; sourceUnchanged=True; output=$output"
