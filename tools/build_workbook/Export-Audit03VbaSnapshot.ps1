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
$postAuditAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PostAudit03')) + [IO.Path]::DirectorySeparatorChar
$performanceAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
$plateauAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/PlateauSlope_2026-10-11')) + [IO.Path]::DirectorySeparatorChar
$openIssuesAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/OpenIssues_2026-10-11')) + [IO.Path]::DirectorySeparatorChar
$geometryAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryDeclarative_2026-10-12')) + [IO.Path]::DirectorySeparatorChar
$reviewAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryReview_2026-10-12')) + [IO.Path]::DirectorySeparatorChar
$endpointAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/RoundedEndpoint_2026-10-12')) + [IO.Path]::DirectorySeparatorChar
$fullAllowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryFull_2026-10-13')) + [IO.Path]::DirectorySeparatorChar
$source = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$output = [IO.Path]::GetFullPath((Join-Path $root $OutputPath))
$publishedWorkbook = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
foreach ($path in @($source, $output)) {
    # Итоговая приёмка может читать именно опубликованную книгу. Это исключение
    # только для source; dump по-прежнему создаётся в evidence, без Save/macros.
    if ($path -eq $source -and $source -eq $publishedWorkbook) {continue}
    if (-not $path.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($postAuditAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($performanceAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($plateauAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($openIssuesAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($geometryAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($reviewAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($endpointAllowed, [StringComparison]::OrdinalIgnoreCase) -and
        -not $path.StartsWith($fullAllowed, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Read-only export paths must stay in their regression evidence directory.'
    }
}
if (Test-Path -LiteralPath $output) { throw 'Export already exists; use a new evidence path.' }
$beforeHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$excel = $null
$book = $null
$text = New-Object Text.StringBuilder
$componentCount = 0
$ownedPid=0;$ownedTicks=0
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class VbaSnapshotWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $excel.EnableEvents=$false
    [uint32]$id=0;[void][VbaSnapshotWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
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
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
$afterHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
if ($beforeHash -ne $afterHash) { throw 'Verification source changed; export is not accepted.' }
[IO.File]::WriteAllText($output, $text.ToString(), (New-Object Text.UTF8Encoding($true)))
Write-Output "VBA_READ_ONLY_EXPORT: components=$componentCount; sourceSHA256=$afterHash; sourceUnchanged=True; output=$output"
