# Создает отдельную отрицательную книгу для проверки чувствительности тестов.
# Подмена существует только в ее VBE: исходники и исходная книга неизменны.
# Пока проверяется потеря передачи Solver.Method; новые мутации требуют
# отдельного точного контракта, а не произвольного replace по всему проекту.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$OutputWorkbook
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath (Join-Path $root $SourceWorkbook)).Path
$destination = [IO.Path]::GetFullPath((Join-Path $root $OutputWorkbook))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $destination.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Мутация разрешена только в изолированной Audit03-книге.'
}
if (Test-Path -LiteralPath $destination) { throw 'Отрицательная fixture уже существует; она не перезаписывается.' }
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
Copy-Item -LiteralPath $source -Destination $destination
$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($destination)
    $module = $book.VBProject.VBComponents.Item('CSectionSolver').CodeModule
    $body = $module.Lines(1, $module.CountOfLines)
    $needle = 'mSolverMethod = settings.GetRawString("Solver.Method", vbNullString)'
    if (($body.Split(@($needle), [StringSplitOptions]::None).Count - 1) -ne 1) {
        throw 'Точный call site Solver.Method не найден или неоднозначен.'
    }
    $body = $body.Replace($needle, 'mSolverMethod = "Newton"')
    $module.DeleteLines(1, $module.CountOfLines)
    $module.AddFromString($body)
    $book.Save()
    $book.Close($false)
    $book = $null
    if ((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne $sourceHash) {
        throw 'Исходная книга изменилась при создании отрицательного fixture.'
    }
    Write-Output "SETTINGS_MUTATION: Solver.Method disconnected; sourceSHA=$sourceHash; sourceUnchanged=True; fixture=$OutputWorkbook"
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null
    }
}
