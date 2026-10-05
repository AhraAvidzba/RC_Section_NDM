param([string]$OriginalPath, [string]$CandidatePath, [string]$ReportPath, [string]$CodeExportPath)

$ErrorActionPreference = 'Stop'
$excel = $null
$original = $null
$candidate = $null
$lines = New-Object System.Collections.Generic.List[string]
$code = New-Object Text.StringBuilder
$allowed = @('modWorkbookMessages', 'modWorkbookCalculation', 'modAutoCADStressExport')
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $original = $excel.Workbooks.Open([IO.Path]::GetFullPath($OriginalPath), 0, $true)
    $candidate = $excel.Workbooks.Open([IO.Path]::GetFullPath($CandidatePath), 0, $true)
    $lines.Add('Component counts: original=' + $original.VBProject.VBComponents.Count + '; candidate=' + $candidate.VBProject.VBComponents.Count)
    if ($candidate.VBProject.VBComponents.Count -ne ($original.VBProject.VBComponents.Count + 1)) { throw $lines[0] }
    $baseline = @{}
    foreach ($component in $original.VBProject.VBComponents) {
        $text = ''
        if ($component.CodeModule.CountOfLines -gt 0) { $text = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines).Trim() }
        $baseline[$component.Name] = $text
    }
    [void]$code.AppendLine('VBA PROJECT EXPORT').AppendLine('Workbook: RC_Section_NDM.xlsm')
    [void]$code.AppendLine('Date: ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach ($component in $candidate.VBProject.VBComponents) {
        $text = ''
        if ($component.CodeModule.CountOfLines -gt 0) { $text = $component.CodeModule.Lines(1, $component.CodeModule.CountOfLines).Trim() }
        if ($component.Name -notin $allowed) {
            if (-not $baseline.ContainsKey($component.Name)) { throw "Unexpected new component: $($component.Name)" }
            if (-not [string]::Equals($text, $baseline[$component.Name], [StringComparison]::OrdinalIgnoreCase)) {
                throw "Unrelated component changed: $($component.Name)"
            }
        }
        if ($component.Name -in @('modWorkbookCalculation', 'modAutoCADStressExport')) {
            if ($text -match '\bMsgBox\b') { throw "ANSI MsgBox remains: $($component.Name)" }
        }
        $lines.Add('PASS: component ' + $component.Name)
        [void]$code.AppendLine('=' * 100).AppendLine('COMPONENT: ' + $component.Name)
        [void]$code.AppendLine('TYPE: ' + $component.Type).AppendLine('LINES: ' + $component.CodeModule.CountOfLines).AppendLine('=' * 100)
        [void]$code.AppendLine($text).AppendLine()
    }
    foreach ($name in $allowed) {
        $path = Join-Path 'src/Excel' ($name + '.bas')
        $sourceLines = [IO.File]::ReadAllLines([IO.Path]::GetFullPath($path), [Text.Encoding]::UTF8) | Where-Object { $_ -notmatch '^\uFEFF?Attribute\s+VB_' }
        $sourceText = ($sourceLines -join "`r`n").Trim()
        $module = $candidate.VBProject.VBComponents.Item($name).CodeModule
        $actual = $module.Lines(1, $module.CountOfLines).Trim()
        # VBE removes redundant Double suffixes from fractional literals during import.
        $sourceText = [regex]::Replace($sourceText, '(?<![\w])([0-9]*\.[0-9]+)#', '$1')
        $actual = [regex]::Replace($actual, '(?<![\w])([0-9]*\.[0-9]+)#', '$1')
        if (-not [string]::Equals($sourceText, $actual, [StringComparison]::OrdinalIgnoreCase)) { throw "Source mismatch: $name" }
        $lines.Add('PASS: source matches ' + $name)
    }
    $lines.Add('PASS: 114 VBA components; only three intended modules changed')
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($CodeExportPath), $code.ToString(), (New-Object Text.UTF8Encoding($true)))
}
finally {
    if ($candidate) { $candidate.Close($false) }
    if ($original) { $original.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [IO.File]::WriteAllLines([IO.Path]::GetFullPath($ReportPath), $lines, (New-Object Text.UTF8Encoding($true)))
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
'VBA: ' + $lines.Count + ' checks passed'
