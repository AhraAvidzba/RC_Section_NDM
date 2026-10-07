# Imports only the selected source set into a private copy of the accepted book.
# No migration, calculation, output replacement or user-process cleanup occurs here.
param(
    [string]$Directory = 'docs/regression/Performance/StorageCandidate',
    [switch]$FrozenProduction,
    [string]$SourceWorkbook = 'docs/regression/Performance/Baseline/RC_Section_NDM.xlsm',
    [switch]$TestModuleOnly,
    [string[]]$SourceModules = @(),
    [switch]$RefreshHelp,
    [switch]$SkipSmoke,
    [string]$SmokeMacro = 'modTestPerformance.RunPerformanceStorageTests'
)
$ErrorActionPreference = 'Stop'
$SourceModules = @($SourceModules | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$base = [IO.Path]::GetFullPath((Join-Path $root $SourceWorkbook))
$directoryPath = [IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Performance')) + [IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture must remain inside Performance evidence.' }
if (-not $base.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase)) { throw 'Source fixture is outside Performance evidence.' }
New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
$target = Join-Path $directoryPath 'RC_Section_NDM.xlsm'
if (Test-Path -LiteralPath $target) { throw 'Fixture exists; use a fresh evidence directory.' }
Copy-Item -LiteralPath $base -Destination $target
$printAreas = @(Get-WorkbookPrintAreas $target)
$baseHash = (Get-FileHash -LiteralPath $base -Algorithm SHA256).Hash
$excel = $null; $book = $null
function Get-ComProperty([object]$Target, [string]$Property) {
    if ($null -eq $Target) { throw "Missing COM target: $Property" }
    try { $value = $Target.GetType().InvokeMember($Property, [Reflection.BindingFlags]::GetProperty, $null, $Target, $null) }
    catch { throw "COM property $Property failed: $($_.Exception.Message)" }
    if ($null -eq $value) { throw "Missing COM property: $Property" }
    return ,$value
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $books = Get-ComProperty $excel 'Workbooks'
    $book = $books.Open($target)
    $components = Get-ComProperty (Get-ComProperty $book 'VBProject') 'VBComponents'
    $files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.cls','.bas')
    foreach ($name in $SourceModules) {
        if (@($files | Where-Object BaseName -eq $name).Count -ne 1) { throw "Source component not found uniquely: $name" }
    }
    if ($FrozenProduction -or $TestModuleOnly) { $files = @($files | Where-Object BaseName -eq 'modTestPerformance') }
    elseif ($SourceModules.Count -gt 0) { $files = @($files | Where-Object {$_.BaseName -eq 'modTestPerformance' -or $_.BaseName -in $SourceModules}) }
    foreach ($file in $files) {
        $component = $null
        for ($i = 1; $i -le $components.Count; $i++) {
            $candidate = $components.Item($i)
            if ([string]$candidate.Name -eq $file.BaseName) { $component = $candidate; break }
        }
        $type = 1; if ($file.Extension -eq '.cls') { $type = 2 }
        if ($null -eq $component) { $component = $components.Add($type); $component.Name = $file.BaseName }
        if ([int]$component.Type -ne $type) { throw "Unexpected component type: $($file.BaseName)" }
        $body = [regex]::Match([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8), '(?ms)^Option Explicit.*').Value
        if (-not $body) { throw "Missing VBA body: $($file.FullName)" }
        if ($FrozenProduction -and $file.BaseName -eq 'modTestPerformance') {
            $body = $body.Replace('#Const PERFORMANCE_CURRENT = True', '#Const PERFORMANCE_CURRENT = False')
        }
        $code = Get-ComProperty $component 'CodeModule'
        if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
        $code.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    $project = Get-ComProperty $book 'VBProject'
    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'VBA compile command is unavailable.' }
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) {
        $pane = $excel.VBE.ActiveCodePane
        $startLine=0; $startColumn=0; $endLine=0; $endColumn=0
        $pane.GetSelection([ref]$startLine,[ref]$startColumn,[ref]$endLine,[ref]$endColumn)
        throw "VBA compile failed at $($pane.CodeModule.Parent.Name):${startLine}: $($pane.CodeModule.Lines($startLine,1))"
    }
    if ($RefreshHelp) {
        Add-SettingsInstructions $book $book.Worksheets.Item('Config') $book.Worksheets.Item(2)
    }
    $book.Save()
    if ($SmokeMacro -and -not $SkipSmoke) {
        $result = [string]$excel.Run("'$($book.Name)'!$SmokeMacro")
        $result | Set-Content -LiteralPath (Join-Path $directoryPath 'Smoke.txt') -Encoding UTF8
        Write-Output $result
        if ($result -match '(?m)^FAIL:' -or $result -notmatch 'failed=0') { throw 'Smoke test failed.' }
    }
    $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    [ordered]@{BaselineSHA256=$baseHash; FrozenProduction=[bool]$FrozenProduction; ImportedComponents=@($files.BaseName); HelpRefreshed=[bool]$RefreshHelp; VBACompileCompleted=$true; SourceGitSHA=(& git -C $root rev-parse HEAD); CandidateSHA256=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directoryPath 'Preparation.json') -Encoding UTF8
    if ((Get-FileHash -LiteralPath $base -Algorithm SHA256).Hash -ne $baseHash) { throw 'Baseline changed.' }
}
finally {
    if ($null -ne $book) { try {$book.Close($false)} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) { try {$excel.Quit()} catch {Write-Warning $_}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
