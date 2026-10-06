# Обновляет код и справку в собственной копии, сохраняя пользовательские
# настройки, расчетный снимок, списки и ширины. НДС здесь не пересчитывается.
param(
    [string]$ReportDirectory = '',
    [string[]]$UpdatedCommentKeys = @('SLS.Crack.SP35.NeighborRatioLimit'),
    [switch]$HelpOnly
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')

# Читает бинарный проект для диагностики служебных изменений Excel при Save.
function Get-VbaProjectHash([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($Path); $stream = $null; $sha = $null
    try {
        $entry = $zip.GetEntry('xl/vbaProject.bin')
        if ($null -eq $entry) { throw 'Workbook has no VBA project.' }
        $stream = $entry.Open(); $sha = [Security.Cryptography.SHA256]::Create()
        return [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '')
    } finally {
        if ($null -ne $sha) { $sha.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
        $zip.Dispose()
    }
}

# Сравнивает точные тексты и состав всех модулей, включая модули листов.
# Бинарный VBA-кеш Excel может изменяться даже без правок исходного кода.
function Get-VbaSourceHash([object]$Book) {
    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($component in $Book.VBProject.VBComponents | Sort-Object Name) {
        $module = $component.CodeModule; $body = ''
        if ($module.CountOfLines -gt 0) { $body = [string]$module.Lines(1, $module.CountOfLines) }
        $parts.Add("$($component.Name)|$($component.Type)|$($module.CountOfLines)")
        $parts.Add($body)
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))).Replace('-', '') }
    finally { $sha.Dispose() }
}
$source = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$directory = Join-Path $PSScriptRoot 'Publication'
if ($ReportDirectory) { $directory = Join-Path $root $ReportDirectory }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$target = Join-Path $directory 'RC_Section_NDM.xlsm'
$sourceHash = (Get-FileHash -LiteralPath $source).Hash
$sourceVbaHash = ''; $preparedVbaHash = ''
$sourceVbaSourceHash = ''; $preparedVbaSourceHash = ''
if ($HelpOnly) { $sourceVbaHash = Get-VbaProjectHash $source }
Copy-Item -LiteralPath $source -Destination (Join-Path $directory 'BeforePatch.xlsm') -Force
Copy-Item -LiteralPath $source -Destination $target -Force
$printAreas = @(Get-WorkbookPrintAreas $target)
$snapshots = @{}; $validations = @{}; $configWidths = @(); $resultWidths = @()
$resultsSnapshot = $null; $resultsAddress = ''; $commentRows = @{}
$excel = $null; $book = $null

# Сравнивает сохраненные исходные данные и Results, разрешая указанные комментарии.
function Assert-Preserved([object]$Book) {
    $cells = 0
    foreach ($name in $snapshots.Keys) {
        $range = $Book.Names.Item($name).RefersToRange
        $actual = $range.Formula; $expected = $snapshots[$name]
        if ($actual.GetLength(0) -ne $expected.GetLength(0) -or $actual.GetLength(1) -ne $expected.GetLength(1)) { throw "Range resized: $name" }
        for ($r = 1; $r -le $expected.GetLength(0); $r++) {
            for ($c = 1; $c -le $expected.GetLength(1); $c++) {
                if ($name -eq 'rngSystemSettings' -and $commentRows.ContainsKey($r) -and $c -eq 4) { continue }
                if ([string]$actual[$r,$c] -cne [string]$expected[$r,$c]) { throw "Input changed: $name/$r/$c" }
                $key = "$name/$r/$c"
                if ($validations.ContainsKey($key) -and [string]$range.Cells.Item($r,$c).Validation.Formula1 -cne $validations[$key]) { throw "List changed: $key" }
                $cells++
            }
        }
    }
    for ($c = 1; $c -le $configWidths.Count; $c++) {
        if ($Book.Worksheets.Item('Config').Columns.Item($c).ColumnWidth -ne $configWidths[$c-1]) { throw "Config width changed: $c" }
    }
    for ($c = 1; $c -le $resultWidths.Count; $c++) {
        if ($Book.Worksheets.Item('Results').Columns.Item($c).ColumnWidth -ne $resultWidths[$c-1]) { throw "Results width changed: $c" }
    }
    $actual = $Book.Worksheets.Item('Results').Range($resultsAddress).Formula
    for ($r = 1; $r -le $resultsSnapshot.GetLength(0); $r++) {
        for ($c = 1; $c -le $resultsSnapshot.GetLength(1); $c++) {
            if ([string]$actual[$r,$c] -cne [string]$resultsSnapshot[$r,$c]) { throw "Results changed: $r/$c" }
        }
    }
    return $cells
}

# Выгружает именно сохраненный VBE, включая встроенные тесты.
function Export-Vba([object]$Book, [string]$Path) {
    $writer = New-Object IO.StreamWriter($Path, $false, (New-Object Text.UTF8Encoding($true)))
    try {
        $writer.WriteLine('VBA PROJECT EXPORT'); $writer.WriteLine("Workbook: $($Book.Name)")
        $writer.WriteLine("Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        foreach ($component in $Book.VBProject.VBComponents) {
            $module = $component.CodeModule
            $writer.WriteLine(('=' * 100)); $writer.WriteLine("COMPONENT: $($component.Name)")
            $writer.WriteLine("TYPE: $($component.Type)"); $writer.WriteLine("LINES: $($module.CountOfLines)")
            $writer.WriteLine(('=' * 100))
            if ($module.CountOfLines) {
                foreach ($line in $module.Lines(1, $module.CountOfLines) -split "`r?`n") { $writer.WriteLine($line.TrimEnd()) }
            }
        }
    } finally { $writer.Dispose() }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($target)
    if ($HelpOnly) { $sourceVbaSourceHash = Get-VbaSourceHash $book }
    foreach ($name in (Get-ConfigNamedRangeNames)) {
        $range = $book.Names.Item($name).RefersToRange; $snapshots[$name] = $range.Formula
        for ($r = 1; $r -le $range.Rows.Count; $r++) {
            for ($c = 1; $c -le $range.Columns.Count; $c++) {
                try { if ($range.Cells.Item($r,$c).Validation.Type -eq 3) { $validations["$name/$r/$c"] = [string]$range.Cells.Item($r,$c).Validation.Formula1 } } catch { }
            }
        }
    }
    $config = $book.Worksheets.Item('Config'); $results = $book.Worksheets.Item('Results')
    for ($c = 1; $c -le $config.UsedRange.Columns.Count; $c++) { $configWidths += $config.Columns.Item($c).ColumnWidth }
    for ($c = 1; $c -le $results.UsedRange.Columns.Count; $c++) { $resultWidths += $results.Columns.Item($c).ColumnWidth }
    $resultsAddress = $results.UsedRange.Address(); $resultsSnapshot = $results.UsedRange.Formula
    if (-not $HelpOnly) {
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas', '.cls') {
            $component = $null
            try { $component = $book.VBProject.VBComponents.Item($file.BaseName) } catch { }
            if ($null -ne $component) { $book.VBProject.VBComponents.Remove($component) }
            $body = [regex]::Match([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8), '(?ms)^Option Explicit.*').Value
            if ([string]::IsNullOrWhiteSpace($body)) { throw "Missing Option Explicit: $file" }
            $type = 1; if ($file.Extension -eq '.cls') { $type = 2 }
            $component = $book.VBProject.VBComponents.Add($type); $component.Name = $file.BaseName
            $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
        }
    }
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    for ($r = 1; $r -le $settings.Rows.Count; $r++) {
        if ($UpdatedCommentKeys -contains [string]$settings.Cells.Item($r,1).Value2) { $commentRows[$r] = [string]$settings.Cells.Item($r,1).Value2 }
    }
    if ($commentRows.Count -ne $UpdatedCommentKeys.Count) { throw 'Requested comment setting not found.' }
    foreach ($r in $commentRows.Keys) {
        foreach ($group in (Get-SystemSettingsCatalog)) {
            foreach ($entry in $group.Rows) {
                if ($entry[0] -eq $commentRows[$r]) { $settings.Cells.Item($r,4).Value2 = [string]$entry[3] }
            }
        }
    }
    Add-SettingsInstructions $book $config ($book.Worksheets.Item('Справка'))
    for ($c = 1; $c -le $configWidths.Count; $c++) { $config.Columns.Item($c).ColumnWidth = $configWidths[$c-1] }
    $preserved = Assert-Preserved $book
    $book.Save(); $book.Close($false); $book = $null
    Restore-WorkbookPrintAreas $target $printAreas
    $book = $excel.Workbooks.Open($target, 0, $true)
    $preserved = Assert-Preserved $book
    if ($HelpOnly) {
        $preparedVbaHash = Get-VbaProjectHash $target
        $preparedVbaSourceHash = Get-VbaSourceHash $book
        if ($preparedVbaSourceHash -ne $sourceVbaSourceHash) { throw 'Help-only edit changed VBA source or module list.' }
    } else { Export-Vba $book (Join-Path $directory 'VBA_All_Code.txt') }
    if ((Get-FileHash -LiteralPath $source).Hash -ne $sourceHash) { throw 'User workbook changed; publication must rebase.' }
    [pscustomobject]@{
        SourceSHA256=$sourceHash; PreparedSHA256=(Get-FileHash -LiteralPath $target).Hash; SourceUnchanged=$true
        ConfigCellsPreserved=$preserved; ConfigCommentsUpdated=$commentRows.Count; ConfigSettingsAdded=0
        ConfigValidationListsPreserved=$validations.Count; ConfigWidthsPreserved=$configWidths.Count
        ResultsAddress=$resultsAddress; ResultsCellsPreserved=$resultsSnapshot.Length; ResultsWidthsPreserved=$resultWidths.Count
        StateSolveExecuted=$false
        HelpOnly=$HelpOnly.IsPresent; VBAProjectUnchanged=($HelpOnly -and $preparedVbaHash -eq $sourceVbaHash)
        SourceVBAProjectSHA256=$sourceVbaHash; PreparedVBAProjectSHA256=$preparedVbaHash
        VBASourceUnchanged=($HelpOnly -and $preparedVbaSourceHash -eq $sourceVbaSourceHash)
        SourceVBASourceSHA256=$sourceVbaSourceHash; PreparedVBASourceSHA256=$preparedVbaSourceHash
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Manifest.json') -Encoding UTF8
    Write-Output "PATCH_OK: Config=$preserved; Results=$($resultsSnapshot.Length); widths=$($resultWidths.Count)"
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
