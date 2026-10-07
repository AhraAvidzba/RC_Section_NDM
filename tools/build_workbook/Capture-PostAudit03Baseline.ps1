# Сохраняет исходные артефакты и metadata перед локальными post-audit правками.
param([string]$ReportDirectory = 'docs/regression/PostAudit03/Baseline')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$target = Join-Path $root $ReportDirectory
if (Test-Path -LiteralPath (Join-Path $target 'Baseline.json')) {
    throw 'Baseline already exists; it must not be overwritten.'
}
New-Item -ItemType Directory -Path $target -Force | Out-Null
$artifacts = @()
foreach ($name in @('RC_Section_NDM.xlsm', 'VBA_All_Code.txt', 'RC_Section_NDM_execution_report.txt')) {
    $path = Join-Path $root ('workbook/output/' + $name)
    $destination = Join-Path $target $name
    Copy-Item -LiteralPath $path -Destination $destination
    $artifacts += [ordered]@{Path = ('workbook/output/' + $name); SHA256 = (Get-FileHash -LiteralPath $path).Hash; Length = (Get-Item -LiteralPath $path).Length}
}
$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
    Where-Object {$_.Extension -in '.cls', '.bas'} | Sort-Object FullName | ForEach-Object {
        [ordered]@{Path = $_.FullName.Substring($root.Length + 1); SHA256 = (Get-FileHash -LiteralPath $_.FullName).Hash}
    })
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open((Join-Path $target 'RC_Section_NDM.xlsm'), 0, $true)
    $snapshots = @()
    foreach ($sheetName in @('Config', 'Results')) {
        $sheet = $book.Worksheets.Item($sheetName)
        $range = $sheet.UsedRange
        $widths = @()
        for ($column = 1; $column -le $range.Column + $range.Columns.Count - 1; $column++) {
            $widths += [ordered]@{Column = $column; Width = $sheet.Columns.Item($column).ColumnWidth}
        }
        $snapshots += [pscustomobject]@{Sheet = $sheetName; Address = $range.Address(); Formula = $range.Formula; Widths = $widths}
    }
    $snapshots | Export-Clixml -LiteralPath (Join-Path $target 'UserSheets.clixml') -Depth 6
    $components = @()
    foreach ($component in $book.VBProject.VBComponents) {
        $components += [ordered]@{Name = $component.Name; Type = $component.Type; Lines = $component.CodeModule.CountOfLines}
    }
    $anchors = @()
    foreach ($name in $book.Names) {
        $anchors += [ordered]@{Name = $name.Name; RefersTo = $name.RefersTo}
    }
    [ordered]@{
        GitSHA = (& git -C $root rev-parse HEAD)
        TrackedStatus = @(& git -C $root status --short --untracked-files=no)
        CapturedAt = (Get-Date).ToString('o'); Artifacts = $artifacts; Sources = $sources
        SourceClassCount = @($sources | Where-Object {$_.Path -like '*.cls'}).Count
        Components = $components; Names = $anchors; SolveExecuted = $false; ImportExecuted = $false
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $target 'Baseline.json') -Encoding UTF8
    Write-Output ('BASELINE: sources=' + $sources.Count + '; components=' + $components.Count)
}
finally {
    if ($null -ne $book) {$book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
    if ($null -ne $excel) {$excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
