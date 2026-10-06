# Проверяет адресность справки в каталоге и в сохраненной Excel-книге.
# Ссылки должны вести к пояснениям конкретного параметра, не к общему дампу.
param([string]$WorkbookPath = 'docs/regression/ContourExportHelp_2026-10-06/Publication/RC_Section_NDM.xlsm')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
. (Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1')
$keys = @('GroupGapTolerance', 'RowTolerance', 'NeighborRatioLimit', 'RadiusDiameterMode', 'InteractionRadiusMode', 'RebarProfile')
$lines = New-Object 'System.Collections.Generic.List[string]'; $failed = 0; $passed = 0

# Сохраняет все независимые проверки до итогового протокола.
function Assert-Help([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; $lines.Add("OK: $Name") }
    else { $script:failed++; $lines.Add("FAIL: $Name") }
}

$entries = @{}
foreach ($group in Get-SystemSettingsCatalog) { foreach ($entry in $group.Rows) { $entries[[string]$entry[0]] = $entry } }
foreach ($key in $keys) {
    $full = "SLS.Crack.SP35.$key"; $description = @(Get-SettingInstructionLines $full $entries[$full][3]) -join "`n"
    foreach ($other in $keys) {
        if ($other -ne $key) { Assert-Help "source.$key.noUnrelated.$other" (-not $description.Contains($other)) }
    }
    Assert-Help "source.$key.concise" (($description -split "`n").Count -le 5)
}
$description = @(Get-SettingInstructionLines 'SLS.Crack.SP35.NeighborRatioLimit' $entries['SLS.Crack.SP35.NeighborRatioLimit'][3]) -join "`n"
Assert-Help 'source.neighbor.rebarAndProjection' ($description.Contains('арматурных стержней') -and $description.Contains('проекциями на НЛ'))
$source = [IO.File]::ReadAllText((Join-Path $root 'tools/build_workbook/SettingsCatalog.ps1'))
Assert-Help 'source.noSP35WildcardBundle' (-not $source.Contains('"SLS.Crack.SP35.*"'))
Assert-Help 'source.sourceLabels' ($source.Contains('Метки источников') -and $source.Contains('[METHOD] - пояснения') -and $source.Contains('[NDM] - принятые'))
$excel = $null; $book = $null; $path = Join-Path $root $WorkbookPath
$hash = (Get-FileHash -LiteralPath $path).Hash
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($path, 0, $true)
    $settings = $book.Names.Item('rngSystemSettings').RefersToRange
    $guide = $book.Worksheets.Item('Справка'); $data = $guide.UsedRange.Value2
    foreach ($key in $keys) {
        $full = "SLS.Crack.SP35.$key"; $row = 0
        for ($r = 1; $r -le $settings.Rows.Count; $r++) { if ([string]$settings.Cells.Item($r,1).Value2 -eq $full) { $row = $r; break } }
        if ($row -eq 0) { throw "Missing key: $full" }
        $link = [string]$settings.Cells.Item($row,5).Hyperlinks.Item(1).SubAddress
        if ($link -notmatch '!\$?A\$?(\d+)$') { throw "Unexpected Help link: $link" }
        $start = [int]$Matches[1]; $text = ''
        for ($r = $start + 1; $r -le $data.GetLength(0); $r++) {
            if ([string]$data[$r,1] -notmatch '^\d+$') { break }
            $text += [string]$data[$r,2] + "`n"
        }
        Assert-Help "saved.$key.linkTarget" ([string]$data[$start,1] -like "*$full*")
        foreach ($other in $keys) { if ($other -ne $key) { Assert-Help "saved.$key.noUnrelated.$other" (-not $text.Contains($other)) } }
        Assert-Help "saved.$key.textMatchesCatalog" ($text.Trim() -ceq (@(Get-SettingInstructionLines $full $entries[$full][3]) -join "`n").Trim())
    }
    $joined = ($data | ForEach-Object { [string]$_ }) -join "`n"
    Assert-Help 'saved.sourceLabels' ($joined.Contains('Метки источников') -and $joined.Contains('[METHOD] - пояснения'))
    Assert-Help 'saved.additiveContract' ($joined.Contains('без поиска совпадений и без автоматической очистки'))
    Assert-Help 'saved.colors' ($joined.Contains('ACI 30') -and $joined.Contains('ACI 4') -and $joined.Contains('ACI 31'))
    Assert-Help 'saved.exactDimensions' ($joined.Contains('габаритные размеры импортированного сечения берутся по точному наружному контуру'))
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Assert-Help 'saved.readOnly' ((Get-FileHash -LiteralPath $path).Hash -eq $hash)
$lines.Add("TOTAL_HELP: passed=$passed; failed=$failed")
$lines | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'HelpChecks.txt') -Encoding UTF8
Write-Output $lines[$lines.Count-1]
if ($failed -gt 0) { throw 'Help checks failed.' }
