# Сравнивает весь сохраненный текст и объединения справки clean/update книг.
# Читает OpenXML без запуска Excel; не меняет исходники и не подтверждает
# нормативную правильность текста либо его пиксельную компоновку.
param(
    [Parameter(Mandatory=$true)][string]$CleanWorkbook,
    [Parameter(Mandatory=$true)][string]$UpdatedWorkbook,
    [Parameter(Mandatory=$true)][string]$ReportPath
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$lines = [Collections.Generic.List[string]]::new()
$failed = 0

# Возвращает обязательную часть книги как XML. Отсутствующий лист или часть
# считаются ошибкой проверки, а не эквивалентной пустой справкой.
function Read-Part($archive, [string]$name) {
    $entry = $archive.GetEntry($name)
    if (-not $entry) { throw "В книге отсутствует $name." }
    $reader = [IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

# Раскрывает rich/shared text и сохраняет точный адрес каждой непустой ячейки.
# Пустые style-only ячейки не считаются содержанием; объединения проверяются
# отдельно, поскольку старое Merge может поглотить новый абзац при update.
function Read-Help([string]$path) {
    $archive = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $workbook = Read-Part $archive 'xl/workbook.xml'
        $relations = Read-Part $archive 'xl/_rels/workbook.xml.rels'
        $help = @($workbook.workbook.sheets.sheet | Where-Object name -eq 'Справка')
        if ($help.Count -ne 1) { throw 'Ожидался один лист Справка.' }
        $id = $help[0].GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relation = @($relations.Relationships.Relationship | Where-Object Id -eq $id)
        if ($relation.Count -ne 1) { throw 'Не найдена однозначная часть листа Справка.' }
        $target = [string]$relation[0].Target
        $part = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }
        $sheet = Read-Part $archive $part
        $strings = @()
        if ($archive.GetEntry('xl/sharedStrings.xml')) {
            $shared = Read-Part $archive 'xl/sharedStrings.xml'
            $strings = @($shared.sst.si | ForEach-Object {
                @($_.SelectNodes('.//*[local-name()="t"]') | ForEach-Object InnerText) -join ''
            })
        }
        $cells = @{}
        foreach ($cell in $sheet.SelectNodes('//*[local-name()="sheetData"]/*[local-name()="row"]/*[local-name()="c"]')) {
            $valueNode = $cell.SelectSingleNode('./*[local-name()="v"]')
            $text = ''
            if ([string]$cell.t -eq 's') { $text = $strings[[int]$valueNode.InnerText] }
            elseif ([string]$cell.t -eq 'inlineStr') {
                $text = @($cell.SelectNodes('.//*[local-name()="t"]') | ForEach-Object InnerText) -join ''
            } elseif ($valueNode) { $text = [string]$valueNode.InnerText }
            $formula = $cell.SelectSingleNode('./*[local-name()="f"]')
            if ($formula) { $text = 'FORMULA=' + $formula.InnerText + '; VALUE=' + $text }
            if ($text.Length -gt 0) { $cells[[string]$cell.r] = $text }
        }
        $merges = @($sheet.SelectNodes('//*[local-name()="mergeCells"]/*[local-name()="mergeCell"]') |
            ForEach-Object { [string]$_.ref } | Sort-Object)
        return @{Cells=$cells; Merges=$merges}
    } finally { $archive.Dispose() }
}

$cleanPath = (Resolve-Path -LiteralPath $CleanWorkbook).Path
$updatedPath = (Resolve-Path -LiteralPath $UpdatedWorkbook).Path
$cleanHash = (Get-FileHash -LiteralPath $cleanPath -Algorithm SHA256).Hash
$updatedHash = (Get-FileHash -LiteralPath $updatedPath -Algorithm SHA256).Hash
$clean = Read-Help $cleanPath
$updated = Read-Help $updatedPath
$addresses = @($clean.Cells.Keys + $updated.Cells.Keys | Sort-Object -Unique)
foreach ($address in $addresses) {
    if ([string]$clean.Cells[$address] -cne [string]$updated.Cells[$address]) {
        $failed++
        $lines.Add("HELP_DIFFERENCE|$address|clean=$($clean.Cells[$address])|updated=$($updated.Cells[$address])")
    }
}
$mergeEqual = ($clean.Merges -join '|') -ceq ($updated.Merges -join '|')
if (-not $mergeEqual) {
    $failed++
    $lines.Add("HELP_MERGE_DIFFERENCE|clean=$($clean.Merges -join ',')|updated=$($updated.Merges -join ',')")
}
$unchanged = $cleanHash -eq (Get-FileHash -LiteralPath $cleanPath -Algorithm SHA256).Hash -and
    $updatedHash -eq (Get-FileHash -LiteralPath $updatedPath -Algorithm SHA256).Hash
if (-not $unchanged) { $failed++ }
$lines.Add("HELP_COMPARE|cleanCells=$($clean.Cells.Count)|updatedCells=$($updated.Cells.Count)|mergesEqual=$mergeEqual|sourceUnchanged=$unchanged")
$lines.Add("TOTAL_AUDIT03_HELP_COMPARE: failed=$failed")
$lines | Set-Content -LiteralPath $ReportPath -Encoding UTF8
$lines | Select-Object -Last 2 | Write-Output
if ($failed -gt 0) { exit 1 }
