# Заменяет только точные согласованные пояснения, включая длинные строки и повторы.
function Update-StandaloneHelpText($Workbook) {
    $mapping=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'StandaloneHelpText.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $replacements=New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::Ordinal)
    foreach ($item in $mapping.PSObject.Properties) {$replacements.Add($item.Name,[string]$item.Value)}
    $changes=@()
    foreach ($sheet in @($Workbook.Worksheets.Item('Справка'),$Workbook.Worksheets.Item('Config'))) {
        $used=$sheet.UsedRange;$values=$used.Value2
        $rows=[int]$used.Rows.Count;$columns=[int]$used.Columns.Count
        for ($r=1;$r -le $rows;$r++) {
            for ($c=1;$c -le $columns;$c++) {
                $value=$values[$r,$c]
                if ($value -is [string] -and $replacements.ContainsKey($value)) {
                    $cell=$used.Cells.Item($r,$c)
                    $changes+=,[ordered]@{Sheet=$sheet.Name;Address=$cell.Address();Before=$value;After=$replacements[$value]}
                    $cell.Value2=[string]$replacements[$value]
                }
            }
        }
    }
    return $changes
}
