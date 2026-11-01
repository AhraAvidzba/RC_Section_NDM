# Статическая структура малого индекса Results при полной сборке; расчётных данных не создаёт.
function Add-IndexedResultsTable {
    param([object]$Workbook,[object]$Layout)
    $item=@($Layout.Names|Where-Object Name -eq 'rngNDMElementResultIndex')
    if($item.Count -ne 1){throw 'В WorkbookLayout должен быть один якорь rngNDMElementResultIndex.'}
    $sheet=$Workbook.Worksheets.Item($item[0].Sheet)
    $sheet.Cells.Locked=$false # Статическая сборка: пользовательская защита блокирует только помеченные технические области.
    $anchor=$sheet.Range($item[0].Address)
    Add-WorkbookName $Workbook 'rngNDMElementResultIndex' $sheet $item[0].Address
    $headers=@('IndexKind','RunID','TableRows','TableColumns','FirstRow','RowCount','LoadCase','ProfileId','StateType')
    $header=$anchor.Resize(1,$headers.Count)
    $header.Locked=$true
    $Workbook.Names.Item('rngNDMElementResults').RefersToRange.Resize(1,9).Locked=$true
    $header.Font.Name='Arial';$header.Font.Size=9;$header.Font.Bold=$true
    $header.Interior.Color=14277081;$header.HorizontalAlignment=-4131;$header.VerticalAlignment=-4108
    $header.WrapText=$false;$header.Borders.LineStyle=1;$header.Borders.Color=13158600
    for($column=0;$column -lt $headers.Count;$column++){$anchor.Offset(0,$column).Value2=$headers[$column]}
    $title=$anchor.Offset(-1,0).Resize(1,$headers.Count)
    $title.Locked=$true
    $title.Merge();$title.Value2='Индекс сохранённых состояний элементов'
    $title.Font.Name='Arial';$title.Font.Size=9;$title.Font.Bold=$true
    $title.Interior.Color=14277081;$title.HorizontalAlignment=-4108;$title.VerticalAlignment=-4108
}

# Продлевает верхнюю полосу до индекса, расположенного после остальных нижних таблиц.
function Set-ResultsSnapshotTitleBand {
    param([object]$Workbook)
    $geometry=$Workbook.Names.Item('rngNDMSectionGeometry').RefersToRange
    $index=$Workbook.Names.Item('rngNDMElementResultIndex').RefersToRange
    if($geometry.Worksheet.Name -ne $index.Worksheet.Name -or $geometry.Row -lt 3){throw 'Некорректное расположение полосы Results'}
    $right=[int]$index.Column+8
    $band=$geometry.Cells.Item(1,1).Offset(-2,0).Resize(1,$right-[int]$geometry.Column+1)
    $band.Interior.Color=15652797
}
