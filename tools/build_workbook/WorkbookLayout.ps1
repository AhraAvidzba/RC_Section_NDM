# Сохраняет расположение элементов книги отдельно от исходных данных и расчётного VBA.
# Полная сборка применяет этот снимок после создания таблиц и перед сохранением.
function Read-WorkbookLayout {
    param([string]$Path=(Join-Path $PSScriptRoot 'WorkbookLayout.json'))
    $layout=Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    if($layout.Schema -ne 1){throw 'Неизвестный формат расположения книги.'}
    return $layout
}

# Возвращает закреплённую колонку служебного списка, не создавая имён в книге.
function Get-WorkbookLayoutHelperColumn {
    param([string]$Key)
    if($null -eq $script:WorkbookLayout){return 0}
    $property=$script:WorkbookLayout.HelperColumns.PSObject.Properties[$Key]
    if($property){return [int]$property.Value}
    return 0
}

# Единое оформление заголовков типов по образцу сохранённой книги.
function Format-SectionTypeHeaders {
    param([object]$Workbook,[object]$Style)
    foreach($name in @('rngCircleGeometry','rngRectSetGeometry','rngPointedRectGeometry','rngHollowRectangleGeometry')){
        $table=$Workbook.Names.Item($name).RefersToRange;$row=[int]$table.Row-1
        $band=$table.Worksheet.Rows.Item($row)
        $band.Interior.Color=[int]$Style.Color;$band.RowHeight=[double]$Style.Height
        $cell=$table.Worksheet.Cells.Item($row,$table.Column)
        $cell.Font.Name=[string]$Style.FontName;$cell.Font.Size=[double]$Style.FontSize
        $cell.Font.Bold=$true;$cell.Font.Color=0;$cell.Font.Underline=-4142
        $cell.HorizontalAlignment=-4131;$cell.VerticalAlignment=-4108;$cell.WrapText=$false
        $text=[string]$cell.Value2;$start=$text.IndexOf('(Подробнее)')+1
        if($start -gt 0){$link=$cell.Characters($start,11);$link.Font.Color=16711680;$link.Font.Underline=2;$link.Font.Size=[double]$Style.LinkSize}
    }
}

# Читает метаданные сохранённого файла для ограниченного обхода строк и столбцов.
function Get-WorkbookLayoutXmlSheets {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($Path)
    function Read-LayoutXml([string]$Part){$reader=New-Object IO.StreamReader($zip.GetEntry($Part).Open());try{return [xml]$reader.ReadToEnd()}finally{$reader.Dispose()}}
    try{
        $wb=Read-LayoutXml 'xl/workbook.xml';$rels=Read-LayoutXml 'xl/_rels/workbook.xml.rels'
        $result=@{}
        foreach($sheet in $wb.workbook.sheets.sheet){
            $target=@($rels.Relationships.Relationship | Where-Object Id -eq $sheet.id)[0].Target
            $part=if($target.StartsWith('/')){$target.TrimStart('/')}else{'xl/'+$target}
            $result[[string]$sheet.name]=Read-LayoutXml $part
        }
        $areas=@{}
        $orderedSheets=@($wb.workbook.sheets.sheet)
        foreach($name in $wb.workbook.definedNames.definedName){
            if($name.name -ne '_xlnm.Print_Area'){continue}
            $sheetName=[string]$orderedSheets[[int]$name.localSheetId].name
            $text=[string]$name.InnerText
            $text=$text.Replace("'"+$sheetName.Replace("'","''")+"'!",'').Replace($sheetName+'!','')
            $areas[$sheetName]=$text
        }
        $result['__PrintAreas']=$areas
        return $result
    }finally{$zip.Dispose()}
}

# Снимает все имена, таблицы, геометрию объектов и размеры листов. Значения ячеек не копирует.
function Get-WorkbookLayoutSnapshot {
    param([object]$Workbook,[hashtable]$XmlSheets)
    $names=@();$sheets=@()
    foreach($name in $Workbook.Names){
        if($name.Name -match '(^|!)Print_Area$|_xlnm\.'){continue}
        $range=$null;try{$range=$name.RefersToRange}catch{}
        $record=[ordered]@{Name=[string]$name.Name;RefersTo=[string]$name.RefersTo}
        if($range){$record.Sheet=[string]$range.Worksheet.Name;$record.Address=$range.Address();$record.Rows=[int]$range.Rows.Count;$record.Columns=[int]$range.Columns.Count}
        $names+=,[pscustomobject]$record
    }
    foreach($sheet in $Workbook.Worksheets){
        $xml=$XmlSheets[[string]$sheet.Name];$columns=@();$rows=@();$objects=@();$tables=@()
        foreach($column in $xml.worksheet.cols.col){
            $first=[int]$column.min;$last=[int]$column.max;$area=$sheet.Range($sheet.Columns.Item($first),$sheet.Columns.Item($last))
            $hidden=[bool]$area.EntireColumn.Hidden
            if($hidden){$area.EntireColumn.Hidden=$false}
            $width=[double]$area.ColumnWidth
            if($hidden){$area.EntireColumn.Hidden=$true}
            $columns+=,[ordered]@{First=$first;Last=$last;Width=$width;Hidden=$hidden}
        }
        foreach($row in $xml.worksheet.sheetData.row){
            if(-not $row.HasAttribute('ht') -and -not $row.HasAttribute('hidden')){continue}
            $height=if($row.HasAttribute('ht')){[double]::Parse([string]$row.ht,[Globalization.CultureInfo]::InvariantCulture)}else{[double]$sheet.StandardHeight}
            if($height -ne [double]$sheet.StandardHeight -or $row.hidden -eq '1'){$rows+=,[ordered]@{Row=[int]$row.r;Height=$height;Hidden=($row.hidden -eq '1')}}
        }
        foreach($shape in $sheet.Shapes){
            $chartType=$null;if($shape.Type -eq 3){$chartType=[int]$sheet.ChartObjects($shape.Name).Chart.ChartType}
            $objects+=,[ordered]@{Name=[string]$shape.Name;Type=[int]$shape.Type;ChartType=$chartType;Left=[double]$shape.Left;Top=[double]$shape.Top;Width=[double]$shape.Width;Height=[double]$shape.Height;Placement=[int]$shape.Placement;Visible=[int]$shape.Visible;LockAspectRatio=[int]$shape.LockAspectRatio}
        }
        foreach($table in $sheet.ListObjects){$tables+=,[ordered]@{Name=[string]$table.Name;Address=$table.Range.Address()}}
        $sheets+=,[ordered]@{Name=[string]$sheet.Name;Visible=[int]$sheet.Visible;StandardWidth=[double]$sheet.StandardWidth;StandardHeight=[double]$sheet.StandardHeight;Columns=$columns;Rows=$rows;Objects=$objects;Tables=$tables;PrintArea=[string]$XmlSheets['__PrintAreas'][[string]$sheet.Name]}
    }
    $config=$Workbook.Worksheets.Item('Config');$system=$Workbook.Names.Item('rngSystemSettings').RefersToRange
    $caseColumn=0
    for($r=2;$r -le $system.Rows.Count;$r++){if([string]$system.Cells.Item($r,1).Value2 -eq 'Plot.LoadCase'){$formula=[string]$system.Cells.Item($r,2).Validation.Formula1;$caseColumn=[int]$config.Range($formula.Substring(1)).Column;break}}
    $loads=$Workbook.Names.Item('rngLoadCombinations').RefersToRange;$profiles=[string]$loads.Cells.Item(2,5).Validation.Formula1
    $controls=$config.Cells.Find('Контрольные точки диаграмм',[Type]::Missing,-4163,2)
    if($null -eq $controls){throw 'Не найдена таблица контрольных точек диаграмм.'}
    $controlBlock=$controls.CurrentRegion
    $circle=$Workbook.Names.Item('rngCircleGeometry').RefersToRange
    $reference=$circle.Worksheet.Cells.Item($circle.Row-1,$circle.Column)
    $sectionStyle=[ordered]@{Color=[int]$reference.Interior.Color;Height=[double]$reference.RowHeight;FontName=[string]$reference.Characters(1,1).Font.Name;FontSize=[double]$reference.Characters(1,1).Font.Size;LinkSize=9}
    return [pscustomobject][ordered]@{Schema=1;Names=$names;Sheets=$sheets;HelperColumns=[ordered]@{rngSectionRecipeTypes=[int]$Workbook.Names.Item('rngSectionRecipeTypes').RefersToRange.Column;LoadProfiles=[int]$config.Range($profiles.Substring(1)).Column;LoadCases=$caseColumn};OtherTables=@([ordered]@{Key='MaterialDiagramControls';Sheet='Config';Caption='Контрольные точки диаграмм';Address=$controlBlock.Address()});SectionHeaders=$sectionStyle}
}

# Переносит таблицы через временный лист: их исходные и конечные области могут пересекаться.
# Cut сохраняет формулы, проверки данных, форматирование и ссылки на перемещаемые ячейки.
function Move-WorkbookLayoutTables {
    param([object]$Workbook,[object]$Layout)
    $moves=New-Object 'System.Collections.Generic.List[object]'
    foreach($item in $Layout.Names){
        if(-not $item.Sheet){continue}
        $range=$null;try{$range=$Workbook.Names.Item($item.Name).RefersToRange}catch{throw ('В сборке отсутствует именованный диапазон '+$item.Name)}
        if($range.Worksheet.Name -eq $item.Sheet -and $range.Address() -eq $item.Address){continue}
        if($range.Rows.Count -ne $item.Rows -or $range.Columns.Count -ne $item.Columns){throw ('Изменилась структура '+$item.Name+'; обновите описание расположения книги.')}
        $top=[int]$range.Row;$left=[int]$range.Column;$width=[int]$range.Columns.Count;$height=[int]$range.Rows.Count
        $withTitle=($range.Worksheet.Name -eq 'Config' -and $top -gt 1)
        if($withTitle){$top--;$height++}
        $block=$range.Worksheet.Cells.Item($top,$left).Resize($height,$width)
        $destination=$Workbook.Worksheets.Item($item.Sheet).Range($item.Address)
        $targetRow=[int]$destination.Row;if($withTitle){$targetRow--}
        $moves.Add([pscustomobject]@{Sheet=$block.Worksheet.Name;Row=$top;Column=$left;Height=$height;Width=$width;Block=$block;TargetSheet=$item.Sheet;TargetRow=$targetRow;TargetColumn=[int]$destination.Column})
    }
    foreach($item in $Layout.OtherTables){
        $sheet=$Workbook.Worksheets.Item($item.Sheet);$title=$sheet.Cells.Find($item.Caption,[Type]::Missing,-4163,2)
        if($null -eq $title){throw ('В сборке отсутствует таблица '+$item.Caption)}
        $block=$title.CurrentRegion;$target=$sheet.Range($item.Address)
        if($block.Address() -eq $target.Address()){continue}
        if($block.Rows.Count -ne $target.Rows.Count -or $block.Columns.Count -ne $target.Columns.Count){throw ('Изменилась структура таблицы '+$item.Caption)}
        $moves.Add([pscustomobject]@{Sheet=$sheet.Name;Row=[int]$block.Row;Column=[int]$block.Column;Height=[int]$block.Rows.Count;Width=[int]$block.Columns.Count;Block=$block;TargetSheet=$sheet.Name;TargetRow=[int]$target.Row;TargetColumn=[int]$target.Column})
    }
    if(-not $moves.Count){return}
    $links=@()
    foreach($sheet in $Workbook.Worksheets){foreach($link in $sheet.Hyperlinks){$links+=,[pscustomobject]@{Link=$link;Target=[string]$link.SubAddress}}}
    $stage=$Workbook.Worksheets.Add();$stage.Name='__NDMLayoutBuild';$cursor=1
    try{
        foreach($move in $moves){$move | Add-Member StageRow $cursor;$move.Block.Cut($stage.Cells.Item($cursor,1)) | Out-Null;$cursor+=$move.Height+2}
        foreach($move in $moves){$stage.Cells.Item($move.StageRow,1).Resize($move.Height,$move.Width).Cut($Workbook.Worksheets.Item($move.TargetSheet).Cells.Item($move.TargetRow,$move.TargetColumn)) | Out-Null}
        foreach($record in $links){
            if($record.Target -notmatch "^'?([^']+)'?!([$]?[A-Z]+[$]?[0-9]+)$"){continue}
            $sheetName=$Matches[1];$reference=$Matches[2];$cell=$Workbook.Worksheets.Item($sheetName).Range($reference)
            foreach($move in $moves){
                if($sheetName -eq $move.Sheet -and $cell.Row -ge $move.Row -and $cell.Row -lt $move.Row+$move.Height -and $cell.Column -ge $move.Column -and $cell.Column -lt $move.Column+$move.Width){
                    $new=$Workbook.Worksheets.Item($move.TargetSheet).Cells.Item($move.TargetRow+$cell.Row-$move.Row,$move.TargetColumn+$cell.Column-$move.Column)
                    $record.Link.SubAddress="'"+$move.TargetSheet+"'!"+$new.Address($false,$false);break
                }
            }
        }
    }finally{$Workbook.Application.CutCopyMode=$false;$Workbook.Application.DisplayAlerts=$false;$stage.Delete()}
}

# Восстанавливает размеры листов и положения всех объектов после сборки содержимого.
function Apply-WorkbookLayout {
    param([object]$Workbook,[object]$Layout=$script:WorkbookLayout)
    Move-WorkbookLayoutTables $Workbook $Layout
    $previous=$null
    foreach($item in $Layout.Sheets){
        $sheet=$Workbook.Worksheets.Item($item.Name)
        if($previous){$sheet.Move([Type]::Missing,$previous)}
        $previous=$sheet;$sheet.Visible=[int]$item.Visible;$sheet.StandardWidth=[double]$item.StandardWidth
        $sheet.Columns.ColumnWidth=[double]$item.StandardWidth;$sheet.Columns.Hidden=$false
        $sheet.Rows.RowHeight=[double]$item.StandardHeight;$sheet.Rows.Hidden=$false
        foreach($column in $item.Columns){$range=$sheet.Range($sheet.Columns.Item($column.First),$sheet.Columns.Item($column.Last));$range.ColumnWidth=[double]$column.Width;$range.EntireColumn.Hidden=[bool]$column.Hidden}
        foreach($row in $item.Rows){$sheet.Rows.Item($row.Row).RowHeight=[double]$row.Height;$sheet.Rows.Item($row.Row).Hidden=[bool]$row.Hidden}
        $sheet.PageSetup.PrintArea=[string]$item.PrintArea
        foreach($object in $item.Objects){
            $shape=$null;try{$shape=$sheet.Shapes.Item($object.Name)}catch{}
            if($null -eq $shape -and $object.Type -eq 3){
                # Новый шаблон не содержит старых результатов; создаём пустой контейнер схемы.
                $chart=$sheet.ChartObjects().Add($object.Left,$object.Top,$object.Width,$object.Height);$chart.Name=[string]$object.Name;$chart.Chart.ChartType=[int]$object.ChartType;$shape=$sheet.Shapes.Item($object.Name)
            }
            if($null -eq $shape){throw ('В сборке отсутствует объект '+$item.Name+'!'+$object.Name)}
            $shape.LockAspectRatio=0;$shape.Width=[double]$object.Width;$shape.Height=[double]$object.Height
            $shape.Left=[double]$object.Left;$shape.Top=[double]$object.Top;$shape.Placement=[int]$object.Placement;$shape.Visible=[int]$object.Visible;$shape.LockAspectRatio=[int]$object.LockAspectRatio
        }
        foreach($table in $item.Tables){$sheet.ListObjects.Item($table.Name).Resize($sheet.Range($table.Address))}
    }
    foreach($item in $Layout.Names){
        $name=$Workbook.Names.Item($item.Name)
        if($item.Sheet){$range=$name.RefersToRange;if($range.Worksheet.Name -ne $item.Sheet -or $range.Address() -ne $item.Address){throw ('Не восстановлен диапазон '+$item.Name)}}
        elseif([string]$name.RefersTo -ne [string]$item.RefersTo){$name.RefersTo=[string]$item.RefersTo}
    }
    Format-SectionTypeHeaders $Workbook $Layout.SectionHeaders
}
