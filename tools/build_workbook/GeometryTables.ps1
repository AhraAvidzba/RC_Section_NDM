# Общая структура таблиц геометрии. Сборщик не меняет ширины столбцов и высоты строк.
# Схема хранит исходные значения нового шаблона; действующие данные обновляются отдельно.
function Get-GeometrySettingsCatalog {
    $catalog=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'GeometryTables.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach($definition in $catalog){Write-Output $definition}
}

# Находит блоки по строкам key, без привязки к форме сечения и русским подписям.
function Get-GeometryTableBlocks {
    param([object]$Range)
    $data=$Range.Value2
    $rowCount=$data.GetLength(0);$columnCount=$data.GetLength(1)
    $blocks=New-Object 'System.Collections.Generic.List[object]'
    for($r=3;$r -le $rowCount;$r++){
        if([string]$data[$r,1] -ine 'key'){continue}
        $columns=@{}
        for($c=1;$c -le $columnCount;$c++){
            $key=([string]$data[$r,$c]).Trim()
            if($key){if($columns.ContainsKey($key)){throw "Повторный ключ столбца $key."};$columns[$key]=$c}
        }
        $titleRow=$r-2
        $blockKey=[string]$data[$titleRow,1]
        $blocks.Add([pscustomobject]@{Key=$blockKey;TitleRow=$titleRow;HeaderRow=($r-1);KeyRow=$r;FirstRow=($r+1);LastRow=$rowCount;Columns=$columns})
    }
    for($i=0;$i -lt $blocks.Count-1;$i++){$blocks[$i].LastRow=$blocks[$i+1].TitleRow-2}
    return $blocks.ToArray()
}

# Находит поле таблицы по тем же ключам, которыми пользуется VBA-рецепт.
function Get-GeometryTableCell {
    param([object]$Range,[string]$Block,[string]$Row,[string]$Column='val')
    $item=@(Get-GeometryTableBlocks $Range | Where-Object Key -eq $Block)
    $data=$Range.Value2
    if($item.Count -ne 1 -or -not $item[0].Columns.ContainsKey($Column)){throw "Не найдено поле $Block/$Row/$Column."}
    for($r=$item[0].FirstRow;$r -le $item[0].LastRow;$r++){
        if([string]$data[$r,1] -ieq $Row){return $Range.Cells.Item($r,[int]$item[0].Columns[$Column])}
    }
    throw "Не найдена строка $Block/$Row."
}

# Записывает двумерный массив отдельным типизированным вызовом COM.
# Это сохраняет массивный setter после записи строк в том же процессе PowerShell.
function Set-GeometryTableMatrix {
    param([object]$Target,[object[,]]$Matrix)
    $Target.Value2=$Matrix
}

# Создаёт один именованный диапазон из блоков каталога. Комментарий занимает одну ячейку.
function Add-GeometrySettingsTable {
    param([object]$Workbook,[object]$Sheet,[object]$Definition,[int]$HeaderRow,[int]$StartColumn)
    $count=0
    foreach($block in $Definition.Blocks){$count+=3+$block.Rows.Count+1;if($block.Units){$count++}}
    $count--
    $range=$Sheet.Cells.Item($HeaderRow,$StartColumn).Resize($count,11)
    $range.UnMerge() | Out-Null
    $range.Clear() | Out-Null
    $range.FormatConditions.Delete()
    $data=New-Object 'object[,]' $count,11
    $r=0
    foreach($block in $Definition.Blocks){
        $data[$r,0]=[string]$block.Key;$data[$r,1]=[string]$block.Title;$r++
        foreach($values in @($block.Headers,$block.Keys)){
            for($c=0;$c -lt $values.Count;$c++){$data[$r,$c]=$values[$c]};$r++
        }
        if($block.Units){for($c=0;$c -lt $block.Units.Count;$c++){$data[$r,$c]=$block.Units[$c]};$r++}
        foreach($values in $block.Rows){
            for($c=0;$c -lt $values.Count;$c++){$data[$r,$c]=$values[$c]};$r++
        }
        $r++
    }
    Set-GeometryTableMatrix $range $data
    Set-WorkbookNameByBounds $Workbook $Definition.RangeName $Sheet $HeaderRow $StartColumn ($HeaderRow+$count-1) ($StartColumn+10)
    $title=$Sheet.Cells.Item($HeaderRow-1,$StartColumn).Resize(1,11)
    $title.UnMerge() | Out-Null;$title.ClearContents() | Out-Null;$title.Merge() | Out-Null
    $title.Cells.Item(1,1).Value2=[string]$Definition.Title
    $title.Font.Bold=$true;$title.Interior.Color=15921906;$title.HorizontalAlignment=-4131
    Format-GeometrySettingsTable $Workbook $Definition.RangeName $range
    Add-GeometryTableValidation $range
    return $count
}

# Оформляет весь диапазон, включая разделительные строки и свободные ячейки справа.
function Format-GeometrySettingsTable {
    param([object]$Workbook,[string]$Name,[object]$Range)
    $Range.WrapText=$false;$Range.HorizontalAlignment=-4108;$Range.VerticalAlignment=-4108
    $Range.Font.Name='Arial';$Range.Font.Size=9;$Range.Font.Color=0;$Range.Font.Bold=$false;$Range.Font.Italic=$false
    $Range.Interior.Color=16777215
    $Range.FormatConditions.Delete()
    foreach($block in @(Get-GeometryTableBlocks $Range)){
        $width=[int]$block.Columns['note']
        $title=$Range.Cells.Item($block.TitleRow,2).Resize(1,$width-1)
        $title.Merge() | Out-Null;$title.HorizontalAlignment=-4131
        $heading=$Range.Cells.Item($block.TitleRow,1).Resize(1,$width)
        $heading.Font.Bold=$true;$heading.Interior.Color=15921906
        $headers=$Range.Cells.Item($block.HeaderRow,1).Resize(1,$width)
        $headers.Font.Bold=$true;$headers.Interior.Color=14277081
        $keys=$Range.Cells.Item($block.KeyRow,1).Resize(1,$width)
        $keys.Font.Italic=$true;$keys.Font.Color=6710886;$keys.Interior.Color=15921906
        for($r=$block.FirstRow;$r -le $block.LastRow;$r++){
            $rowKey=[string]$Range.Cells.Item($r,1).Value2
            if(-not $rowKey){continue}
            if($rowKey -eq 'Units'){
                $Range.Cells.Item($r,1).Resize(1,$width).Font.Color=6710886
                foreach($c in $block.Columns.Values){$cell=$Range.Cells.Item($r,[int]$c);if([string]$cell.Value2 -eq 'мм'){Set-InputUnitCell $cell 'мм'}}
            }elseif($block.Columns.ContainsKey('unit')){
                $cell=$Range.Cells.Item($r,[int]$block.Columns['unit']);if([string]$cell.Value2 -eq 'мм'){Set-InputUnitCell $cell 'мм'}
            }
        }
        $Range.Cells.Item($block.HeaderRow,$width).Resize($block.LastRow-$block.HeaderRow+1,1).HorizontalAlignment=-4152
    }
    $Range.Borders.Weight=2;$Range.Borders.Color=0;$Range.Borders.LineStyle=-4115
}

# Каждое поле выбора имеет собственный именованный источник, в том числе противоположные стороны.
function Add-GeometryTableValidation {
    param([object]$Range)
    foreach($block in @(Get-GeometryTableBlocks $Range)){
        for($r=$block.FirstRow;$r -le $block.LastRow;$r++){
            $rowKey=[string]$Range.Cells.Item($r,1).Value2
            if(-not $rowKey -or $rowKey -eq 'Units'){continue}
            foreach($key in @($block.Columns.Keys)){
                if($key -in @('key','name','unit','note')){continue}
                $cell=$Range.Cells.Item($r,[int]$block.Columns[$key]);$options=@()
                if($key -like 'bind*'){$options=@('EachBar','EverySecondBar')}
                elseif($key -like 'pos*'){$options=@('Stacked','SideBySide')}
                elseif($block.Key -eq 'Geom' -and $rowKey -eq 'Shape'){$options=@('Rectangle','LSection','TwoRectangles')}
                elseif($block.Key -eq 'Detail' -and $rowKey -eq 'Type'){$options=@('Simple','Tapered')}
                $cell.Validation.Delete()
                if($options.Count){
                    $tableName=(([string]$Range.Name.Name) -split '!')[-1]
                    if(-not $tableName -or $tableName.Length -gt 100){throw "Не удалось определить короткое имя таблицы: $tableName"}
                    $name=($tableName+'_'+$block.Key+'_'+$rowKey+'_'+$key+'_Choices') -replace '[^A-Za-z0-9_]','_'
                    $list=Get-ConfigValidationListAddress $Range.Worksheet $options $name
                    $cell.Validation.Add(3,1,1,$list);$cell.Validation.IgnoreBlank=$false;$cell.Validation.InCellDropdown=$true
                }elseif($key -eq 'n' -and $rowKey -like 'Opening*'){
                    # Прочерк показывает, что количество определяется проекцией наружной арматуры.
                }elseif($key -eq 'n'){
                    $cell.Validation.Add(1,1,7,'0');$cell.Validation.IgnoreBlank=$true
                }elseif($key -in @('as','d','s0','s1','d2','a2','d3','a3')){
                    $cell.Validation.Add(2,1,7,'0');$cell.Validation.IgnoreBlank=$true
                }
                if($options.Count -or ($key -eq 'n' -and $rowKey -notlike 'Opening*') -or $key -in @('as','d','s0','s1','d2','a2','d3','a3')){$cell.Validation.ShowError=$true}
            }
        }
    }
}
