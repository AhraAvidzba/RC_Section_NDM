# Переставляет готовые блоки Справки, сохраняя ячейки, рисунки и точные ссылки.
$script:HelpMainTitles=@(
    'Подготовка геометрии в AutoCAD и импорт',
    'Поиск НДС текущего сочетания',
    'Поиск предельной несущей способности',
    'Расчет ширины раскрытия нормальных трещин',
    'Расчет продольного изгиба и устойчивости',
    'Справка по результатам расчета листа Results',
    'Добавление нового типа сечения',
    'Справка по настройкам листа Config'
)

# Ищет содержимое, а не оставшийся после старого большого гайда UsedRange.
function Get-HelpContentLastRow($Sheet) {
    $cell=$Sheet.Cells.Find('*',$Sheet.Cells.Item(1,1),-4123,2,1,2,$false)
    if ($null -eq $cell) {return 0}
    return [int]$cell.Row
}

# Проверяет уникальность главных заголовков до перемещения стандартных блоков.
function Get-HelpMainRows($Sheet) {
    $last=Get-HelpContentLastRow $Sheet
    $matrix=$Sheet.Range("A1:A$last").Value2
    $rows=@{}
    for ($r=1;$r -le $last;$r++) {
        $title=[string]$matrix[$r,1]
        if ($title -in $script:HelpMainTitles) {
            if ($rows.ContainsKey($title)) {throw ('Duplicate Help section: '+$title)}
            $rows[$title]=$r
        }
    }
    return $rows
}

# Переводит номер строки после переноса непрерывного блока выше него.
function Convert-HelpMovedRow([int]$Row,[int]$First,[int]$Last,[int]$Destination) {
    if ($Destination -gt $Last) {
        $count=$Last-$First+1
        if ($Row -ge $First -and $Row -le $Last) {return $Destination-$count+$Row-$First}
        if ($Row -gt $Last -and $Row -lt $Destination) {return $Row-$count}
        return $Row
    }
    if ($Row -ge $First -and $Row -le $Last) {return $Destination+$Row-$First}
    if ($Row -ge $Destination -and $Row -lt $First) {return $Row+$Last-$First+1}
    return $Row
}

# Восстанавливает свободные рисунки и строковые SubAddress по карте строк Cut/Insert.
function Move-HelpRowsBefore($Workbook,$Sheet,[int]$First,[int]$Last,[int]$Destination) {
    if (($Destination -ge $First -and $Destination -le $Last+1) -or $Last -lt $First) {throw 'Invalid Help block move.'}
    $ownedNames=@(foreach($name in $Workbook.Names){
        try{$range=$name.RefersToRange}catch{continue}
        if($null -eq $range -or $range.Areas.Count -ne 1 -or $range.Worksheet.Name -cne $Sheet.Name){continue}
        if($range.Row -ge $First -and $range.Row+$range.Rows.Count-1 -le $Last){
            [pscustomobject]@{Name=$name;Row=[int]$range.Row;Column=[int]$range.Column;Rows=[int]$range.Rows.Count;Columns=[int]$range.Columns.Count}
        }
    })
    $pictures=@($Sheet.Shapes | ForEach-Object {
        [pscustomobject]@{Shape=$_;Row=[int]$_.TopLeftCell.Row;Offset=[double]$_.Top-[double]$Sheet.Rows.Item($_.TopLeftCell.Row).Top;
            Left=[double]$_.Left;Width=[double]$_.Width;Height=[double]$_.Height}
    })
    $links=New-Object System.Collections.Generic.List[object]
    foreach ($worksheet in $Workbook.Worksheets) {
        foreach ($link in $worksheet.Hyperlinks) {
            if ([string]$link.SubAddress -match "^'?Справка'?!([A-Z]+)([0-9]+)$") {
                $links.Add([pscustomobject]@{Link=$link;Column=$Matches[1];Row=[int]$Matches[2]})
            }
        }
    }
    $Sheet.Rows.Item("${First}:${Last}").Cut() | Out-Null
    $Sheet.Rows.Item($Destination).Insert(-4121) | Out-Null
    $Workbook.Application.CutCopyMode=$false
    foreach($saved in $ownedNames){
        $row=Convert-HelpMovedRow $saved.Row $First $Last $Destination
        $saved.Name.RefersTo="='"+$Sheet.Name.Replace("'","''")+"'!"+$Sheet.Cells.Item($row,$saved.Column).Resize($saved.Rows,$saved.Columns).Address()
    }
    foreach ($picture in $pictures) {
        $row=Convert-HelpMovedRow $picture.Row $First $Last $Destination
        $picture.Shape.Top=[double]$Sheet.Rows.Item($row).Top+$picture.Offset
        $picture.Shape.Left=$picture.Left;$picture.Shape.Width=$picture.Width;$picture.Shape.Height=$picture.Height
    }
    foreach ($saved in $links) {
        $row=Convert-HelpMovedRow $saved.Row $First $Last $Destination
        $saved.Link.SubAddress="'Справка'!"+$saved.Column+$row
    }
    $script:HelpRowMoves+=,[ordered]@{First=$First;Last=$Last;Destination=$Destination}
}

# Меняет число строк только в свободном месте; сохраняет ссылки и положение рисунков.
function Resize-HelpEmptyRows($Workbook,$Sheet,[int]$First,[int]$Count,[bool]$Insert) {
    if ($First -lt 1 -or $Count -lt 1) {throw 'Invalid Help row resize.'}
    $last=$First+$Count-1
    if (-not $Insert -and $Workbook.Application.WorksheetFunction.CountA($Sheet.Rows.Item("${First}:${last}")) -ne 0) {throw 'Refusing to delete nonempty Help rows.'}
    $Sheet.Rows.Hidden=$false
    $pictures=@($Sheet.Shapes | ForEach-Object {
        if (-not $Insert -and $_.TopLeftCell.Row -le $last -and $_.BottomRightCell.Row -ge $First) {throw 'A picture overlaps deleted Help rows.'}
        [pscustomobject]@{Shape=$_;Row=[int]$_.TopLeftCell.Row;Offset=[double]$_.Top-[double]$Sheet.Rows.Item($_.TopLeftCell.Row).Top;
            Left=[double]$_.Left;Width=[double]$_.Width;Height=[double]$_.Height}
    })
    $links=@(foreach ($worksheet in $Workbook.Worksheets) {
        foreach ($link in $worksheet.Hyperlinks) {
            if ([string]$link.SubAddress -match "^'?Справка'?!([A-Z]+)([0-9]+)$") {
                $row=[int]$Matches[2]
                if (-not $Insert -and $row -ge $First -and $row -le $last) {throw 'A hyperlink points to deleted Help rows.'}
                [pscustomobject]@{Link=$link;Column=$Matches[1];Row=$row}
            }
        }
    })
    $delta=$Count
    if ($Insert) {$Sheet.Rows.Item("${First}:${last}").Insert(-4121) | Out-Null}
    else {$Sheet.Rows.Item("${First}:${last}").Delete(-4162) | Out-Null;$delta=-$Count}
    foreach ($picture in $pictures) {
        $row=$picture.Row;if ($row -ge $First) {$row+=$delta}
        $picture.Shape.Top=[double]$Sheet.Rows.Item($row).Top+$picture.Offset
        $picture.Shape.Left=$picture.Left;$picture.Shape.Width=$picture.Width;$picture.Shape.Height=$picture.Height
    }
    foreach ($saved in $links) {
        $row=$saved.Row;if ($row -ge $First) {$row+=$delta}
        $saved.Link.SubAddress="'Справка'!"+$saved.Column+$row
    }
}

# Создаёт одну outline-группу для каждого раздела; Config последний и раскрыт.
function Set-HelpSectionOrderAndOutline($Workbook) {
    $script:HelpRowMoves=@()
    $sheet=$Workbook.Worksheets.Item('Справка')
    $sheet.Rows.Hidden=$false
    Repair-HelpConfigEntries $Workbook
    $rows=Get-HelpMainRows $sheet
    $configTitle=$script:HelpMainTitles[-1]
    foreach ($title in $script:HelpMainTitles) {
        if (-not $rows.ContainsKey($title)) {throw ('Missing Help section: '+$title)}
    }
    foreach ($title in $script:HelpMainTitles[5..6]) {
        $rows=Get-HelpMainRows $sheet
        $first=[int]$rows[$title];$destination=[int]$rows[$configTitle]
        if ($first -gt $destination) {
            $next=@($rows.Values | Where-Object {$_ -gt $first} | Sort-Object)[0]
            $last=Get-HelpContentLastRow $sheet
            if ($null -ne $next) {$last=[int]$next-1}
            # Именованный короткий раздел включает завершающую пустую строку.
            # Перенос только непустых строк оставлял её на старом месте и растягивал имя.
            if($title -eq 'Добавление нового типа сечения'){
                try{$guide=$Workbook.Names.Item('rngSectionRecipeGuide').RefersToRange}catch{$guide=$null}
                if($null -ne $guide -and $guide.Row -eq $first){$last=[Math]::Max($last,$guide.Row+$guide.Rows.Count-1)}
            }
            Move-HelpRowsBefore $Workbook $sheet $first $last $destination
        }
    }
    $rows=Get-HelpMainRows $sheet
    $ordered=@($script:HelpMainTitles | ForEach-Object {[int]$rows[$_]})
    for ($i=1;$i -lt $ordered.Count;$i++) {
        if ($ordered[$i] -le $ordered[$i-1]) {throw 'Help sections have an unexpected order.'}
    }
    $last=Get-HelpContentLastRow $sheet
    $sheet.Cells.ClearOutline() | Out-Null
    $sheet.Rows.Hidden=$false;$sheet.Outline.SummaryRow=0
    for ($i=0;$i -lt $ordered.Count;$i++) {
        $end=$last
        if ($i+1 -lt $ordered.Count) {$end=$ordered[$i+1]-1}
        if ($end -gt $ordered[$i]) {$sheet.Rows.Item(($ordered[$i]+1).ToString()+':'+$end).Group() | Out-Null}
    }
    $sheet.Outline.ShowLevels(1) | Out-Null
    $sheet.Rows.Item([int]$rows[$configTitle]).ShowDetail=$true
    $sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item($last,6)).Borders.Item(10).LineStyle=1
    return [ordered]@{Titles=$script:HelpMainTitles;Rows=$ordered;LastRow=$last;DefaultExpanded=$configTitle;Groups=8;Moves=$script:HelpRowMoves}
}

# Узкие прежние обновления дописывали эти настройки после Results.
# Переносим весь абзац внутрь Config, сохраняя реальные адреса Подробнее.
function Repair-HelpConfigEntries($Workbook) {
    $sheet=$Workbook.Worksheets.Item('Справка')
    foreach ($key in @('SLS.Crack.ToleranceLambda','General.PlateauSlopeFactor')) {
        $rows=Get-HelpMainRows $sheet
        $config=[int]$rows['Справка по настройкам листа Config']
        if ($config -eq 0) {throw 'Config Help section is missing.'}
        $next=@($rows.Values | Where-Object {$_ -gt $config} | Sort-Object)[0]
        $found=$sheet.Cells.Find($key,$sheet.Cells.Item(1,1),-4163,1)
        if ($null -eq $found) {continue}
        if ($found.Row -gt $config -and ($null -eq $next -or $found.Row -lt [int]$next)) {continue}
        $first=[int]$found.Row;$last=$first
        $bottom=Get-HelpContentLastRow $sheet
        while ($last -lt $bottom -and [string]$sheet.Cells.Item($last+1,1).Value2 -ne '') {$last++}
        if ($first -lt $config) {Move-HelpRowsBefore $Workbook $sheet $first $last ($bottom+1)}
        else {Move-HelpRowsBefore $Workbook $sheet $first $last ([int]$next)}
    }
}
