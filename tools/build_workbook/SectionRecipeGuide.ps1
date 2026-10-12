# Общая явная установка последнего раздела Справки; расчёт её не вызывает.
function Get-RecipeGuideTextHash([string[]]$Lines) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($Lines -join "`n")))).Replace('-','').ToLowerInvariant()}
    finally {$sha.Dispose()}
}

function Get-RecipeGuideContentBottom($Sheet) {
    $last=$Sheet.Cells.Find('*',$Sheet.Cells.Item(1,1),-4123,2,1,2,$false)
    $bottom=0
    if ($null -ne $last) {$bottom=[int]$last.Row}
    foreach ($shape in $Sheet.Shapes) {$bottom=[Math]::Max($bottom,[int]$shape.BottomRightCell.Row)}
    return $bottom
}

. (Join-Path $PSScriptRoot 'HelpOutline.ps1')

function Get-RecipeGuideRowRuns($Addresses) {
    $first=0;$last=0
    foreach ($address in $Addresses) {
        if ($address -notmatch '^A([0-9]+):F[0-9]+$') {throw 'Invalid internal guide row.'}
        $row=[int]$Matches[1]
        if ($first -eq 0) {$first=$row;$last=$row}
        elseif ($row -eq $last+1) {$last=$row}
        else {[pscustomobject]@{First=$first;Last=$last};$first=$row;$last=$row}
    }
    if ($first -gt 0) {[pscustomobject]@{First=$first;Last=$last}}
}

function Get-RecipeGuideNativeLines($Sheet,[int]$First,[int]$Last) {
    $matrix=$Sheet.Range($Sheet.Cells.Item($First,1),$Sheet.Cells.Item($Last,6)).Value2
    for ($row=1;$row -le $Last-$First+1;$row++) {
        $fields=for ($column=1;$column -le 6;$column++) {[string]$matrix[$row,$column]}
        $lastValue=5
        while ($lastValue -ge 0 -and $fields[$lastValue] -eq '') {$lastValue--}
        if ($lastValue -lt 0) {''} else {$fields[0..$lastValue] -join "`t"}
    }
}

function Set-RecipeGuideMatrix($Target,[object[,]]$Matrix) {
    # Отдельный типизированный call-site для массива: общая сборка прежде
    # записывает скаляры в Value2 через тот же PowerShell COM adapter.
    $Target.Value2=$Matrix
}

function Update-SectionRecipeGuide {
    param([Parameter(Mandatory=$true)]$Workbook,[Parameter(Mandatory=$true)][string]$Root)
    $application=$Workbook.Application
    $saved=[ordered]@{Calculation=[int]$application.Calculation;ScreenUpdating=[bool]$application.ScreenUpdating;
        EnableEvents=[bool]$application.EnableEvents;DisplayAlerts=[bool]$application.DisplayAlerts}
    try {
        $application.ScreenUpdating=$false;$application.EnableEvents=$false;$application.DisplayAlerts=$false
        $application.Calculation=-4135
        $layoutPath=Join-Path $Root 'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json'
        $layout=Get-Content -LiteralPath $layoutPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($layout.schema -ne 2) {throw 'Unsupported guide layout schema.'}
        foreach ($property in $layout.sourceHashes.PSObject.Properties) {
            if ((Get-FileHash -LiteralPath (Join-Path $Root $property.Name)).Hash.ToLowerInvariant() -cne $property.Value) {
                throw ('Guide content is stale: '+$property.Name)
            }
        }
        if ((Get-RecipeGuideTextHash @($layout.rows | Select-Object -SkipLast 1 | ForEach-Object {$_.text})) -cne $layout.bodySHA256) {throw 'Guide layout hash mismatch.'}
        $sheet=$Workbook.Worksheets.Item('Справка')
        $owner=$null
        foreach ($name in $Workbook.Names) {if ($name.Name -eq 'rngSectionRecipeGuide') {$owner=$name;break}}
        $oldStart=0;$oldEnd=0
        if ($null -ne $owner) {
            $old=$owner.RefersToRange
            $oldStart=[int]$old.Row;$oldEnd=$oldStart+[int]$old.Rows.Count-1
            if ($old.Worksheet.Name -cne $sheet.Name -or $old.Column -ne 1 -or $old.Columns.Count -ne 6 -or $old.Rows.Count -lt 3) {throw 'Guide owner range is invalid.'}
            if ([string]$sheet.Cells.Item($oldStart,1).Value2 -cne $layout.title) {throw 'Guide owner title was changed; refusing to clear user content.'}
            $footer=[string]$sheet.Cells.Item($oldEnd,1).Value2
            if ($footer -notmatch '^NDM\.RECIPE\.GUIDE/[12]\|([0-9a-f]{64})\|[0-9a-f]{64}$') {throw 'Guide owner footer is invalid.'}
            $expected=$Matches[1]
            $oldLines=@(Get-RecipeGuideNativeLines $sheet $oldStart ($oldEnd-1))
            if ((Get-RecipeGuideTextHash $oldLines) -cne $expected) {throw 'Guide owner content was edited; refusing to clear user content.'}
            foreach ($shape in $sheet.Shapes) {
                if ($shape.TopLeftCell.Row -le $oldEnd -and $shape.BottomRightCell.Row -ge $oldStart) {throw 'A user shape overlaps the owned guide.'}
            }
            # Удаляется только принадлежащий нам блок, без удаления строк листа.
            for ($level=0;$level -lt 8;$level++) {try {$old.EntireRow.Ungroup() | Out-Null} catch {break}}
            $old.UnMerge();$old.Clear() | Out-Null
        }
        $bottom=Get-RecipeGuideContentBottom $sheet
        $start=$bottom+3
        # Короткий блок повторно записывается в своё свободное место перед Config.
        if ($oldStart -gt 0) {
            $oldCount=$oldEnd-$oldStart+1
            if ($layout.rows.Count -gt $oldCount) {Resize-HelpEmptyRows $Workbook $sheet ($oldEnd+1) ($layout.rows.Count-$oldCount) $true}
            elseif ($layout.rows.Count -lt $oldCount) {Resize-HelpEmptyRows $Workbook $sheet ($oldStart+$layout.rows.Count) ($oldCount-$layout.rows.Count) $false}
            $start=$oldStart
        }
        $end=$start+$layout.rows.Count-1
        $range=$sheet.Range($sheet.Cells.Item($start,1),$sheet.Cells.Item($end,6))
        $width=[double]$range.Width
        if ($width -lt 400) {throw 'Help columns are too narrow for readable guide text.'}
        $range.Merge($true);$range.NumberFormat='@'
        $values=New-Object 'object[,]' $layout.rows.Count,6
        for ($i=0;$i -lt $layout.rows.Count;$i++) {
            $texts=@([string]$layout.rows[$i].text)
            if ($layout.rows[$i].cells) {
                $row=$start+$i
                $tableRow=$sheet.Range($sheet.Cells.Item($row,1),$sheet.Cells.Item($row,6))
                $tableRow.UnMerge()
                $texts=@($layout.rows[$i].cells)
                # A — имя, B — тип/default, C:F — смысл; старые ширины
                # колонок листа не меняются. Для двух полей B:F объединяются.
                if ($texts.Count -eq 2) {$sheet.Range($sheet.Cells.Item($row,2),$sheet.Cells.Item($row,6)).Merge()}
                else {$sheet.Range($sheet.Cells.Item($row,3),$sheet.Cells.Item($row,6)).Merge()}
            }
            for ($column=0;$column -lt $texts.Count;$column++) {
                $text=[string]$texts[$column]
                # Апостроф комментария VBA сохраняется в Value2, не prefix.
                if ($text.StartsWith("'")) {$text="'"+$text}
                $values[$i,$column]=$text
            }
        }
        Set-RecipeGuideMatrix $range $values
        $range.Font.Name='Arial';$range.Font.Size=11;$range.Font.Color=0
        $range.WrapText=$true;$range.VerticalAlignment=-4160;$range.HorizontalAlignment=-4131
        $range.Borders.LineStyle=1;$range.Borders.Color=14277081
        $range.EntireRow.Hidden=$false;$range.EntireRow.RowHeight=18
        $byStyle=@{};$byHeight=@{}
        for ($i=0;$i -lt $layout.rows.Count;$i++) {
            $row=$start+$i;$style=[string]$layout.rows[$i].style
            if (-not $byStyle.ContainsKey($style)) {$byStyle[$style]=New-Object System.Collections.Generic.List[string]}
            $byStyle[$style].Add("A${row}:F${row}")
            $size=11;$pointPerCharacter=6.7
            if ($style -in 'title','h1') {$size=16;$pointPerCharacter=9.8}
            elseif ($style -in 'h2','h3') {$size=13;$pointPerCharacter=8}
            # Консервативная оценка с запасом; длинная строка остаётся одним
            # исходным значением, чтобы пояснение или вызов можно было скопировать.
            $characters=[Math]::Max(20,[Math]::Floor(($width-16)/$pointPerCharacter))
            $lines=[Math]::Max(1,[Math]::Ceiling(([string]$layout.rows[$i].text).Length/$characters))
            $height=[Math]::Min(409,($size+7)*$lines+3)
            if ($layout.rows[$i].cells) {
                $fields=@($layout.rows[$i].cells);$maximumLines=1
                for ($column=0;$column -lt $fields.Count;$column++) {
                    $cell=$sheet.Cells.Item($row,$column+1)
                    $cellWidth=[double]$cell.MergeArea.Width
                    $span=[Math]::Max(10,[Math]::Floor(($cellWidth-12)/$pointPerCharacter))
                    $maximumLines=[Math]::Max($maximumLines,[Math]::Ceiling(([string]$fields[$column]).Length/$span))
                }
                $height=[Math]::Min(409,($size+7)*$maximumLines+3)
                if ($layout.rows[$i].tableHeader) {
                    $tableRow=$sheet.Range($sheet.Cells.Item($row,1),$sheet.Cells.Item($row,6))
                    $tableRow.Font.Bold=$true;$tableRow.Interior.Color=14277081
                }
            }
            if (-not $byHeight.ContainsKey($height)) {$byHeight[$height]=New-Object System.Collections.Generic.List[string]}
            $byHeight[$height].Add("A${row}:F${row}")
        }
        foreach ($style in $byStyle.Keys) {
            $addresses=$byStyle[$style]
            foreach ($run in (Get-RecipeGuideRowRuns $addresses)) {
                $styled=$sheet.Range($sheet.Cells.Item($run.First,1),$sheet.Cells.Item($run.Last,6))
                if ($style -in 'title','h1','h2','h3') {$styled.Font.Bold=$true;$styled.Interior.Color=14277081}
                if ($style -in 'title','h1') {$styled.Font.Size=16}
                elseif ($style -in 'h2','h3') {$styled.Font.Size=13}
                elseif ($style -eq 'code') {$styled.Font.Name='Consolas';$styled.Interior.Color=15921906}
                elseif ($style -eq 'toc') {$styled.Font.Color=12611584;$styled.Font.Underline=2}
                elseif ($style -eq 'footer') {$styled.Font.Color=7368816}
            }
        }
        foreach ($height in $byHeight.Keys) {
            $addresses=$byHeight[$height]
            foreach ($run in (Get-RecipeGuideRowRuns $addresses)) {
                $sheet.Range($sheet.Cells.Item($run.First,1),$sheet.Cells.Item($run.Last,6)).EntireRow.RowHeight=[double]$height
            }
        }
        $sectionRows=@{}
        for ($i=0;$i -lt $layout.sections.Count;$i++) {
            $section=$layout.sections[$i];$sectionRow=$start+[int]$section.index
            $sectionRows[$section.anchor]=$sectionRow
            # Заголовок виден, тело сворачивается независимо от старой Справки.
            $bodyEnd=$end-1
            if ($i+1 -lt $layout.sections.Count) {$bodyEnd=$start+[int]$layout.sections[$i+1].index-1}
            if ($bodyEnd -gt $sectionRow) {$sheet.Range($sheet.Cells.Item($sectionRow+1,1),$sheet.Cells.Item($bodyEnd,1)).EntireRow.Group() | Out-Null}
        }
        for ($i=0;$i -lt $layout.rows.Count;$i++) {
            if ($layout.rows[$i].link) {
                $target=$sectionRows[$layout.rows[$i].link]
                $cell=$sheet.Cells.Item($start+$i,1)
                $sheet.Hyperlinks.Add($cell,'',("'Справка'!A"+$target),[Type]::Missing,[string]$layout.rows[$i].text) | Out-Null
                $cell.Font.Size=11
            }
        }
        $ref="='Справка'!"+$range.Address()
        if ($null -eq $owner) {$Workbook.Names.Add('rngSectionRecipeGuide',$ref) | Out-Null} else {$owner.RefersTo=$ref}
        $actualLines=@(Get-RecipeGuideNativeLines $sheet $start ($end-1))
        if ((Get-RecipeGuideTextHash $actualLines) -cne $layout.bodySHA256) {
            $different=@()
            for ($i=0;$i -lt $actualLines.Count;$i++) {
                if ($actualLines[$i] -cne [string]$layout.rows[$i].text) {
                    $different+=,[ordered]@{Index=$i;Actual=$actualLines[$i];Expected=[string]$layout.rows[$i].text}
                    if ($different.Count -ge 6) {break}
                }
            }
            throw ('Native guide text differs from source: '+($different | ConvertTo-Json -Compress))
        }
        $outline=Set-HelpSectionOrderAndOutline $Workbook
        $range=$Workbook.Names.Item('rngSectionRecipeGuide').RefersToRange
        $start=[int]$range.Row;$end=$start+[int]$range.Rows.Count-1
        $apiRows=[ordered]@{}
        foreach ($property in $layout.api.PSObject.Properties) {$apiRows[$property.Name]=$start+[int]$property.Value}
        return [ordered]@{Status='NATIVE_GUIDE_TEXT_MATCH';StartRow=$start;EndRow=$end;OldStartRow=$oldStart;OldEndRow=$oldEnd;
            PreviousContentBottom=$bottom;WidthPoints=$width;Rows=$layout.rows.Count;Members=$layout.members;
            BodySHA256=$layout.bodySHA256;ApiStatus=$layout.apiStatus;ApiRows=$apiRows;SectionRows=$sectionRows;Outline=$outline}
    } finally {
        $application.Calculation=$saved.Calculation
        $application.ScreenUpdating=$saved.ScreenUpdating
        $application.EnableEvents=$saved.EnableEvents
        $application.DisplayAlerts=$saved.DisplayAlerts
    }
}
