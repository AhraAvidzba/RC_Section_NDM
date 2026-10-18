# Проверяет готовую книгу: структуру геометрии, исходные значения и формулы единиц.
# Книга открывается только для чтения; смена единиц проверяется без сохранения.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$ReportPath,
    [switch]$CheckDefaults
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$report=[IO.Path]::GetFullPath((Join-Path $root $ReportPath))
if(-not $report.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)){throw 'Report must stay inside the workspace.'}
$beforeHash=(Get-FileHash -LiteralPath $path).Hash
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class StructuredGeometryWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
$unitCells=New-Object 'System.Collections.Generic.List[object]'
$tables=@();$defaultCount=0;$conditionalChecks=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$identity=0;[void][StructuredGeometryWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$identity)
    $ownedPid=[int]$identity;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    $book=$excel.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Read-only opening required.'}
    $sheet=$book.Worksheets.Item('Config')
    for($c=9;$c -le 24;$c++){if($sheet.Columns.Item($c).ColumnWidth -ne 17){throw "Wrong Config column width: $c"}}
    $definitions=@(Get-GeometrySettingsCatalog)
    foreach($definition in $definitions+@(Get-RebarSettingsCatalog)){
        $range=$book.Names.Item($definition.RangeName).RefersToRange
        $width=5;if($definition.RangeName -eq 'rngRebarInput'){$width=11}
        $rules=0;if($definition.RangeName -eq 'rngRebarInput'){$rules=3}
        if($range.Columns.Count -ne $width -or $range.FormatConditions.Count -ne $rules){throw "Wrong table: $($definition.RangeName)"}
        foreach($edge in 7,8,9,10,11,12){if($range.Borders.Item($edge).LineStyle -ne -4115){throw 'Full dotted borders required.'}}
        if($range.WrapText -ne $false){throw 'Input tables must not wrap text.'}
        $blocks=@(Get-GeometryTableBlocks $range)
        if($blocks.Count -ne $definition.Blocks.Count){throw 'Wrong block count.'}
        for($i=0;$i -lt $blocks.Count;$i++){
            $block=$blocks[$i];$expected=$definition.Blocks[$i]
            if($block.Key -cne $expected.Key){throw 'Wrong block order.'}
            for($c=0;$c -lt $expected.Keys.Count;$c++){
                if([string]$range.Cells.Item($block.KeyRow,$c+1).Value2 -cne $expected.Keys[$c]){throw 'Wrong column key.'}
                if([string]$range.Cells.Item($block.HeaderRow,$c+1).Value2 -cne $expected.Headers[$c]){throw 'Wrong column caption.'}
                if($expected.Units -and $expected.Units[$c] -ceq 'мм'){$unitCells.Add($range.Cells.Item($block.FirstRow,$c+1))}
            }
            $first=$block.FirstRow
            if($expected.Units){$first++}
            for($r=0;$r -lt $expected.Rows.Count;$r++){
                $values=$expected.Rows[$r]
                $expectedKey=$values[0]
                if($expectedKey -is [string] -and $definition.RangeName -ne 'rngRebarInput'){
                    if([string]$range.Cells.Item($first+$r,1).Value2 -cne $expectedKey){throw 'Wrong row key.'}
                }
                if($block.Columns.ContainsKey('unit') -and $values[[int]$block.Columns['unit']-1] -ceq 'мм'){
                    $unitCells.Add($range.Cells.Item($first+$r,[int]$block.Columns['unit']))
                }
                foreach($key in @($block.Columns.Keys)){
                    if($key -in @('key','name','unit','note')){continue}
                    $column=[int]$block.Columns[$key];$cell=$range.Cells.Item($first+$r,$column)
                    if($block.Key -eq 'Edges'){
                        if(-not $cell.HasFormula -or ([string]$cell.Formula).Length -gt 80){throw 'Control size must be a short Excel formula.'}
                        if($cell.Interior.Color -eq 15135487){throw 'Computed control size is highlighted as input.'}
                        continue
                    }
                    if($cell.HasFormula){if($cell.Interior.Color -eq 15135487){throw 'Formula is highlighted as input.'};continue}
                    if([string]$cell.Value2 -notin @('-','—') -and $cell.Interior.Color -ne 15135487){throw ('Input is not orange: '+$cell.Address())}
                    if($CheckDefaults){
                        $value=$values[$column-1]
                        if([string]$cell.Value2 -cne [string]$value){throw "Wrong default: $($definition.RangeName)/$($block.Key)/$($values[0])/$key"}
                        $defaultCount++
                    }
                }
            }
        }
        $tables+=,[ordered]@{Name=$definition.RangeName;Address=$range.Address();Blocks=$blocks.Count}
    }
    $common=$book.Names.Item('rngRebarInput').RefersToRange
    $firstGeometry=$book.Names.Item($definitions[0].RangeName).RefersToRange
    if($firstGeometry.Row -ne 87 -or $firstGeometry.Column -ne 8 -or $common.Row -ne $firstGeometry.Row -or $common.Column -ne 14){throw 'Geometry/reinforcement layout changed.'}
    $last=0
    foreach($definition in $definitions){$area=$book.Names.Item($definition.RangeName).RefersToRange;if($area.Row -le $last){throw 'Geometry tables overlap.'};$last=$area.Row+$area.Rows.Count}
    $system=$book.Names.Item('rngSystemSettings').RefersToRange;$selector=$null
    for($r=2;$r -le $system.Rows.Count;$r++){if([string]$system.Cells.Item($r,1).Value2 -eq 'Geometry.Type'){$selector=$system.Cells.Item($r,2)}}
    if($null -eq $selector){throw 'Geometry.Type missing.'}
    $blocks=@(Get-GeometryTableBlocks $common);$main=$blocks | Where-Object Key -eq 'Main';$extra=$blocks | Where-Object Key -eq 'Extra'
    if($main.LastRow-$main.FirstRow -ne 8 -or $extra.LastRow-$extra.FirstRow -ne 8){throw 'Common table must have exactly eight rows in each block.'}
    # Проверяем фактическую заливку Excel, а не только текст условий в файле.
    # Серые поля сохраняют значения; отключённый диаметр сам по себе не запрещает ввод.
    function Test-RebarInputDisplay {
        param([object]$Range,[object[]]$Blocks)
        $checks=0
        foreach($block in $Blocks){
            foreach($columnKey in @($block.Columns.Keys)){
                if($columnKey -in @('key','name','unit','note')){continue}
                for($r=$block.FirstRow+1;$r -le $block.LastRow;$r++){
                    $key=[string]$Range.Cells.Item($r,[int]$block.Columns['key']).Value2
                    $cell=$Range.Cells.Item($r,[int]$block.Columns[$columnKey])
                    $expected=15135487
                    if(-not $key -or ($block.Key -eq 'Main' -and $columnKey -eq 'n' -and $key.StartsWith('HollowRectangle.Opening'))){$expected=15132390}
                    if($cell.DisplayFormat.Interior.Color -ne $expected){throw ('Wrong displayed input color: '+$key+'/'+$columnKey+' '+$cell.Address())}
                    if($cell.Interior.Color -ne 15135487){throw 'Conditional color changed the static input format.'}
                    $checks++
                }
            }
        }
        return $checks
    }
    $sheet.Activate()
    $manualBefore=$common.Formula
    $switches=@()
    foreach($definition in $definitions){
        $selector.Value2=[string]$definition.Type;$sheet.Calculate()
        $source=$book.Names.Item($definition.RangeName).RefersToRange
        $edges=@(Get-GeometryTableBlocks $source | Where-Object Key -eq 'Edges')[0]
        for($i=0;$i -lt 8;$i++){
            $expectedKey='';$expectedName=''
            if($edges.FirstRow+$i -le $edges.LastRow){
                $short=[string]$source.Cells.Item($edges.FirstRow+$i,1).Value2
                if($short){$expectedKey=$definition.Type+'.'+$short;$expectedName=[string]$source.Cells.Item($edges.FirstRow+$i,2).Value2}
            }
            foreach($block in @($main,$extra)){
                $row=$block.FirstRow+1+$i
                if([string]$common.Cells.Item($row,1).Value2 -cne $expectedKey -or [string]$common.Cells.Item($row,2).Value2 -cne $expectedName){throw 'Formula keys/names do not follow selected geometry.'}
                if(-not $common.Cells.Item($row,1).HasFormula -or -not $common.Cells.Item($row,2).HasFormula){throw 'Key/name must be ordinary Excel formulas.'}
            }
        }
        $conditionalChecks+=Test-RebarInputDisplay $common @($main,$extra)
        $switches+=,[ordered]@{Type=$definition.Type;EdgeRows=$edges.LastRow-$edges.FirstRow+1}
    }
    $selector.Value2='RectSet';$rect=$book.Names.Item('rngRectSetGeometry').RefersToRange
    $shape=Get-GeometryTableCell $rect 'General' 'Shape';$originalShape=$shape.Value2
    foreach($subtype in @('Rectangle','LSection','TwoRectangles')){
        $shape.Value2=$subtype;$sheet.Calculate()
        for($i=0;$i -lt 8;$i++){
            $key=[string]$common.Cells.Item($main.FirstRow+1+$i,1).Value2
            if(($subtype -eq 'Rectangle' -and $i -ge 4) -ne ($key -eq '')){throw 'RectSet inactive lower edges are incorrect.'}
        }
        $conditionalChecks+=Test-RebarInputDisplay $common @($main,$extra)
    }
    $shape.Value2=$originalShape;$sheet.Calculate()
    $afterSwitch=$common.Formula
    for($r=1;$r -le $common.Rows.Count;$r++){for($c=1;$c -le $common.Columns.Count;$c++){
        if([string]$manualBefore[$r,$c] -cne [string]$afterSwitch[$r,$c]){throw 'Type switching changed manual reinforcement data.'}
    }}
    $choiceNames=@{}
    foreach($field in @('bind2','pos2','bind3','pos3')){
        for($i=0;$i -lt 8;$i++){
            $cell=$common.Cells.Item($extra.FirstRow+1+$i,[int]$extra.Columns[$field]);$formula=[string]$cell.Validation.Formula1
            if($formula -notmatch '^=[A-Za-z_][A-Za-z0-9_]*$' -or $choiceNames.ContainsKey($formula)){throw 'Choice lists must have independent named sources.'}
            $choiceNames[$formula]=$true
            if($book.Names.Item($formula.Substring(1)).RefersToRange.Rows.Count -ne 2){throw 'Invalid choice source.'}
        }
    }
    $orangeCount=0
    $inputs=@{rngSystemSettings=@(2);rngUnitSettings=@(2,4);rngSignConventionSettings=@(2);rngSteelMaterialParameters=@(2,3);
        rngConcreteMaterialParameters=@(2,3);rngCalculationProfiles=@(3,4,5,6);rngPlotAnnotationSettings=@(2,3);
        rngLoadCombinations=@(1,2,3,4,5,6,7);rngStabilityDurationLoads=@(1,2,3,4);rngSP35Table721=@(1,2,3,4,5,6,7,8)}
    foreach($name in $inputs.Keys){
        $range=$book.Names.Item($name).RefersToRange
        foreach($column in $inputs[$name]){for($r=2;$r -le $range.Rows.Count;$r++){
            $cell=$range.Cells.Item($r,$column);$key=[string]$range.Cells.Item($r,1).Value2
            $isComment=($name -eq 'rngLoadCombinations' -and $column -eq 7)
            if($key.StartsWith('[') -or ($name -eq 'rngSystemSettings' -and -not $key.Contains('.')) -or (-not $isComment -and ($cell.HasFormula -or [string]$cell.Value2 -in @('-','—')))){continue}
            if($cell.MergeCells -and ($cell.MergeArea.Row -ne $cell.Row -or $cell.MergeArea.Column -ne $cell.Column)){continue}
            $alignment=-4108;if($isComment){$alignment=-4131}
            if($cell.Interior.Color -ne 15135487 -or $cell.HorizontalAlignment -ne $alignment){throw ('Input formatting: '+$name+' '+$cell.Address())}
            $orangeCount++
        }}
    }
    foreach($definition in $definitions+@(Get-RebarSettingsCatalog)){
        $range=$book.Names.Item($definition.RangeName).RefersToRange
        foreach($block in @(Get-GeometryTableBlocks $range)){
            $note=$range.Cells.Item($block.HeaderRow,[int]$block.Columns['note']).Resize($block.LastRow-$block.HeaderRow+1,1)
            if($note.HorizontalAlignment -ne -4152 -or $note.WrapText -ne $false){throw 'Comments must stay right-aligned in their own cell.'}
        }
    }
    $profiles=$book.Names.Item('rngCalculationProfiles').RefersToRange
    for($r=2;$r -le $profiles.Rows.Count;$r++){
        $cell=$profiles.Cells.Item($r,7)
        if(-not $cell.MergeCells -and ($cell.HorizontalAlignment -ne -4152 -or $cell.WrapText -ne $false)){throw 'Profile comments must stay right-aligned without wrapping.'}
    }
    $expectedFormula=Get-InputUnitFormulaForUnit 'мм'
    foreach($cell in $unitCells){if([string]$cell.Formula -cne $expectedFormula){throw "Unit formula missing: $($cell.Address())"}}
    $units=$book.Names.Item('rngUnitSettings').RefersToRange;$length=$null
    for($r=2;$r -le $units.Rows.Count;$r++){if([string]$units.Cells.Item($r,1).Value2 -ceq 'Length'){$length=$units.Cells.Item($r,2);break}}
    if($null -eq $length){throw 'Length input setting missing.'}
    foreach($unit in @('mm','cm','m')){
        $length.Value2=$unit;$sheet.Calculate()
        foreach($cell in $unitCells){if([string]$cell.Value2 -cne $unit){throw "Unit caption did not follow INPUT Length: $($cell.Address())"}}
    }
    [ordered]@{Passed=$true;WorkbookSHA256=$beforeHash;Tables=$tables;DefaultFieldsChecked=$defaultCount;
        DynamicLengthUnitCells=$unitCells.Count;InputLengthUnits=@('mm','cm','m');UniformConfigWidth=17;
        MacroSecurity=3;FormulaTypeSwitches=$switches;OtherConfigInputCells=$orangeCount;InputColorRGB=@(255,242,230);
        RectSetSubtypes=@('Rectangle','LSection','TwoRectangles');ManualInputsPreservedOnSwitch=$true;IndependentRebarChoiceLists=$choiceNames.Count;
        RebarConditionalFormattingRules=3;DisplayedInputColorsChecked=$conditionalChecks;UnusedInputColorRGB=@(230,230,230);
        Saved=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $beforeHash){throw 'Read-only check changed the workbook.'}
Write-Output "STRUCTURED_GEOMETRY: defaults=$defaultCount; dynamic units=$($unitCells.Count); three INPUT units; source unchanged."
