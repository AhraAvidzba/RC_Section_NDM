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
$tables=@();$defaultCount=0
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$identity=0;[void][StructuredGeometryWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$identity)
    $ownedPid=[int]$identity;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    $book=$excel.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Read-only opening required.'}
    $sheet=$book.Worksheets.Item('Config')
    for($c=9;$c -le 21;$c++){if($sheet.Columns.Item($c).ColumnWidth -ne 17){throw "Wrong Config column width: $c"}}
    foreach($definition in @(Get-GeometrySettingsCatalog)){
        $range=$book.Names.Item($definition.RangeName).RefersToRange
        if($range.Columns.Count -ne 11 -or $range.FormatConditions.Count -ne 0){throw "Wrong table: $($definition.RangeName)"}
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
                if([string]$range.Cells.Item($first+$r,1).Value2 -cne $values[0]){throw 'Wrong row key.'}
                if($block.Columns.ContainsKey('unit') -and $values[[int]$block.Columns['unit']-1] -ceq 'мм'){
                    $unitCells.Add($range.Cells.Item($first+$r,[int]$block.Columns['unit']))
                }
                foreach($key in @($block.Columns.Keys)){
                    if($key -in @('key','name','unit','note')){continue}
                    $column=[int]$block.Columns[$key];$cell=$range.Cells.Item($first+$r,$column)
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
        Saved=$false} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $beforeHash){throw 'Read-only check changed the workbook.'}
Write-Output "STRUCTURED_GEOMETRY: defaults=$defaultCount; dynamic units=$($unitCells.Count); three INPUT units; source unchanged."
