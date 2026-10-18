# Приёмка формул и оформления отдельного макета в настоящем Excel.
# Временные изменения восстанавливаются перед сохранением; рабочая книга не меняется.
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$path=[string]$manifest.file
$main=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$mainHash=(Get-FileHash -LiteralPath $main).Hash
$existing=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SharedRebarWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$checks=0
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
function Set-NumberCell($Cell,[double]$Value){$Cell.Value2=$Value}
function Keys([string]$Prefix){
    $active=0
    if($Prefix -ceq 'LSection'){$active=$manifest.sides.Count}
    if($Prefix -ceq 'Circle'){$active=1}
    for($i=0;$i -lt $manifest.capacity;$i++){
        $expected='';$caption=''
        if($i -lt $active){
            $row=67
            if($Prefix -ceq 'LSection'){$row=[int]$manifest.geometryRows[$i]}
            $key=[string]$sheet.Cells.Item($row,8).Value2
            if($key -cne ''){$expected=$Prefix+'.'+$key;$caption=[string]$sheet.Cells.Item($row,9).Value2}
        }
        foreach($r in @([int]$manifest.mainRows[$i],[int]$manifest.extraRows[$i])){
            $cell=$sheet.Cells.Item($r,8)
            Check ($cell.HasFormula -and [string]$cell.Value2 -ceq $expected) ('Wrong formula key at '+$cell.Address())
            $name=$sheet.Cells.Item($r,9)
            Check ($name.HasFormula -and [string]$name.Value2 -ceq $caption) ('Wrong formula side name at '+$name.Address())
        }
    }
}
function Lengths {
    foreach($side in $manifest.sides){
        $cell=$sheet.Range([string]$side.lengthCell)
        $source=$sheet.Range([string]$side.sourceCell)
        Check ($cell.HasFormula -and [string]$cell.Formula -ceq ('='+$side.sourceCell)) ('Side length must refer to General: '+$side.lengthCell)
        Check ($cell.Value2 -eq $source.Value2) ('Side length did not recalculate: '+$side.key)
    }
}
try{
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SharedRebarWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id
    if($ownedPid -in $existing){throw 'A separate Excel process is required.'}
    $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item('Config')
    # Нативные ссылки — отсутствующая возможность автора XLSX; формулы ключей не меняются.
    $sheet.Hyperlinks.Add($sheet.Range('L4'),'',"'Config'!A9",'','Подробнее') | Out-Null
    $sheet.Hyperlinks.Add($sheet.Range('R28'),'',"'Config'!A12",'','Подробнее') | Out-Null
    $sheet.Hyperlinks.Add($sheet.Range('L55'),'',"'Config'!A9",'','Подробнее') | Out-Null
    foreach($address in @('L4','R28','L55')){$sheet.Range($address).Font.Name='Arial';$sheet.Range($address).Font.Size=9}
    # Excel учитывает служебные поля OOXML при переводе ширины. Фиксируем итоговые размеры.
    $sheet.Columns.Item('I:U').ColumnWidth=17
    $sheet.Columns.Item(8).ColumnWidth=32
    Check (-not $book.HasVBProject) 'The example must have no VBA project.'
    Check ($book.Names.Item('rngLSectionGeometry').RefersToRange.Address($false,$false) -ceq $manifest.geometry) 'Geometry range changed.'
    Check ($book.Names.Item('rngRebarInput').RefersToRange.Address($false,$false) -ceq $manifest.rebar) 'Rebar range changed.'
    Check ($book.Names.Item('rngCircleGeometry').RefersToRange.Address($false,$false) -ceq $manifest.circleGeometry) 'Circle geometry range changed.'
    Check ([string]$sheet.Range('H16').Value2 -ceq 'Edges' -and [string]$sheet.Range('H64').Value2 -ceq 'Edges') 'Wrong edge block key.'
    $sheet.Calculate();Keys 'LSection';Lengths
    foreach($dimension in $manifest.dimensions){
        Check (-not $sheet.Range([string]$dimension.cell).HasFormula) ('General dimension must be an input: '+$dimension.key)
        Check ($sheet.Range([string]$dimension.cell).Value2 -eq $dimension.value) ('Wrong initial dimension: '+$dimension.key)
    }
    for($c=9;$c -le 21;$c++){Check ($sheet.Columns.Item($c).ColumnWidth -eq 17) 'I-U widths differ.'}
    foreach($area in @($manifest.geometry,$manifest.rebar,$manifest.circleGeometry)){
        $range=$sheet.Range($area)
        Check ($range.RowHeight -eq 15) 'Row heights differ.'
        Check ($range.WrapText -eq $false) 'Text wrapping must be disabled.'
        Check ($range.FormatConditions.Count -eq 0) 'Conditional formatting is not permitted.'
        Check ($range.Font.Name -ceq 'Arial' -and $range.Font.Size -eq 9) 'Wrong table font.'
        foreach($cell in $range.Cells){
            if($cell.MergeCells){continue}
            foreach($side in 7..10){Check ($cell.Borders.Item($side).LineStyle -eq -4115 -and $cell.Borders.Item($side).Weight -eq 2) ('Missing thin dashed border at '+$cell.Address())}
        }
    }
    foreach($area in $manifest.commentRanges){Check ($sheet.Range([string]$area).HorizontalAlignment -eq -4152) 'Comments must be right aligned.'}
    foreach($address in $manifest.inputCells){
        $cell=$sheet.Range([string]$address)
        Check ($cell.Interior.Color -eq 16777215 -and $cell.HorizontalAlignment -eq -4108) ('Wrong input format: '+$address)
    }
    foreach($item in $manifest.validations){
        $cell=$sheet.Range([string]$item.address)
        # Автор XLSX не поддерживает разрешение пустых значений; необязательные поля могут быть пустыми.
        $cell.Validation.IgnoreBlank=$true
        Check ($cell.Validation.Type -eq 3 -and $cell.Validation.Formula1 -ceq ('='+$item.name) -and $cell.Validation.IgnoreBlank -and ([string]$cell.Value2 -ceq '' -or $cell.Validation.Value)) ('Wrong independent list: '+$item.address)
    }
    foreach($area in @('J33:N40','J46:K53','N46:O53')){
        foreach($cell in $sheet.Range($area).Cells){$cell.Validation.IgnoreBlank=$true;Check $cell.Validation.IgnoreBlank 'Optional rebar inputs must allow blank values.'}
    }
    $savedInputs=@{}
    foreach($address in $manifest.inputCells){if($address -cne 'B7'){$savedInputs[$address]=$sheet.Range([string]$address).Value2}}
    $sheet.Range('B7').Value2='Circle';$sheet.Calculate();Keys 'Circle'
    foreach($address in $savedInputs.Keys){Check ($sheet.Range([string]$address).Value2 -eq $savedInputs[$address]) ('Type switch changed a manual input: '+$address)}
    $sheet.Range('B7').ClearContents() | Out-Null;$sheet.Calculate();Keys ''
    $sheet.Range('B7').Value2='LSection';$sheet.Calculate();Keys 'LSection'
    $sheet.Range('H19').Value2='R1.TopCheck';$sheet.Calculate();Keys 'LSection'
    $sheet.Range('H19').Value2='R1.Top';$sheet.Calculate()
    $sheet.Range('I19').Value2='R1: проверка';$sheet.Calculate();Keys 'LSection'
    $sheet.Range('I19').Value2='R1: сверху'
    $sheet.Range('H19').ClearContents() | Out-Null;$sheet.Calculate();Keys 'LSection'
    $sheet.Range('H19').Value2='R1.Top';$sheet.Calculate();Keys 'LSection'
    foreach($unit in @('mm','cm','m')){
        $sheet.Range('B5').Value2=$unit;$sheet.Calculate()
        foreach($address in $manifest.unitCells){
            $cell=$sheet.Range([string]$address)
            Check ($cell.HasFormula -and [string]$cell.Value2 -ceq $unit) ('Unit formula did not recalculate: '+$address)
        }
    }
    $sheet.Range('B5').Value2='mm'
    foreach($dimension in $manifest.dimensions){
        Set-NumberCell $sheet.Range([string]$dimension.cell) ([double]$dimension.value+50)
        $sheet.Calculate();Lengths
        Set-NumberCell $sheet.Range([string]$dimension.cell) ([double]$dimension.value)
        $sheet.Calculate();Lengths
    }
    Set-NumberCell $sheet.Range('J62') 400;$sheet.Calculate()
    Check ($sheet.Range('J67').HasFormula -and [Math]::Abs($sheet.Range('J67').Value2-400*[Math]::PI) -lt 1e-9) 'Circle perimeter did not recalculate.'
    Set-NumberCell $sheet.Range('J62') 300;$sheet.Calculate()
    Keys 'LSection'
    Check ([string]$sheet.Range('L4').Value2 -ceq 'Подробнее' -and [string]$sheet.Range('R28').Value2 -ceq 'Подробнее') 'Help hyperlinks did not calculate.'
    $book.Worksheets.Item('_Lists').Visible=0
    $sheet.Activate();$excel.ActiveWindow.Zoom=85;$excel.ActiveWindow.ScrollColumn=8;$excel.ActiveWindow.ScrollRow=1
    $book.Save();$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$sheet=$book.Worksheets.Item('Config')
    Check $book.ReadOnly 'Reopen must be read-only.'
    Check ($book.Worksheets.Item('_Lists').Visible -eq 0) 'Internal list sheet must be hidden.'
    Keys 'LSection';Lengths
    foreach($address in $manifest.unitCells){Check ($sheet.Range([string]$address).HasFormula -and [string]$sheet.Range([string]$address).Value2 -ceq 'mm') 'Unit formulas must survive reopening.'}
    $formulaErrors=$null
    try{$formulaErrors=$sheet.Range('H5:R67').SpecialCells(-4123,16)}catch{if($_.Exception.HResult -ne -2146827284){throw}}
    Check ($null -eq $formulaErrors) 'Excel formula error in final workbook.'
    foreach($item in $manifest.captureRanges){
        if($item.type){$sheet.Range('B7').Value2=[string]$item.type;$sheet.Calculate();Keys ([string]$item.type)}
        $sheet.Activate();$excel.ActiveWindow.Zoom=50;$excel.ActiveWindow.ScrollColumn=8;$excel.ActiveWindow.ScrollRow=1
        $source=$sheet.Range([string]$item.range);$source.CopyPicture(1,-4147) | Out-Null
        $scratch=$book.Worksheets.Add()
        try{
            $chart=$scratch.ChartObjects().Add(0,0,[double]$source.Width,[double]$source.Height)
            $chart.Activate();$chart.Chart.Paste()
            Check ($chart.Chart.Export((Join-Path $PSScriptRoot ($item.name+'_Native.png')),'PNG')) 'Native screenshot failed.'
        }finally{$scratch.Delete();$sheet.Activate()}
        if($item.type){$sheet.Range('B7').Value2='LSection';$sheet.Calculate();Keys 'LSection'}
    }
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid -and $ownedTicks){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
Check ((Get-FileHash -LiteralPath $main).Hash -ceq $mainHash -and $mainHash -ceq $manifest.mainWorkbookSHA256) 'Production workbook changed.'
[ordered]@{Passed=$true;Checks=$checks;File=$path;SHA256=(Get-FileHash -LiteralPath $path).Hash;NoVBA=$true;KeyFormulas=16;FormulaSideNames=16;Capacity=8;ActiveRows=@{LSection=8;Circle=1};FormulaSideLengths=9;GeneralDimensionChanges=@('B1','H1','B2','H2','D');FormulaKeyRenaming=$true;InputLengthUnits=@('mm','cm','m');Reopened=$true;WorkingWorkbookUnchanged=$true} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Acceptance.json') -Encoding UTF8
Write-Output ('NATIVE_MOCKUP_PASS: checks='+$checks+'; VBA=False; key/type/side changes recalculated; source unchanged.')
