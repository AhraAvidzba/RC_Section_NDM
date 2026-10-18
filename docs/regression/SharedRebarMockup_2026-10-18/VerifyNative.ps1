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
    for($i=0;$i -lt 6;$i++){
        $expected=$Prefix+'.'+[string]$sheet.Cells.Item(15+$i,8).Value2
        foreach($r in @([int]$manifest.mainRows[$i],[int]$manifest.extraRows[$i])){
            $cell=$sheet.Cells.Item($r,8)
            Check ($cell.HasFormula -and [string]$cell.Value2 -ceq $expected) ('Wrong formula key at '+$cell.Address())
        }
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
    $sheet.Hyperlinks.Add($sheet.Range('R22'),'',"'Config'!A12",'','Подробнее') | Out-Null
    foreach($address in @('L4','R22')){$sheet.Range($address).Font.Name='Arial';$sheet.Range($address).Font.Size=9}
    # Excel учитывает служебные поля OOXML при переводе ширины. Фиксируем итоговые размеры.
    $sheet.Columns.Item('I:U').ColumnWidth=17
    $sheet.Columns.Item(8).ColumnWidth=32
    Check (-not $book.HasVBProject) 'The example must have no VBA project.'
    Check ($book.Names.Item('rngLSectionGeometry').RefersToRange.Address() -ceq '$H$5:$L$20') 'Geometry range changed.'
    Check ($book.Names.Item('rngRebarInput').RefersToRange.Address() -ceq '$H$23:$R$43') 'Rebar range changed.'
    $sheet.Calculate();Keys 'LSection'
    for($c=9;$c -le 21;$c++){Check ($sheet.Columns.Item($c).ColumnWidth -eq 17) 'I-U widths differ.'}
    foreach($area in @('H5:L20','H23:R43')){
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
    foreach($area in @('L8:L10','L15:L20','O27:O32','R38:R43')){Check ($sheet.Range($area).HorizontalAlignment -eq -4152) 'Comments must be right aligned.'}
    foreach($address in $manifest.inputCells){
        $cell=$sheet.Range([string]$address)
        Check ($cell.Interior.Color -eq 16777215 -and $cell.HorizontalAlignment -eq -4108) ('Wrong input format: '+$address)
    }
    foreach($item in $manifest.validations){
        $cell=$sheet.Range([string]$item.address)
        Check ($cell.Validation.Type -eq 3 -and $cell.Validation.Formula1 -ceq ('='+$item.name) -and $cell.Validation.Value) ('Wrong independent list: '+$item.address)
    }
    $sheet.Range('J8').Value2='ExampleType';$sheet.Calculate();Keys 'ExampleType'
    $sheet.Range('J8').Value2='LSection'
    $sheet.Range('H15').Value2='TopCheck';$sheet.Calculate();Keys 'LSection'
    $sheet.Range('H15').Value2='Top';$sheet.Calculate()
    $savedSide=[string]$sheet.Range('I27').Value2
    $sheet.Range('I27').Value2='Слева';$sheet.Calculate()
    Check ([string]$sheet.Range('H27').Value2 -ceq 'LSection.Left' -and [string]$sheet.Range('H38').Value2 -ceq 'LSection.Left') 'Side selection did not update both keys.'
    $sheet.Range('I27').ClearContents() | Out-Null;$sheet.Calculate()
    Check ([string]$sheet.Range('H27').Value2 -ceq '' -and [string]$sheet.Range('H38').Value2 -ceq '') 'Blank side must leave both keys blank.'
    $sheet.Range('I27').Value2=$savedSide
    foreach($unit in @('mm','cm','m')){
        $sheet.Range('B5').Value2=$unit;$sheet.Calculate()
        foreach($address in $manifest.unitCells){
            $cell=$sheet.Range([string]$address)
            Check ($cell.HasFormula -and [string]$cell.Value2 -ceq $unit) ('Unit formula did not recalculate: '+$address)
        }
    }
    $sheet.Range('B5').Value2='mm'
    Set-NumberCell $sheet.Range('J19') 650;$sheet.Calculate()
    Check ($sheet.Range('J17').Value2 -eq 400 -and $sheet.Range('J20').Value2 -eq 800) 'Dependent side lengths did not recalculate.'
    Set-NumberCell $sheet.Range('J19') 600;$sheet.Calculate()
    Check ($sheet.Range('J17').Value2 -eq 350 -and $sheet.Range('J20').Value2 -eq 800) 'Initial side lengths were not restored.'
    Keys 'LSection'
    Check ([string]$sheet.Range('L4').Value2 -ceq 'Подробнее' -and [string]$sheet.Range('R22').Value2 -ceq 'Подробнее') 'Help hyperlinks did not calculate.'
    $book.Worksheets.Item('_Lists').Visible=0
    $sheet.Activate();$excel.ActiveWindow.Zoom=85;$excel.ActiveWindow.ScrollColumn=8;$excel.ActiveWindow.ScrollRow=1
    $book.Save();$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($path,0,$true);$sheet=$book.Worksheets.Item('Config')
    Check $book.ReadOnly 'Reopen must be read-only.'
    Check ($book.Worksheets.Item('_Lists').Visible -eq 0) 'Internal list sheet must be hidden.'
    Keys 'LSection'
    foreach($address in $manifest.unitCells){Check ($sheet.Range([string]$address).HasFormula -and [string]$sheet.Range([string]$address).Value2 -ceq 'mm') 'Unit formulas must survive reopening.'}
    $formulaErrors=$null
    try{$formulaErrors=$sheet.Range('H5:R43').SpecialCells(-4123,16)}catch{if($_.Exception.HResult -ne -2146827284){throw}}
    Check ($null -eq $formulaErrors) 'Excel formula error in final workbook.'
    foreach($item in @(@('Geometry','H4:L20'),@('Rebar','H22:R43'))){
        $source=$sheet.Range([string]$item[1]);$source.CopyPicture(1,-4147)
        $scratch=$book.Worksheets.Add()
        try{
            $chart=$scratch.ChartObjects().Add(0,0,[double]$source.Width,[double]$source.Height)
            $chart.Activate();$chart.Chart.Paste()
            Check ($chart.Chart.Export((Join-Path $PSScriptRoot ($item[0]+'_Native.png')),'PNG')) 'Native screenshot failed.'
        }finally{$scratch.Delete();$sheet.Activate()}
    }
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid -and $ownedTicks){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
Check ((Get-FileHash -LiteralPath $main).Hash -ceq $mainHash -and $mainHash -ceq $manifest.mainWorkbookSHA256) 'Production workbook changed.'
[ordered]@{Passed=$true;Checks=$checks;File=$path;SHA256=(Get-FileHash -LiteralPath $path).Hash;NoVBA=$true;KeyFormulas=12;FormulaKeyRenaming=$true;InputLengthUnits=@('mm','cm','m');Reopened=$true;WorkingWorkbookUnchanged=$true} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'Acceptance.json') -Encoding UTF8
Write-Output ('NATIVE_MOCKUP_PASS: checks='+$checks+'; VBA=False; key/type/side changes recalculated; source unchanged.')
