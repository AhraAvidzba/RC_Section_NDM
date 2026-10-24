# ==========================================================================
# ДЛЯ ТЕСТОВ: единое оформление настроек и справки после полной сборки
# ==========================================================================
# Открывает только read-only копию в собственном Excel. Проверяет реальную
# проверку данных, все блоки справки, ссылки, группировку и пустые Results.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
      [switch]$AllowSavedResults)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'HelpOutline.ps1')
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
$directory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$hash=(Get-FileHash -LiteralPath $path).Hash
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class SettingsStyleTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Подтверждает принадлежность Excel по PID и времени старта перед закрытием.
function Test-OwnedExcel {
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Находит ровно один заголовок; неверная или повторная привязка — ошибка проверки.
function Test-FindRow($Sheet,[string]$Title) {
    $last=[int]$Sheet.UsedRange.Row+[int]$Sheet.UsedRange.Rows.Count-1
    $values=$Sheet.Range('A1:A'+$last).Value2
    $found=@(for($row=1;$row -le $last;$row++){if([string]$values[$row,1] -ceq $Title){$row}})
    if($found.Count -ne 1){throw ('Ambiguous/missing heading: '+$Title)}
    return [int]$found[0]
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сохраняет нативный вид диапазона для локального визуального контроля.
function Export-TestView($Sheet,[string]$Address,[string]$Name) {
    $Sheet.Activate();$range=$Sheet.Range($Address);$range.CopyPicture(1,-4147)
    $scratch=$app.Workbooks.Add()
    try{
        $chart=$scratch.Worksheets.Item(1).ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if(-not $chart.Chart.Export((Join-Path $directory $Name),'PNG')){throw 'Native render failed.'}
    }finally{if(Test-OwnedExcel){$scratch.Close($false)}}
}

try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SettingsStyleTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedExcel)){throw 'Excel ownership was not established.'}
    Write-Output ('OWNED_EXCEL: id='+$owned+'; utcStartTicks='+$ticks)
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1;$app.EnableEvents=$true
    $book=$app.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Expected read-only verification.'}
    $updateButton=$book.Worksheets.Item('Расчет').Shapes.Item('btnUpdateSectionPlot')
    if([string]$updateButton.TextFrame.Characters().Text -cne 'Обновить схему и результаты' -or [string]$updateButton.OnAction -notmatch '(^|!)UpdateSectionPlot$'){throw 'Snapshot update button caption/action differs.'}
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка')
    $row=Test-FindRow $config 'Mesh.StepMode';$mode=$config.Cells.Item($row,2)
    $options=@('Auto 0.05','Auto 0.1','Auto 0.2','User');$separator=[string]$app.International(5)
    if([string]$mode.Value2 -cne 'Auto 0.1' -or $mode.Validation.Type -ne 3 -or -not $mode.Validation.InCellDropdown){throw 'Mode default/dropdown is wrong.'}
    if([string]$mode.Validation.Formula1 -cne ($options -join $separator)){throw 'Mode list uses wrong locale separator.'}
    $validation=@()
    foreach($option in $options){
        $mode.Value2=$option;$accepted=[bool]$mode.Validation.Value
        if(-not $accepted){throw ('Valid option rejected: '+$option)}
        $validation+=,[ordered]@{Value=$option;Accepted=$accepted}
    }
    foreach($invalid in @('Wrong',($options -join ','))){$mode.Value2=$invalid;if($mode.Validation.Value){throw ('Invalid option accepted: '+$invalid)}}
    $mode.Value2='Auto 0.1'
    $edge=$config.Cells.Item($row,5).Borders.Item(10)
    if($edge.LineStyle -ne -4115 -or $edge.Weight -ne 2 -or $edge.Color -ne 0){throw 'Missing dotted right border.'}
    $snapshotStyles=@()
    $referenceRow=Test-FindRow $config 'General.WorstCombinationCriterion'
    foreach($key in @('Results.SnapshotStorage','Results.SnapshotName')){
        $settingRow=Test-FindRow $config $key
        for($column=1;$column -le 5;$column++){
            $cell=$config.Cells.Item($settingRow,$column);$reference=$config.Cells.Item($referenceRow,$column)
            foreach($field in @('HorizontalAlignment','VerticalAlignment','WrapText','NumberFormat')){
                if($cell.$field -cne $reference.$field){throw ('New setting style differs: '+$key+' column '+$column+' '+$field)}
            }
            foreach($field in @('Name','Size','Bold','Italic','Color','Underline')){
                if($cell.Font.$field -cne $reference.Font.$field){throw ('New setting font differs: '+$key+' column '+$column+' '+$field)}
            }
            foreach($index in @(7,8,9,10)){
                $border=$cell.Borders.Item($index);$refBorder=$reference.Borders.Item($index)
                if($border.LineStyle -ne $refBorder.LineStyle -or $border.Weight -ne $refBorder.Weight -or $border.Color -ne $refBorder.Color){throw ('New setting border differs: '+$key+' column '+$column+' edge '+$index+'; actual='+$border.LineStyle+','+$border.Weight+','+$border.Color+'; expected='+$refBorder.LineStyle+','+$refBorder.Weight+','+$refBorder.Color)}
            }
        }
        $input=$config.Cells.Item($settingRow,2);$edge=$config.Cells.Item($settingRow,5).Borders.Item(10)
        if($key -ceq 'Results.SnapshotStorage'){
            if($input.Validation.Type -ne 3 -or -not $input.Validation.InCellDropdown -or [string]$input.Validation.Formula1 -cne (@('Excel','File') -join $separator)){throw 'Snapshot storage is not a native two-item inline list.'}
            $savedValue=$input.Value2
            foreach($option in @('Excel','File')){$input.Value2=$option;if(-not $input.Validation.Value){throw ('Storage option rejected: '+$option)}}
            $input.Value2='Excel,File';if($input.Validation.Value){throw 'Combined comma-separated storage string accepted as an item.'}
            $input.Value2=$savedValue
        }else{
            $hasValidation=$false;try{$hasValidation=($input.Validation.Type -eq 3)}catch{}
            if($hasValidation){throw 'Snapshot name must be free text, not a list.'}
        }
        if($input.Interior.Color -ne 15135487 -or $input.HorizontalAlignment -ne -4108){throw ('New setting input style differs: '+$key)}
        if($edge.LineStyle -ne -4115 -or $edge.Weight -ne 2 -or $edge.Color -ne 0){throw ('New setting right border missing: '+$key)}
        $snapshotStyles+=,[ordered]@{Key=$key;Input='RGB(255,242,230)/center';CellsCompared=5;BordersCompared=20;RightBorder='dotted/black/thin'}
    }
    $snapshotStyles|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'SnapshotSettingStyles.json') -Encoding UTF8
    $inlineLists=@();$seenLists=@{}
    foreach($sheetName in @('Config','Сечения')){
        $sheet=$book.Worksheets.Item($sheetName);$validated=$null
        try{$validated=$sheet.UsedRange.SpecialCells(-4174)}catch{}
        if($null -eq $validated){continue}
        foreach($cell in $validated.Cells){
            if($cell.Validation.Type -ne 3){continue}
            $formula=[string]$cell.Validation.Formula1
            if($formula.StartsWith('=')){continue}
            $identity=$sheetName+'|'+$formula
            if($seenLists.ContainsKey($identity)){continue};$seenLists[$identity]=$true
            if($separator -cne ',' -and $formula.Contains(',') -and -not $formula.Contains($separator)){throw ('Comma-separated single-item validation: '+$sheetName+'!'+$cell.Address())}
            $options=@($formula -split [regex]::Escape($separator));$saved=$cell.Formula
            foreach($option in $options){$cell.Value2=$option;if(-not $cell.Validation.Value){throw ('Native list rejects its item: '+$sheetName+'!'+$cell.Address()+' '+$option)}}
            if($options.Count -gt 1){$cell.Value2=($options -join ',');if($cell.Validation.Value){throw ('Native list accepts the combined string: '+$sheetName+'!'+$cell.Address())}}
            $cell.Formula=$saved
            $inlineLists+=,[ordered]@{Sheet=$sheetName;Address=$cell.Address();Formula=$formula;Items=$options;NativeItemsAccepted=$true}
        }
    }
    $inlineLists|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $directory 'InlineValidationLists.json') -Encoding UTF8
    $links=0
    foreach($link in $config.Hyperlinks){
        $cell=$link.Range
        if([string]$cell.Value2 -cne 'Подробнее'){continue}
        if($cell.Font.Name -cne 'Arial' -or $cell.Font.Size -ne 9 -or $cell.Font.Color -ne 16711680 -or $cell.Font.Underline -ne 2 -or $cell.Font.Bold -or $cell.Font.Italic){throw ('Link font differs: '+$cell.Address())}
        $links++
    }
    $help.Outline.ShowLevels(8)
    $paragraphs=0;$blocks=@()
    foreach($item in (Get-SettingsInstructionCatalog)){
        $heading=Test-FindRow $help $item.Title
        if([Math]::Abs([double]$help.Rows.Item($heading).RowHeight-[double]$help.StandardHeight) -gt 0.1){throw ('Setting header height differs: '+$item.Key)}
        $lines=@($item.Lines | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
        for($index=0;$index -lt $lines.Count;$index++){
            $number=$help.Cells.Item($heading+$index+1,1);$text=$help.Cells.Item($heading+$index+1,2)
            if($number.Value2 -isnot [double] -or $number.Value2 -ne ($index+1) -or $number.HorizontalAlignment -ne -4152 -or $number.NumberFormat -cne '0'){throw ('Paragraph numbering differs: '+$item.Key+' '+($index+1))}
            if($number.Font.Name -cne 'Arial' -or $number.Font.Size -ne 10 -or $number.Font.Color -ne 0 -or $number.Font.Bold -or $number.Font.Italic){throw ('Number style differs: '+$item.Key)}
            if([string]$text.Value2 -cne [string]$lines[$index]){throw ('Help content differs from catalog: '+$item.Key)}
            $height=[double]$text.EntireRow.RowHeight
            $text.EntireRow.AutoFit() | Out-Null
            if([double]$text.EntireRow.RowHeight -gt $height+0.1){throw ('Help paragraph is clipped: '+$item.Key+' '+($index+1))}
            $paragraphs++
        }
        $blocks+=,[ordered]@{Key=$item.Key;Row=$heading;Paragraphs=$lines.Count}
    }
    $pictures=@($help.Shapes | Where-Object Type -eq 13)
    if($pictures.Count -ne 76 -or @($pictures | Where-Object {$_.Height -le 0 -or $_.Width -le 0}).Count){throw 'Help pictures are missing or collapsed.'}
    $sections=Get-HelpMainRows $help;$last=Get-HelpContentLastRow $help
    $headers=@($sections.Values | Sort-Object);$cycles=0
    for($index=0;$index -lt $headers.Count;$index++){
        $first=[int]$headers[$index];$end=if($index+1 -lt $headers.Count){[int]$headers[$index+1]-1}else{$last}
        $body=$help.Rows.Item(($first+1).ToString()+':'+$end)
        $help.Rows.Item($first).ShowDetail=$false
        if($body.Hidden -ne $true){throw ('Partial collapse at '+$first)}
        $help.Rows.Item($first).ShowDetail=$true
        if($body.Hidden -ne $false){throw ('Partial expansion at '+$first)}
        $cycles++
    }
    $navigation=@()
    foreach($key in @('General.ReformatResultsEnabled','Mesh.StepMode','Results.SnapshotStorage','Results.SnapshotName')){
        $r=Test-FindRow $config $key;$target=Test-FindRow $help $key
        $help.Outline.ShowLevels(1);$config.Activate();$config.Cells.Item($r,5).Hyperlinks.Item(1).Follow()
        $start=[int]$sections[$script:HelpMainTitles[-1]]
        if($app.ActiveSheet.Name -cne 'Справка' -or $app.ActiveWindow.ScrollRow -ne $target -or $help.Rows.Item(($start+1).ToString()+':'+$last).Hidden -ne $false){throw ('Help navigation failed: '+$key)}
        $navigation+=,[ordered]@{Key=$key;ScrollRow=[int]$app.ActiveWindow.ScrollRow;WholeSectionExpanded=$true}
    }
    $app.EnableEvents=$false;$help.Outline.ShowLevels(8)
    Export-TestView $config 'A3:E19' 'Config.png'
    foreach($key in @('Mesh.StepMode','General.ReformatResultsEnabled')){
        $r=Test-FindRow $help $key
        $count=@($blocks | Where-Object Key -ceq $key)[0].Paragraphs
        Export-TestView $help ('A'+$r+':F'+($r+$count)) ($key+'.png')
    }
    $hasSavedResults=[bool]($book.Worksheets.Item('Results').Cells.Item(157,1).Value2 -or $book.Worksheets.Item('Results').Cells.Item(2,1).Value2)
    if(-not $AllowSavedResults -and $hasSavedResults){throw 'Unexpected saved calculation results in fresh template.'}
    $app.VBE.ActiveVBProject=$book.VBProject;$compile=$app.VBE.CommandBars.FindControl(1,578)
    if($null -eq $compile){throw 'VBA compile command unavailable.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){throw 'VBA compilation failed.'}
    $code=New-Object 'System.Collections.Generic.List[string]'
    $code.Add('VBA PROJECT EXPORT');$code.Add('Workbook: RC_Section_NDM.xlsm');$code.Add('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach($component in $book.VBProject.VBComponents){
        $module=$component.CodeModule;$code.Add(('='*100));$code.Add('COMPONENT: '+$component.Name);$code.Add('TYPE: '+$component.Type);$code.Add('LINES: '+$module.CountOfLines);$code.Add(('='*100))
        if($module.CountOfLines){$code.Add([string]$module.Lines(1,$module.CountOfLines))}
    }
    [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),($code -join "`r`n"),[Text.UTF8Encoding]::new($true))
    [ordered]@{Status='PASS';WorkbookSHA256=$hash;ListSeparator=$separator;Validation=$validation;InvalidValuesRejected=2;RightBorder='dotted/black/thin';LinkFontsChecked=$links;HelpBlocks=$blocks;NumberedParagraphs=$paragraphs;HelpPictures=$pictures.Count;OutlineCycles=$cycles;Navigation=$navigation;UpdateButtonCaption=[string]$updateButton.TextFrame.Characters().Text;FreshResults=(-not $hasSavedResults);SavedResultsAllowed=[bool]$AllowSavedResults;VBACompiled=$true;OwnedPID=$owned;StartTicks=[string]$ticks} |
        ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directory 'SettingsHelpStyle.json') -Encoding UTF8
    Write-Output ('SETTINGS_HELP_STYLE_PASS: blocks='+$blocks.Count+'; paragraphs='+$paragraphs+'; links='+$links+'; pictures='+$pictures.Count)
}catch{
    Write-Output $_.ScriptStackTrace
    throw
}finally{
    if(Test-OwnedExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if(Test-OwnedExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Read-only verification changed the workbook.'}
