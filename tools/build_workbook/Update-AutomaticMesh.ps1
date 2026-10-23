# Добавляет новую настройку и её справку только в изолированную копию книги.
# Существующие данные сохраняются; рабочий VBA не содержит одноразовой миграции.
param([string]$EvidenceDirectory='docs/regression/AutomaticMesh_2026-10-23')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$script:WorkbookLayout=Read-WorkbookLayout
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
if (-not $directory.StartsWith((Join-Path $root 'docs/regression/'),[StringComparison]::OrdinalIgnoreCase)) {throw 'Evidence must stay inside docs/regression.'}
if (Test-Path -LiteralPath $directory) {throw 'Existing evidence must not be overwritten.'}
New-Item -ItemType Directory -Path $directory | Out-Null
$source=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination (Join-Path $directory 'Before.xlsm')
Copy-Item -LiteralPath $source -Destination $path
$printAreas=@(Get-WorkbookPrintAreas $path)
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class AutoMeshWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$beforePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

# Проверяет принадлежность процесса перед изменениями и закрытием COM-книг.
function Test-OwnedExcel {
    if (-not $owned -or $owned -in $beforePids) {return $false}
    $p=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $null -ne $p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# Находит строку ключа в первом столбце; номера ячеек не зашиты в обновление.
function Find-KeyRow($Sheet,[string]$Key) {
    $last=[int]$Sheet.UsedRange.Row+[int]$Sheet.UsedRange.Rows.Count-1
    $values=$Sheet.Range('A1:A'+$last).Value2
    for($r=1;$r -le $last;$r++){if([string]$values[$r,1] -ceq $Key){return $r}}
    throw ('Missing key: '+$Sheet.Name+'! '+$Key)
}

try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][AutoMeshWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if (-not (Test-OwnedExcel)) {throw 'Excel ownership was not established.'}
    Write-Output ('OWNED_EXCEL: id='+$owned+'; utcStartTicks='+$ticks)
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Candidate is read-only.'}
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка')
    $help.Outline.ShowLevels(8)
    $configRow=Find-KeyRow $config 'Mesh.StepX'
    $config.Range('A'+$configRow+':E'+$configRow).Insert(-4121) | Out-Null
    $config.Range('A'+($configRow+1)+':E'+($configRow+1)).Copy($config.Range('A'+$configRow+':E'+$configRow))
    $catalog=@((Get-SystemSettingsCatalog | Where-Object Name -eq 'GeometrySettings').Rows)
    foreach ($key in @('Mesh.StepMode','Mesh.StepX','Mesh.StepY','Mesh.BoundarySubdivisions')) {
        $definition=@($catalog | Where-Object {$_[0] -ceq $key})[0]
        $r=if ($key -eq 'Mesh.StepMode') {$configRow} else {Find-KeyRow $config $key}
        $config.Cells.Item($r,1).Value2=$key
        $config.Cells.Item($r,4).Value2=[string]$definition[3]
        if ($key -eq 'Mesh.StepMode') {
            $config.Cells.Item($r,2).Value2='Auto 0.1';$config.Cells.Item($r,3).Value2='-'
            $cell=$config.Cells.Item($r,2);$cell.Validation.Delete();$cell.Validation.Add(3,1,1,'Auto 0.05,Auto 0.1,Auto 0.2,User')
            $cell.Validation.IgnoreBlank=$false;$cell.Validation.InCellDropdown=$true;$cell.Validation.ShowError=$true
            $cell.Interior.Color=15135487;$cell.HorizontalAlignment=-4108
            $config.Cells.Item($r,5).Hyperlinks.Delete()
        }
        if ($key -eq 'Mesh.BoundarySubdivisions') {$config.Cells.Item($r,2).Value2=2}
    }
    $helperColumn=Get-WorkbookLayoutHelperColumn 'LoadCases'
    $helperRows=$book.Names.Item('rngLoadCombinations').RefersToRange.Rows.Count-1
    $config.Range($config.Cells.Item(2,$helperColumn),$config.Cells.Item($helperRows+1,$helperColumn)).Formula='=IF(ROW()-1>ROWS(rngLoadCombinations)-1,"",INDEX(rngLoadCombinations,ROW(),1))'
    $helpRow=Find-KeyRow $help 'Mesh.StepX'
    $modeDefinition=@(Get-SettingsInstructionCatalog | Where-Object Key -ceq 'Mesh.StepMode')[0]
    $lines=@($modeDefinition.Lines | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
    $inserted=$lines.Count+2 # Заголовок и обычная пустая разделительная строка соседних блоков.
    $links=@(foreach($sheet in $book.Worksheets){foreach($link in $sheet.Hyperlinks){
        $target=[string]$link.SubAddress
        if ($target -match '^''?Справка''?!(\$?)([A-Z]+)(\$?)(\d+)$' -and [int]$Matches[4] -ge $helpRow) {
            [pscustomobject]@{Link=$link;Target=("'Справка'!"+$Matches[1]+$Matches[2]+$Matches[3]+([int]$Matches[4]+$inserted))}
        }
    }})
    $help.Range([string]$helpRow+':'+($helpRow+$inserted-1)).Insert(-4121) | Out-Null
    foreach($record in $links){$record.Link.SubAddress=$record.Target}
    $template=$helpRow+$inserted
    $help.Range('A'+$template+':F'+$template).Copy($help.Range('A'+$helpRow+':F'+$helpRow))
    $help.Cells.Item($helpRow,1).Value2='Mesh.StepMode'
    $help.Rows.Item($helpRow).RowHeight=$help.Rows.Item($template).RowHeight
    for($i=0;$i -lt $lines.Count;$i++) {
        $r=$helpRow+$i+1
        $help.Range('A'+($template+1)+':F'+($template+1)).Copy($help.Range('A'+$r+':F'+$r))
        $help.Cells.Item($r,1).Value2=[string]($i+1);$help.Cells.Item($r,2).Value2=[string]$lines[$i]
        $help.Rows.Item($r).AutoFit() | Out-Null
    }
    $config.Hyperlinks.Add($config.Cells.Item($configRow,5),'',("'Справка'!A"+$helpRow),'Открыть справку по выбору шага сетки','Подробнее') | Out-Null
    foreach($key in @('Mesh.StepX','Mesh.StepY','Mesh.BoundarySubdivisions')) {
        $r=Find-KeyRow $help $key
        $item=@(Get-SettingsInstructionCatalog | Where-Object Key -ceq $key)[0]
        $paragraphs=@($item.Lines | Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})
        for($i=0;$i -lt $paragraphs.Count;$i++) {
            if ([string]$help.Cells.Item($r+$i+1,1).Value2 -ne [string]($i+1)) {throw ('Help paragraph count mismatch: '+$key)}
            $help.Cells.Item($r+$i+1,2).Value2=[string]$paragraphs[$i];$help.Rows.Item($r+$i+1).AutoFit() | Out-Null
        }
    }
    $changed=@(& git -C $root diff --name-only -- src tests | Where-Object {$_ -match '\.(bas|cls)$'})
    foreach($relative in $changed) {
        $file=Join-Path $root $relative;$name=[IO.Path]::GetFileNameWithoutExtension($file)
        $body=[regex]::Match([IO.File]::ReadAllText($file,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if (-not $body) {throw ('No module body: '+$relative)}
        $module=$book.VBProject.VBComponents.Item($name).CodeModule
        if ($module.CountOfLines) {$module.DeleteLines(1,$module.CountOfLines)}
        $module.AddFromString(($body -split '\r?\n') -join "`r`n")
    }
    $app.VBE.ActiveVBProject=$book.VBProject
    $compile=$app.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {throw 'VBA compilation failed.'}
    $app.Calculate() # Обновляет единицы и зависимые формулы, не запускает НДС и не меняет режим Excel.
    foreach($picture in $help.Shapes){if ($picture.Type -eq 13 -and $picture.Height -le 0) {throw ('Zero-height Help picture: '+$picture.Name)}}
    $book.Save()
    $code=New-Object 'System.Collections.Generic.List[string]'
    $code.Add('VBA PROJECT EXPORT');$code.Add('Workbook: RC_Section_NDM.xlsm');$code.Add('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach($component in $book.VBProject.VBComponents) {
        $module=$component.CodeModule;$code.Add(('='*100));$code.Add('COMPONENT: '+$component.Name);$code.Add('TYPE: '+$component.Type);$code.Add('LINES: '+$module.CountOfLines);$code.Add(('='*100))
        if($module.CountOfLines){$code.Add([string]$module.Lines(1,$module.CountOfLines))}
    }
    [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),($code -join "`r`n"),[Text.UTF8Encoding]::new($true))
    [ordered]@{ConfigRow=$configRow;HelpRow=$helpRow;HelpRowsAdded=$inserted;RemappedLinks=$links.Count;Modules=$changed;Default='Auto 0.1';BoundarySubdivisions=2;Compiled=$true} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Update.json') -Encoding UTF8
} catch {
    Write-Output $_.ScriptStackTrace
    throw
} finally {
    if (Test-OwnedExcel) {
        if($book){$book.Close($false)}
        if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if(Test-OwnedExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
    }
}
Restore-WorkbookPrintAreas $path $printAreas
if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Published workbook changed during preparation.'}
Write-Output ('AUTOMATIC_MESH_CANDIDATE: '+$path)
