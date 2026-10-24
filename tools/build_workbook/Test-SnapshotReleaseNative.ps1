# ==========================================================================
# ДЛЯ ТЕСТОВ: компиляция и фактический экспорт книги без импорта исходников.
# ==========================================================================
# По умолчанию книга открывается только для чтения. Compile разрешает лишь
# компиляцию уже собранного проекта и сохранение в частном каталоге evidence.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [switch]$Compile
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$directory=[IO.Path]::GetFullPath((Join-Path $root $EvidenceDirectory))
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'A private evidence directory is required.'}
if($Compile -and -not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Compile must use a private workbook.'}
[void][IO.Directory]::CreateDirectory($directory)
$export=Join-Path $directory 'VBA_All_Code.txt'
if(Test-Path -LiteralPath $export){throw 'An existing native export cannot be overwritten.'}
$before=(Get-FileHash -LiteralPath $path).Hash
$previous=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
Add-Type 'using System;using System.Runtime.InteropServices;public static class SnapshotNativeOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr w,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$count=0
$text=New-Object Text.StringBuilder

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Проверяет PID и время старта; чужие экземпляры не закрываются.
function Test-OwnedSnapshotNativeExcel{
    if(-not $owned -or $owned -in $previous){return $false}
    $process=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks
}

try{
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][SnapshotNativeOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedSnapshotNativeExcel)){throw 'Excel ownership not established.'}
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($path,0,(-not [bool]$Compile))
    if([bool]$book.ReadOnly -eq [bool]$Compile){throw 'The requested workbook access mode was not obtained.'}
    $project=$book.VBProject
    if(-not $book.HasVBProject -or $project.VBComponents.Count -lt 1){throw 'The VBA project is absent.'}
    $app.VBE.ActiveVBProject=$project
    if($Compile){
        $command=$app.VBE.CommandBars.FindControl(1,578)
        if(-not $command){throw 'Compile command is unavailable.'}
        if($command.Enabled){$command.Execute()}
        if($command.Enabled){throw 'The VBA project did not compile.'}
        $book.Save()
    }
    [void]$text.AppendLine('VBA PROJECT EXPORT')
    [void]$text.AppendLine('Workbook: '+$book.Name)
    [void]$text.AppendLine('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    $count=[int]$project.VBComponents.Count
    for($index=1;$index -le $count;$index++){
        $component=$project.VBComponents.Item($index);$code=$component.CodeModule
        [void]$text.AppendLine('='*100)
        [void]$text.AppendLine('COMPONENT: '+$component.Name)
        [void]$text.AppendLine('TYPE: '+$component.Type)
        [void]$text.AppendLine('LINES: '+$code.CountOfLines)
        [void]$text.AppendLine('='*100)
        if($code.CountOfLines){[void]$text.AppendLine($code.Lines(1,$code.CountOfLines))}
        [void]$text.AppendLine()
    }
    if($project.VBComponents.Count -ne $count){throw 'The component count changed during export.'}
}finally{
    if(Test-OwnedSnapshotNativeExcel){if($book){$book.Close($false)};if($app){$app.Quit()}}
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
$after=(Get-FileHash -LiteralPath $path).Hash
if(-not $Compile -and $before -cne $after){throw 'A read-only export changed its source.'}
[IO.File]::WriteAllText($export,$text.ToString(),(New-Object Text.UTF8Encoding($true)))
$proof=[ordered]@{Status='PASS';Workbook=$WorkbookPath;BeforeSHA256=$before;WorkbookSHA256=$after;CompiledWithoutImport=[bool]$Compile;ImportedComponents=0;Components=$count;ExportSHA256=(Get-FileHash -LiteralPath $export).Hash;OwnedPID=$owned;StartTicks=[string]$ticks}
$proof|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $directory 'NativeRelease.json') -Encoding UTF8
$proof|ConvertTo-Json
