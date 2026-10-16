param([switch]$Retry)
$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=$PSScriptRoot
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt')
function Read-SharedHash([string]$path) {
    $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $digest=[Security.Cryptography.SHA256]::Create()
    try {[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}
    finally {$digest.Dispose();$stream.Dispose()}
}
# VBA автоматически меняет регистр одинаковых идентификаторов во всём проекте.
# Строковые литералы и комментарии сравниваются без каких-либо преобразований.
function Canonical-Code([string]$value) {
    [regex]::Replace($value, '"(?:[^"]|"")*"|''[^\r\n]*|[^"''\r\n]+', [Text.RegularExpressions.MatchEvaluator]{param($m)
        if($m.Value.StartsWith('"') -or $m.Value.StartsWith("'")){$m.Value}else{$m.Value.ToLowerInvariant()}
    })
}
$before=@{}
foreach($path in $protected){$before[$path]=Read-SharedHash (Join-Path $root $path)}
$candidate=Join-Path $directory 'Candidate'
New-Item -ItemType Directory -Path $candidate -Force | Out-Null
$path=Join-Path $candidate 'FaceDiameters.xlsm'
if(Test-Path -LiteralPath $path){if($Retry){Remove-Item -LiteralPath $path}else{throw 'Private copy already exists.'}}
Copy-Item -LiteralPath (Join-Path $root $protected[0]) -Destination $path
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class FaceProbeExcel{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][FaceProbeExcel]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$ownedPid)
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Read-only private copy.'}
    $original=@{}
    foreach($component in $book.VBProject.VBComponents){$original[$component.Name]=[string]$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)}
    $module=$book.VBProject.VBComponents.Item('modTestSectionRecipes').CodeModule
    $body=[IO.File]::ReadAllText((Join-Path $directory 'Probe.bas'),[Text.Encoding]::UTF8)
    $module.InsertLines($module.CountOfLines+1,$body)
    $recased=@()
    foreach($component in $book.VBProject.VBComponents){
        $actual=[string]$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
        if($component.Name -eq 'modTestSectionRecipes'){
            if(-not (Canonical-Code $actual).StartsWith((Canonical-Code $original[$component.Name]),[StringComparison]::Ordinal)){throw 'Existing test code changed.'}
        }elseif($actual -cne $original[$component.Name]){
            if((Canonical-Code $actual) -cne (Canonical-Code $original[$component.Name])){throw ('Unrelated code changed: '+$component.Name)}
            $recased+=,$component.Name
        }
    }
    if($book.VBProject.VBComponents.Count -ne $original.Count){throw 'Component inventory changed.'}
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($null -eq $compile){throw 'No compiler command.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){
        $pane=$excel.VBE.ActiveCodePane
        [int]$a=0;[int]$b=0;[int]$c=0;[int]$d=0
        $pane.GetSelection([ref]$a,[ref]$b,[ref]$c,[ref]$d)
        Write-Output ($pane.CodeModule.Parent.Name+':'+$a)
        Write-Output $pane.CodeModule.Lines([Math]::Max(1,$a-2),5)
        throw 'Compilation failed.'
    }
    $book.Save();$book.Close($false);$book=$null
    foreach($key in $protected){if((Read-SharedHash (Join-Path $root $key)) -cne $before[$key]){throw ('Protected file changed: '+$key)}}
    [ordered]@{Head=(& git rev-parse HEAD);Hashes=$before;CandidateSHA256=Read-SharedHash $path;ProbeSHA256=Read-SharedHash (Join-Path $directory 'Probe.bas');Components=$original.Count;IdentifierCaseOnly=$recased;WorkingCodeUnchanged=$true;ExistingTestCodePreserved=$true;Compiled=$true;ProtectedUnchanged=$true} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Encoding UTF8
    Write-Output 'PRIVATE_PROBE_READY'
}finally{
    if($book){$book.Close($false)}
    if($excel){try{$excel.Quit()}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){$p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}}
}
