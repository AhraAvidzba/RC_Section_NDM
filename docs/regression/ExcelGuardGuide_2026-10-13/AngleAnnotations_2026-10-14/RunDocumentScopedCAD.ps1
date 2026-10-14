# Native regression in new documents of the verified ordinary AutoCAD.
# Existing documents and the application are never closed or written.
$ErrorActionPreference='Stop';$root=(Get-Location).Path;$directory=$PSScriptRoot
$target=Join-Path $directory 'DocumentScopedCAD02'
if(Test-Path -LiteralPath $target){throw 'Existing CAD evidence must not be replaced.'}
New-Item -ItemType Directory -Path $target | Out-Null
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class ScopedCadWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
function Document-Snapshot($app){
    $items=@();foreach($doc in $app.Documents){$items+=,[ordered]@{Name=[string]$doc.Name;FullName=[string]$doc.FullName;Saved=[bool]$doc.Saved}}
    return ,@($items | Sort-Object Name)
}
function Invoke-CadReady([scriptblock]$action){
    $deadline=[DateTime]::UtcNow.AddSeconds(60)
    while($true){
        try{return ,(& $action)}catch{
            $inner=$_.Exception;while($inner.InnerException){$inner=$inner.InnerException}
            if($inner.HResult -notin @(-2147418111,-2147417846) -or [DateTime]::UtcNow -ge $deadline){throw}
            Start-Sleep -Milliseconds 500
        }
    }
}
$source=Join-Path $directory 'Candidate.xlsm';$sourceHash=(Get-FileHash -LiteralPath $source).Hash
$path=Join-Path $target 'CADProbe.xlsm';Copy-Item -LiteralPath $source -Destination $path
$protected=@('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt','workbook/output/RC_Section_NDM_execution_report.txt');$hashes=@{}
foreach($name in $protected){$hashes[$name]=(Get-FileHash -LiteralPath (Join-Path $root $name)).Hash}
$acad=$null;$excel=$null;$book=$null;$bootstrap=$null;$previous=$null;$excelPid=0;$excelTicks=0;$before=$null;$passed=$false
try{
    $acad=[Runtime.InteropServices.Marshal]::GetActiveObject('AutoCAD.Application.24.2')
    [uint32]$cadId=0;[void][ScopedCadWindow]::GetWindowThreadProcessId([IntPtr]$acad.HWND,[ref]$cadId)
    $cadProcess=Get-Process -Id $cadId
    if([string]$acad.FullName -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe' -or $cadId -ne 9344 -or -not $acad.GetAcadState().IsQuiescent){throw 'Unexpected or busy CAD application; leave it alone.'}
    $before=Document-Snapshot $acad;$originalNames=@($before | ForEach-Object {$_.Name})
    if($acad.Documents.Count){$previous=$acad.ActiveDocument}
    $visible=[bool]$acad.Visible
    $template=Join-Path $env:LOCALAPPDATA 'Autodesk/AutoCAD 2023/R24.2/rus/Template/acadiso.dwt'
    if(-not (Test-Path -LiteralPath $template)){throw 'Standard template missing.'}
    $bootstrap=Invoke-CadReady { $acad.Documents.Add($template) }
    if($bootstrap.Name -in $originalNames -or $acad.Documents.Count -ne $before.Count+1){throw 'Bootstrap ownership check failed.'}
    Write-Output ('VERIFIED_CAD: '+$cadId+'; originalDocuments='+$before.Count+'; ownBootstrap='+$bootstrap.Name)
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=1
    [uint32]$id=0;[void][ScopedCadWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $excelPid=[int]$id;$excelTicks=(Get-Process -Id $excelPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: '+$excelPid)
    $book=$excel.Workbooks.Open($path,0,$false);$excel.Calculation=-4135
    $result=[string]$excel.Run("'$($book.Name)'!modTestAutoCADContours.RunRealAutoCADContourFormatTests",$acad,$template)
    $result | Set-Content -LiteralPath (Join-Path $target 'NativeCAD.txt') -Encoding UTF8
    Write-Output (($result -split "`r?`n" | Where-Object {$_ -match '^TOTAL|^FAIL:|^NATIVE_DWG:'}) -join "`n")
    if($result -notmatch 'TOTAL[^\r\n]*failed=0\b' -or $result -match '(?m)^FAIL:'){throw 'Native CAD tests failed.'}
    $passed=$true
}catch{
    [ordered]@{ErrorMessage=$_.Exception.ToString();ErrorRecord=$_.ToString()} | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath (Join-Path $target 'Error.json') -Encoding UTF8
    throw
}finally{
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    if($excelPid){$p=Get-Process -Id $excelPid -ErrorAction SilentlyContinue;if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $excelTicks){Stop-Process -Id $excelPid -Force}}
    if($bootstrap){if($bootstrap.Name -notin $originalNames){Invoke-CadReady { $bootstrap.Close($false) }}}
    $originalUnchanged=$false
    if($acad -and $before){
        # Any leaked test document may be closed only after verifying its saved path.
        foreach($doc in @($acad.Documents)){
            if($doc.Name -notin $originalNames -and [string]$doc.FullName -like ($target+'\ContourFormat_*.dwg')){$doc.Close($false)}
        }
        if($previous){$previous.Activate()}
        $after=Document-Snapshot $acad
        $originalUnchanged=($before | ConvertTo-Json -Compress) -ceq ($after | ConvertTo-Json -Compress)
        if([bool]$acad.Visible -ne $visible){throw 'CAD visibility changed.'}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($acad)
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    foreach($name in $protected){if((Get-FileHash -LiteralPath (Join-Path $root $name)).Hash -cne $hashes[$name]){throw 'Production changed during native CAD tests.'}}
    if((Get-FileHash -LiteralPath $source).Hash -cne $sourceHash){throw 'Candidate changed.'}
    [ordered]@{Passed=$passed -and $originalUnchanged;NativeTestsPassed=$passed;CADPID=[int]$cadId;Executable='Autodesk AutoCAD 2023';OriginalDocumentsCount=$before.Count;OriginalDocumentNamesPathsAndSavedFlagsUnchanged=$originalUnchanged;UserDocumentNamesLogged=$false;ApplicationClosed=$false;PreviousActiveDocumentRestored=$true;SourceUnchanged=$true} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $target 'Run.json') -Encoding UTF8
    if(-not $originalUnchanged){throw 'Original CAD document state differs; inspect before taking further action.'}
}
