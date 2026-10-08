# Runs native contour acceptance in an owned Autodesk AutoCAD instance.
# Existing user/SOFiPLUS instances are never edited, closed, or reused.
param([string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
      [string]$ReportDirectory='docs/regression/PostAudit03/NativeCAD',
      [ValidateSet('RunRealAutoCADContourTests','RunRealAutoCADContourFormatTests')]
      [string]$Macro='RunRealAutoCADContourTests')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory=Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $root $SourceWorkbook) -Destination $path -Force
$class=(Get-ItemProperty -LiteralPath 'Registry::HKEY_CLASSES_ROOT\AutoCAD.Application.24.2\CLSID').'(default)'
$server=(Get-ItemProperty -LiteralPath "Registry::HKEY_CLASSES_ROOT\CLSID\$class\LocalServer32").'(default)'
if ($server -notlike 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe*') {throw "Unsafe CAD COM registration: $server"}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class PostAuditCADIdentity {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
}
'@
$existing=@(Get-Process acad -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$excel=$null; $book=$null; $acad=$null; $owned=$false; $pidValue=0
try {
    Write-Output ('CAD_ACTIVATE: ' + $server)
    $acad=New-Object -ComObject AutoCAD.Application.24.2
    $fullName=[string]$acad.GetType().InvokeMember('FullName',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null)
    if ([string]::IsNullOrWhiteSpace($fullName)) {throw 'AutoCAD FullName is unavailable through IDispatch.'}
    if ($fullName -notlike 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe') {throw "Not Autodesk AutoCAD: $fullName"}
    [uint32]$processId=0
    $window=$acad.GetType().InvokeMember('HWND',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null)
    [void][PostAuditCADIdentity]::GetWindowThreadProcessId([IntPtr][long]$window,[ref]$processId)
    $pidValue=[int]$processId
    $owned=($pidValue -gt 0 -and $existing -notcontains $pidValue)
    if (-not $owned) {throw "CAD instance is not owned by this test: PID=$pidValue"}
    $acad.Visible=$false
    Write-Output "CAD_OWNED: PID=$pidValue; executable=$fullName"
    # Готовим первый документ в управляющем COM-процессе до передачи
    # application в Excel: скрытый CAD без документа может зависнуть
    # на первом межпроцессном чтении Documents из VBA.
    $template=Join-Path $env:LOCALAPPDATA 'Autodesk/AutoCAD 2023/R24.2/rus/Template/acadiso.dwt'
    if (-not (Test-Path -LiteralPath $template)) {throw 'Standard AutoCAD test template is missing.'}
    $documents=$acad.GetType().InvokeMember('Documents',[Reflection.BindingFlags]::GetProperty,$null,$acad,$null)
    $bootstrap=$documents.Add($template)
    Write-Output ('CAD_BOOTSTRAP_DOCUMENT: ' + [string]$bootstrap.Name)
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open($path)
    # The saved candidate already contains the guarded optional CAD entrypoint.
    # Test it without replacing its VBA or saving fixture mutations.
    $result=[string]$excel.Run("'$($book.Name)'!modTestAutoCADContours.$Macro",$acad,$template)
    $result | Set-Content -LiteralPath (Join-Path $directory 'NativeCAD.txt') -Encoding UTF8
    [ordered]@{Executable=$fullName; OwnedPID=$pidValue; COMServer=$server; UserDocumentsUsed=$false} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'NativeCADIdentity.json') -Encoding UTF8
    Write-Output (($result -split "`r?`n" | Where-Object {$_ -match '^TOTAL|^FAIL:|^NATIVE_DWG:'}) -join "`n")
    if ($result -notmatch 'TOTAL[^\r\n]*failed=0\b' -or $result -match '(?m)^FAIL:') {throw 'Native CAD test failed.'}
} finally {
    if ($book) {try {$book.Close($false)} catch {}}
    if ($excel) {try {$excel.Quit()} catch {}; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    if ($acad) {
        if ($owned) {try {$acad.Quit()} catch {Write-Warning $_.Exception.Message}}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
