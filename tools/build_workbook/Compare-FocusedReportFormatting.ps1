# Compares the previous and current formatter in one isolated Excel instance.
param(
    [Parameter(Mandatory=$true)][string]$BaselineWorkbook,
    [Parameter(Mandatory=$true)][string]$LargeInputPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$FormatterFile='src/Excel/CExecutionReport.cls'
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source=(Resolve-Path -LiteralPath (Join-Path $root $BaselineWorkbook)).Path
$hash=(Get-FileHash -LiteralPath $source).Hash
$directory=[IO.Path]::GetFullPath((Join-Path $root $OutputDirectory))
if(-not $directory.StartsWith($root.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must stay within the workspace.' }
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$path=Join-Path $directory 'RC_Section_NDM.xlsm'
if(Test-Path -LiteralPath $path) { throw 'Existing evidence cannot be overwritten.' }
Copy-Item -LiteralPath $source -Destination $path
$method=@'
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Открывает в частной копии существующие функции форматирования отчёта для сравнения с исходной реализацией.
Public Function TestProbeFormat(ByVal kind As String, ByVal body As String) As String
    Select Case kind
        Case "Crack": TestProbeFormat = CompactCrackDiagnosticText(body)
        Case "State": TestProbeFormat = CompactStateSolutionDiagnosticText(body)
        Case "Capacity": TestProbeFormat = CompactCapacityDiagnosticText(body)
        Case "Block"
            mEnabled = True: Set mLines = New Collection: mStartTimer = Timer
            AddBlock "probe", body
            TestProbeFormat = CStr(mLines.Count)
            If mLines.Count > 1 Then TestProbeFormat = TestProbeFormat & "|" & CStr(mLines.Item(2))
    End Select
End Function
'@
$entry=@'
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пропускает одинаковый текст через исходный либо текущий formatter и возвращает результат или исходную ошибку.
Public Function TestRunReportFormatProbe(ByVal kind As String, ByVal body As String, ByVal before As Boolean) As String
    On Error GoTo Failed
    If before Then
        Dim oldReport As CTestExecutionReportBaseline
        Set oldReport = New CTestExecutionReportBaseline
        TestRunReportFormatProbe = oldReport.TestProbeFormat(kind, body)
    Else
        Dim newReport As CExecutionReport
        Set newReport = New CExecutionReport
        TestRunReportFormatProbe = newReport.TestProbeFormat(kind, body)
    End If
    Exit Function
Failed:
    TestRunReportFormatProbe = "RUNTIME ERROR: " & CStr(Err.Number) & "; " & Err.Description
End Function
'@
$excel=$null; $book=$null; $ownedPid=0; $ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class FocusedReportWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
    [uint32]$id=0
    [void][FocusedReportWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id; $ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly) { throw 'Comparison copy is read-only.' }
    $original=$book.VBProject.VBComponents.Item('CExecutionReport').CodeModule
    $baseline=$book.VBProject.VBComponents.Add(2); $baseline.Name='CTestExecutionReportBaseline'
    $baseline.CodeModule.AddFromString("' ==========================================================================`r`n' ДЛЯ ТЕСТОВ — исходный formatter для независимого сравнения.`r`n' ==========================================================================`r`n"+$original.Lines(1,$original.CountOfLines)+"`r`n"+$method)
    $formatter=(Resolve-Path -LiteralPath (Join-Path $root $FormatterFile)).Path
    $body=[regex]::Match([IO.File]::ReadAllText($formatter),'(?ms)^Option Explicit.*').Value
    $original.DeleteLines(1,$original.CountOfLines); $original.AddFromString($body+"`r`n"+$method)
    $book.VBProject.VBComponents.Item('modTestWorkbookInterface').CodeModule.AddFromString($entry)
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if($compile.Enabled) { $compile.Execute() }
    if($compile.Enabled) { throw 'Formatter comparison compilation failed.' }
    $large=(Resolve-Path -LiteralPath (Join-Path $root $LargeInputPath)).Path
    $markers=[regex]::Matches($original.Lines(1,$original.CountOfLines), 'ReplaceDiagnosticMarker translated, "(?<key>(?:""|[^"])*)",', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $allMarkers=(@($markers | ForEach-Object { $_.Groups['key'].Value.Replace('""','"') }) -join '0; ')
    if ($FormatterFile -ne 'src/Excel/CExecutionReport.cls' -and $markers.Count -ne 51) { throw 'Candidate marker coverage changed.' }
    $special='ultimatePath pureAxial redirected=LoadMultiplier'
    $russianSpecial='Силовая траектория UltimateStrain перенаправлена на LoadMultiplier.'
    $inputs=@('', " `r`n  `r`n", "initialGuess=warmStartSolver`r`ninitialGuess=directStateActiveSet`r`n  step=1  `r`n`r`n",
        "ultimatePath pureAxial redirected=LoadMultiplier`r`nprobe=1; lambda=0.95`r`n initialGuess=warmStartSolver `r`n",
        "Силовая траектория UltimateStrain перенаправлена на LoadMultiplier.`r`nultimate iter=1; rN=0; rLimit=0",
        $allMarkers, $allMarkers.ToUpperInvariant(), ("Кириллица, lambda="+[char]0x3bb+"; "+$allMarkers+"`r`n"+$allMarkers),
        $special, ($special+"`r`n"+$special+"`r`n"+$russianSpecial), ('prefix '+$special), ($special+' suffix'),
        (" `r`n  "+$special+"  `r`n `r`n"), ($special+"`n"+$special), $special.ToUpperInvariant(),
        [IO.File]::ReadAllText($large,[Text.Encoding]::Unicode))
    $records=@()
    foreach($kind in @('Crack','State','Capacity','Block')) {
        for($index=0; $index -lt $inputs.Count; $index++) {
            Write-Output ('REPORT_FORMAT_START: '+$kind+'/'+$index)
            $watch=[Diagnostics.Stopwatch]::StartNew()
            $old=[string]$excel.Run("'"+$book.Name+"'!modTestWorkbookInterface.TestRunReportFormatProbe",$kind,[string]$inputs[$index],$true)
            $oldSeconds=$watch.Elapsed.TotalSeconds
            $watch.Restart()
            $new=[string]$excel.Run("'"+$book.Name+"'!modTestWorkbookInterface.TestRunReportFormatProbe",$kind,[string]$inputs[$index],$false)
            $newSeconds=$watch.Elapsed.TotalSeconds
            if($old.StartsWith('RUNTIME ERROR:') -or $new.StartsWith('RUNTIME ERROR:')) { throw ('Format runtime: '+$old+'; '+$new) }
            if(-not [string]::Equals($old,$new,[StringComparison]::Ordinal)) { throw ('Formatted text changed: '+$kind+'/'+$index) }
            $sha=[Security.Cryptography.SHA256]::Create()
            try { $outputHash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::Unicode.GetBytes($new))).Replace('-','') }
            finally { $sha.Dispose() }
            $records+=@{Kind=$kind; Case=$index; InputLength=$inputs[$index].Length; OutputLength=$new.Length;
                ExactTextEqual=$true; OldSeconds=$oldSeconds; NewSeconds=$newSeconds; OutputSHA256=$outputHash}
            Write-Output ('REPORT_FORMAT_EQUAL: '+$kind+'/'+$index+'; old='+$oldSeconds+'; new='+$newSeconds)
        }
    }
    if((Get-FileHash -LiteralPath $source).Hash -cne $hash) { throw 'Source changed.' }
    [ordered]@{Passed=$true; SourceSHA256=$hash; CurrentFormatterSHA256=(Get-FileHash -LiteralPath $formatter).Hash;
        InputSHA256=(Get-FileHash -LiteralPath $large).Hash; Cases=$records; ProductionChanged=$false} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $directory 'Results.json') -Encoding UTF8
} finally {
    if($book) { $book.Close($false) }
    if($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {
            Stop-Process -Id $ownedPid -Force -ErrorAction SilentlyContinue
        }
    }
}
