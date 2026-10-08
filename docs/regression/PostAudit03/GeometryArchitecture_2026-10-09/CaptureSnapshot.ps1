param([string]$Workbook = 'Before.xlsm', [string]$Output = 'BeforeSnapshot.tsv', [string]$Macro = 'DumpShapeArchitectureSnapshot', [ValidateSet('SP35','SP63')][string]$CrackCode='SP35')
$ErrorActionPreference='Stop'
$root=(Resolve-Path (Join-Path $PSScriptRoot '../../../..')).Path
$excel=$null; $book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $book=$excel.Workbooks.Open((Join-Path $PSScriptRoot $Workbook),0,$true)
    $name='modTestShapeArchitecture'
    $existing=@($book.VBProject.VBComponents | Where-Object Name -eq $name)
    foreach ($component in $existing) { $book.VBProject.VBComponents.Remove($component) }
    $component=$book.VBProject.VBComponents.Add(1); $component.Name=$name
    $body=[regex]::Match([IO.File]::ReadAllText((Join-Path $root ('tests/'+$name+'.bas'))),'(?ms)^Option Explicit.*').Value
    $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    $started=Get-Date
    if($Macro -eq 'DumpShapeCalculationSnapshot') { $result=[string]$excel.Run("'$($book.Name)'!$name.$Macro",$CrackCode) }
    else { $result=[string]$excel.Run("'$($book.Name)'!$name.$Macro") }
    [IO.File]::WriteAllText((Join-Path $PSScriptRoot $Output),$result,[Text.UTF8Encoding]::new($false))
    if ($result -match '(?m)^FAIL:') { throw $result.Split("`n")[-2] }
    Write-Output ('SNAPSHOT: '+$Output+'; rows='+($result -split "`n").Count+'; seconds='+((Get-Date)-$started).TotalSeconds)
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
