# Выполняет адресные макросы в собственной копии baseline или текущего кода.
param(
    [string]$ReportDirectory = 'docs/regression/PostAudit03/Focused',
    [string[]]$Macros = @('modTestBatchCalculation.RunPostAudit03LifecycleTests'),
    [switch]$FrozenProduction,
    [string[]]$TestModules = @('modTestBatchCalculation'),
    [switch]$AllowFailure
)
$ErrorActionPreference = 'Stop'
$Macros = @($Macros | ForEach-Object {$_ -split ','})
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory = Join-Path $root $ReportDirectory
New-Item -ItemType Directory -Path $directory -Force | Out-Null
$target = Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath (Join-Path $root 'docs/regression/PostAudit03/Baseline/RC_Section_NDM.xlsm') -Destination $target -Force
$excel = $null; $book = $null; $records = @(); $failed = $false
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
    $book = $excel.Workbooks.Open($target)
    $project=$book.GetType().InvokeMember('VBProject',[Reflection.BindingFlags]::GetProperty,$null,$book,$null)
    if ($null -eq $project) {throw 'Workbook VBProject is unavailable.'}
    $components=$project.GetType().InvokeMember('VBComponents',[Reflection.BindingFlags]::GetProperty,$null,$project,$null)
    if ($null -eq $components) {throw 'Workbook VBComponents is unavailable.'}
    $files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
        Where-Object {$_.Extension -in '.cls', '.bas'})
    if ($FrozenProduction) {$files = @($files | Where-Object {$_.DirectoryName -eq (Join-Path $root 'tests') -and $TestModules -contains $_.BaseName})}
    foreach ($file in $files) {
        $component = $null
        for ($index=1;$index -le $components.Count;$index++) {
            $existing=$components.Item($index)
            if ([string]$existing.Name -eq $file.BaseName) {$component=$existing; break}
        }
        $body = [regex]::Match([IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8), '(?ms)^Option Explicit.*').Value
        if (-not $body) {throw "Missing VBA body: $($file.FullName)"}
        $type = 1; if ($file.Extension -eq '.cls') {$type = 2}
        if ($null -ne $component) {
            if ($component.CodeModule.CountOfLines -gt 0) {$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)}
        } else {
            $component = $components.Add($type); $component.Name = $file.BaseName
        }
        $component.CodeModule.AddFromString(($body -split "`r?`n") -join "`r`n")
    }
    if (-not $FrozenProduction) {
        Write-Output ([string]$excel.Run("'$($book.Name)'!MigrateSavedSectionContours"))
    }
    $book.Save()
    foreach ($macro in $Macros) {
        # Каждый макрос получает чистую COM-сессию и исходный сохраненный Config.
        # Направленные fixtures создают/удаляют листы; их состояние не переносится
        # в следующий тест через старый RCW-прокси Excel.
        if ($records.Count -gt 0) {
            try {$book.Close($false)} catch {Write-Warning ('Own fixture close: ' + $_.Exception.Message)}
            try {$excel.Quit()} catch {Write-Warning ('Own fixture quit: ' + $_.Exception.Message)}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
            $book = $null; $excel = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false; $excel.DisplayAlerts = $false; $excel.EnableEvents = $false; $excel.AutomationSecurity = 1
            $book = $excel.Workbooks.Open($target)
        }
        Write-Output ('RUN: ' + $macro)
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $result = [string]$excel.Run("'$($book.Name)'!$macro")
        $watch.Stop()
        $result | Set-Content -LiteralPath (Join-Path $directory ($macro + '.txt')) -Encoding UTF8
        $totals = [regex]::Matches($result, '(?im)^TOTAL[^\r\n]*passed=(\d+);?\s*failed=(\d+)')
        if ($totals.Count -eq 0) {throw "Macro has no total: $macro"}
        $last = $totals[$totals.Count - 1]
        $bad = [int]$last.Groups[2].Value
        if ($bad -gt 0 -or $result -match '(?m)^FAIL:') {$failed = $true}
        $record = [ordered]@{Macro = $macro; Passed = [int]$last.Groups[1].Value; Failed = $bad; Seconds = $watch.Elapsed.TotalSeconds}
        $records += $record
        Write-Output ($record | ConvertTo-Json -Compress)
    }
    [ordered]@{FrozenProduction = [bool]$FrozenProduction; SourceGitSHA = (& git -C $root rev-parse HEAD); Tests = $records; Failed = $failed} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'Summary.json') -Encoding UTF8
}
finally {
    if ($null -ne $book) {
        try {$book.Close($false)} catch {Write-Warning ('Test workbook close: ' + $_.Exception.Message)}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($null -ne $excel) {
        try {$excel.Quit()} catch {Write-Warning ('Own Excel quit: ' + $_.Exception.Message)}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
if ($failed -and -not $AllowFailure) {exit 1}
