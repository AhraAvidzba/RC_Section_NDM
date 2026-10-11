# Read-only comparison: approved changes must match their current source files.
param(
    [string]$BaselinePath='docs/regression/GeometryDeclarative_2026-10-11/Baseline/RC_Section_NDM.xlsm',
    [string]$ActualPath='docs/regression/GeometryDeclarative_2026-10-11/BaselineInstrumented/RC_Section_NDM.xlsm',
    [string]$ReportPath='docs/regression/GeometryDeclarative_2026-10-11/BaselineInstrumented/CodeComparison.json',
    [string[]]$ApprovedModules=@('modTestShapeArchitecture')
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$paths=@((Resolve-Path -LiteralPath (Join-Path $root $BaselinePath)).Path,(Resolve-Path -LiteralPath (Join-Path $root $ActualPath)).Path)
$before=@($paths | ForEach-Object {(Get-FileHash -LiteralPath $_).Hash})
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class GeometryBaselineWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hWnd,out uint id);}'
function TextHash([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally {$sha.Dispose()}
}
function Normalize-IdentifierCase([string]$body) {
    # VBE changes identifier capitalization across the entire project at compile.
    # Keep literals, comments, punctuation and whitespace exact.
    return [regex]::Replace($body,'"(?:[^"]|"")*"|''[^\r\n]*|[\p{L}_][\p{L}\p{Nd}_]*', [Text.RegularExpressions.MatchEvaluator]{
        param($match)
        if ($match.Value.StartsWith('"') -or $match.Value.StartsWith("'")) {return $match.Value}
        return $match.Value.ToLowerInvariant()
    })
}
$excel=$null;$book=$null;$ownedId=0;$ownedTicks=0
$approved=@{}
$sourceBodies=@{};$actualBodies=@{}
foreach ($name in $ApprovedModules) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object {$_.BaseName -eq $name -and $_.Extension -in '.bas','.cls'})
    if ($files.Count -ne 1) {throw ('Approved source module is not unique: '+$name)}
    $body=[regex]::Match([IO.File]::ReadAllText($files[0].FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
    if (-not $body) {throw ('Approved source body is missing: '+$name)}
    $body=($body -replace "`r`n","`n").TrimEnd("`r","`n")
    $sourceBodies[$name]=$body
    $approved[$name]=[ordered]@{Type=$(if ($files[0].Extension -eq '.cls') {2} else {1});IdentifierCaseSHA256=(TextHash (Normalize-IdentifierCase $body));SourceSHA256=(Get-FileHash -LiteralPath $files[0].FullName).Hash}
}
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][GeometryBaselineWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedId=[int]$id;$ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    $inventories=@()
    foreach ($path in $paths) {
        $book=$excel.Workbooks.Open($path,0,$true)
        if (-not $book.ReadOnly) {throw 'Comparison workbook is not read-only.'}
        $inventory=@{}
        foreach ($component in $book.VBProject.VBComponents) {
            $module=$component.CodeModule;$body=''
            if ($module.CountOfLines) {$body=[string]$module.Lines(1,$module.CountOfLines)}
            $body=($body -replace "`r`n","`n").TrimEnd("`r","`n")
            if ($approved.ContainsKey([string]$component.Name)) {$actualBodies[[string]$component.Name]=$body}
            $inventory[[string]$component.Name]=[ordered]@{Type=[int]$component.Type;SHA256=(TextHash $body);IdentifierCaseSHA256=(TextHash (Normalize-IdentifierCase $body))}
        }
        $inventories+=,$inventory
        $book.Close($false);$book=$null
    }
    $changes=@()
    foreach ($name in @($inventories[0].Keys+$inventories[1].Keys | Sort-Object -Unique)) {
        $left=$inventories[0][$name];$right=$inventories[1][$name]
        if (-not $left -or -not $right -or $left.Type -ne $right.Type -or $left.SHA256 -cne $right.SHA256) {
            $caseOnly=($left -and $right -and $left.Type -eq $right.Type -and $left.IdentifierCaseSHA256 -ceq $right.IdentifierCaseSHA256)
            $matchesSource=($approved.ContainsKey($name) -and $right -and $right.Type -eq $approved[$name].Type -and $right.IdentifierCaseSHA256 -ceq $approved[$name].IdentifierCaseSHA256)
            $changes+=,[ordered]@{Name=$name;Baseline=$left;Actual=$right;IdentifierCaseOnly=[bool]$caseOnly;MatchesApprovedSource=[bool]$matchesSource;Allowed=($matchesSource -or $caseOnly)}
        }
    }
    $sourceChecks=@()
    foreach ($name in $approved.Keys) {
        $actual=$inventories[1][$name];$expected=$approved[$name]
        $matches=($actual -and $actual.Type -eq $expected.Type -and $actual.IdentifierCaseSHA256 -ceq $expected.IdentifierCaseSHA256)
        $differences=@()
        if (-not $matches -and $actual) {
            $sourceLines=(Normalize-IdentifierCase $sourceBodies[$name]) -split "`n"
            $actualLines=(Normalize-IdentifierCase $actualBodies[$name]) -split "`n"
            for ($i=0;$i -lt [Math]::Max($sourceLines.Count,$actualLines.Count) -and $differences.Count -lt 20;$i++) {
                if ($i -ge $sourceLines.Count -or $i -ge $actualLines.Count -or $sourceLines[$i] -cne $actualLines[$i]) {
                    $differences+=,[ordered]@{Line=($i+1);Source=$(if ($i -lt $sourceLines.Count) {$sourceLines[$i]});Actual=$(if ($i -lt $actualLines.Count) {$actualLines[$i]})}
                }
            }
        }
        $sourceChecks+=,[ordered]@{Name=$name;Expected=$expected;Matches=[bool]$matches;Differences=$differences}
    }
    $unexpected=@($changes | Where-Object {-not $_.Allowed})
    $sourceMismatches=@($sourceChecks | Where-Object {-not $_.Matches})
    for ($i=0;$i -lt $paths.Count;$i++) {if ((Get-FileHash -LiteralPath $paths[$i]).Hash -cne $before[$i]) {throw 'Read-only comparison changed a workbook.'}}
    $result=[ordered]@{Status=$(if ($unexpected.Count -or $sourceMismatches.Count) {'FAIL'} else {'PASS'});ExcelVersion=[string]$excel.Version;ExcelBuild=[string]$excel.Build;BaselineSHA256=$before[0];ActualSHA256=$before[1];Components=$inventories[0].Count;ActualComponents=$inventories[1].Count;Changes=$changes;SourceChecks=$sourceChecks;SourceMismatchCount=$sourceMismatches.Count;UnexpectedCount=$unexpected.Count;ReadOnly=$true}
    $result | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
    Write-Output ('BASELINE_CODE_COMPARISON: status='+$result.Status+'; components='+$result.Components+'; changes='+$changes.Count+'; unexpected='+$unexpected.Count)
    if ($unexpected.Count) {throw ('Unexpected code changes: '+(($unexpected | ForEach-Object Name) -join ', '))}
    if ($sourceMismatches.Count) {throw ('Approved sources differ: '+(($sourceMismatches | ForEach-Object Name) -join ', '))}
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedId) {
        $process=Get-Process -Id $ownedId -ErrorAction SilentlyContinue
        if ($process -and $process.ProcessName -eq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedId -Force}
    }
}
