# Compares actual read-only VBA export with repo bodies. VBE canonicalizes identifier
# casing; string literals and comments are preserved, not lowercased as code tokens.
param([string]$ExportPath='docs/regression/Performance/Release/VBA_All_Code.txt',
    [string]$ReportPath='docs/regression/Performance/Release/SourceEquality.json',
    [string]$ReferenceExportPath='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path=Join-Path $root $ExportPath
$export=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
$components=@{}
$pattern='(?ms)^={100}\r?\nCOMPONENT: ([^\r\n]+)\r?\nTYPE: (\d+)\r?\nLINES: (\d+)\r?\n={100}\r?\n(.*?)(?=^={100}\r?\nCOMPONENT: |\z)'
foreach($match in [regex]::Matches($export,$pattern)) {
    $components[$match.Groups[1].Value]=@{Type=[int]$match.Groups[2].Value;Body=$match.Groups[4].Value}
}
$reference=@{}
if($ReferenceExportPath) {
    $referenceText=[IO.File]::ReadAllText((Join-Path $root $ReferenceExportPath),[Text.Encoding]::UTF8)
    foreach($match in [regex]::Matches($referenceText,$pattern)) {
        $reference[$match.Groups[1].Value]=@{Type=[int]$match.Groups[2].Value;Body=$match.Groups[4].Value}
    }
}
function Normalize-Code([string]$body) {
    $body=[regex]::Match($body,'(?ms)^Option Explicit.*').Value.Replace("`r`n","`n").TrimEnd()
    if(-not $body) {throw 'Missing Option Explicit body.'}
    $builder=New-Object Text.StringBuilder
    foreach($line in $body.Split("`n")) {
        $inString=$false
        for($i=0;$i -lt $line.Length;$i++) {
            $char=$line[$i]
            if($char -eq '"') {$inString=-not $inString;[void]$builder.Append($char)}
            elseif(-not $inString -and $char -eq "'") {[void]$builder.Append($line.Substring($i));break}
            elseif($inString) {[void]$builder.Append($char)}
            else {[void]$builder.Append([char]::ToLowerInvariant($char))}
        }
        [void]$builder.Append("`n")
    }
    return $builder.ToString().TrimEnd()
}
$checks=@()
$documentChecks=@()
$files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
foreach($file in $files) {
    $actual=$components[$file.BaseName]
    if(-not $actual) {throw "Missing actual component: $($file.BaseName)"}
    $expectedType=1;if($file.Extension -eq '.cls') {$expectedType=2}
    $rawExpected=Normalize-Code ([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8))
    $expected=$rawExpected
    if($ReferenceExportPath) {
        if(-not $reference.ContainsKey($file.BaseName) -or $reference[$file.BaseName].Type -ne $expectedType) {throw 'Fresh source-built reference is incomplete.'}
        $expected=Normalize-Code $reference[$file.BaseName].Body
    }
    $observed=Normalize-Code $actual.Body
    $passed=$actual.Type -eq $expectedType -and $expected -ceq $observed
    $differences=@()
    if(-not $passed) {
        $expectedLines=$expected -split "`n";$observedLines=$observed -split "`n"
        for($i=0;$i -lt [Math]::Max($expectedLines.Count,$observedLines.Count);$i++) {
            $left='';$right='';if($i -lt $expectedLines.Count){$left=$expectedLines[$i]};if($i -lt $observedLines.Count){$right=$observedLines[$i]}
            if($left -cne $right) {$differences += [ordered]@{Line=$i+1;Source=$left;Book=$right};if($differences.Count -ge 20){break}}
        }
    }
    $checks += [ordered]@{Component=$file.BaseName;Type=$actual.Type;SourceLines=($rawExpected -split "`n").Count;SourceSHA256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash;RawTextSame=($rawExpected -ceq $observed);Passed=$passed;Differences=$differences}
}
if($ReferenceExportPath) {
    $sourceNames=@($files.BaseName)
    $extraNames=@($components.Keys | Where-Object {$_ -notin $sourceNames})
    $referenceExtraNames=@($reference.Keys | Where-Object {$_ -notin $sourceNames})
    if($extraNames.Count -ne $referenceExtraNames.Count) {throw 'Generated document component count differs.'}
    foreach($name in $extraNames) {
        $actual=$components[$name];$expected=$reference[$name]
        $passed=$null -ne $expected -and $actual.Type -eq 100 -and $expected.Type -eq 100
        if($passed) {
            if([string]::IsNullOrWhiteSpace($actual.Body) -or [string]::IsNullOrWhiteSpace($expected.Body)) {
                $passed=[string]::IsNullOrWhiteSpace($actual.Body) -and [string]::IsNullOrWhiteSpace($expected.Body)
            } else {$passed=(Normalize-Code $actual.Body) -ceq (Normalize-Code $expected.Body)}
        }
        $documentChecks += [ordered]@{Component=$name;Type=$actual.Type;Passed=$passed}
    }
}
$allPassed=@($checks | Where-Object {-not $_.Passed}).Count -eq 0 -and @($documentChecks | Where-Object {-not $_.Passed}).Count -eq 0
[ordered]@{ActualComponents=$components.Count;RepositoryComponents=$files.Count;ProductionClasses=@($files | Where-Object {$_.Extension -eq '.cls' -and $_.FullName.StartsWith((Join-Path $root 'src'))}).Count;AllPassed=$allPassed;Checks=$checks;GeneratedDocumentChecks=$documentChecks;BuildScriptSHA256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Build-Workbook.ps1') -Algorithm SHA256).Hash;ReferenceExport=$ReferenceExportPath;Normalization='CRLF/LF, trailing blank lines, VBE identifier case; string literals and comments unchanged. Optional native reference is an independently compiled fresh build from the listed source hashes, not ad-hoc numeric/comment rounding.'} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
if(-not $allPassed) {$checks | Where-Object {-not $_.Passed} | Format-Table;$documentChecks | Where-Object {-not $_.Passed} | Format-Table;throw 'Source/book code differs.'}
Write-Output "SOURCE/BOOK: $($files.Count) source and $($documentChecks.Count) generated components matched; no new production class."
