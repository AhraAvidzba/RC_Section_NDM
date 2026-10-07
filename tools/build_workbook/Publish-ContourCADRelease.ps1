# Publishes the contour/CAD change only after native and source-equality gates.
# User inputs, Results, named anchors, widths and the execution report are guarded.
param(
    [string]$Directory='docs/regression/Performance/ContourCADRelease',
    [string]$CandidateWorkbook='docs/regression/Performance/CADCandidateV6/RC_Section_NDM.xlsm',
    [string]$NativeDirectory='docs/regression/Performance/NativeCADV4',
    [string]$NativeContractDirectory='docs/regression/Performance/ContourCADNativeContracts',
    [string]$FreshDirectory='docs/regression/Performance/ContourCADFreshCompiled'
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$evidence=Join-Path $root $Directory
$candidate=(Resolve-Path -LiteralPath (Join-Path $root $CandidateWorkbook)).Path
$native=Join-Path $root $NativeDirectory
$fresh=Join-Path $root $FreshDirectory
$expectedBook='DE653B7ACFA9EE4A97D1C7114A1067AADDC0229F118E50EE6849A016473DE282'
$expectedExport='5113A6C88851D65EEDF80F09F4DDAA3A15F6AF18CE40769CFAE8AB976BE504B0'
$expectedReport='D65B41D99B4E8969F5ADBB7196C9FB41989DEBB1287827F4F18D5D17B98ECCBC'
function Read-Json([string]$path) {return Get-Content -LiteralPath $path -Raw | ConvertFrom-Json}
function Assert-Hash([string]$path,[string]$expected) {
    if((Get-FileHash -LiteralPath $path).Hash -cne $expected) {throw "Artifact changed: $path"}
}
function Text-Hash([string]$text) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)))).Replace('-','')}
    finally {$sha.Dispose()}
}
function Book-Fingerprint($book) {
    $config=$book.Worksheets.Item('Config').UsedRange
    $names=@();foreach($name in $book.Names) {$names += ([string]$name.Name+'|'+[string]$name.RefersTo)}
    $sheet=$book.Worksheets.Item('Results');$results=$sheet.UsedRange;$data=$results.Value2
    $anchor=$book.Names.Item('rngNDMSectionContours').RefersToRange
    $data[($anchor.Row-$results.Row+1),($anchor.Column-$results.Column+1)]='RunID'
    $widths=@();for($column=1;$column -le 260;$column++) {$widths += [double]$sheet.Columns.Item($column).ColumnWidth}
    return [ordered]@{
        Inputs=(Text-Hash ($config.Formula | ConvertTo-Json -Depth 6 -Compress));
        Names=(Text-Hash ((@($names | Sort-Object)) -join "`n"));
        Widths=(Text-Hash ($widths | ConvertTo-Json -Compress));
        ResultsExceptHeader=(Text-Hash ($data | ConvertTo-Json -Depth 6 -Compress));
        ResultAddress=$results.Address()
    }
}
$nativeAcceptance=Read-Json (Join-Path $native 'Acceptance.json')
$nativeEnvironment=Read-Json (Join-Path $native 'Environment.json')
if(-not $nativeAcceptance.Passed -or -not $nativeAcceptance.ExactABGeometry -or $nativeAcceptance.Repetitions -lt 5) {throw 'Native CAD acceptance is incomplete.'}
Assert-Hash $nativeEnvironment.Sources.B $nativeEnvironment.SourceHashes.B
$nativeEquality=Read-Json (Join-Path $evidence 'NativeSourceEquality.json')
if(-not $nativeEquality.AllPassed -or $nativeEquality.ActualComponents -ne 126 -or $nativeEquality.RepositoryComponents -ne 121 -or @($nativeEquality.GeneratedDocumentChecks).Count -ne 5) {throw 'Measured CAD code equality is incomplete.'}
$candidateHash=(Get-FileHash -LiteralPath $candidate).Hash
$nativeContracts=Join-Path $root $NativeContractDirectory
$nativeContractReport=Get-Content -LiteralPath (Join-Path $nativeContracts 'NativeCAD.txt') -Raw
$nativeContractIdentity=Read-Json (Join-Path $nativeContracts 'NativeCADIdentity.json')
if($nativeContractReport -notmatch 'TOTAL_REAL_AUTOCAD_CONTOURS: passed=[1-9][0-9]*; failed=0' -or $nativeContractReport -match '(?m)^FAIL:' -or
    $nativeContractIdentity.Executable -cne 'C:\Program Files\Autodesk\AutoCAD 2023\acad.exe' -or $nativeContractIdentity.UserDocumentsUsed) {throw 'Final native contour contracts are incomplete.'}
Assert-Hash (Join-Path $nativeContracts 'RC_Section_NDM.xlsm') $candidateHash
$equality=Read-Json (Join-Path $evidence 'SourceEquality.json')
if(-not $equality.AllPassed -or $equality.ActualComponents -ne 126 -or $equality.RepositoryComponents -ne 121 -or @($equality.GeneratedDocumentChecks).Count -ne 5) {throw 'Source equality is incomplete.'}
foreach($check in $equality.Checks) {
    $files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object {$_.BaseName -eq $check.Component -and $_.Extension -in '.bas','.cls'})
    if($files.Count -ne 1) {throw 'Ambiguous source identity.'}
    Assert-Hash $files[0].FullName $check.SourceSHA256
}
Assert-Hash (Join-Path $PSScriptRoot 'Build-Workbook.ps1') $equality.BuildScriptSHA256
foreach($folder in @($evidence,$fresh)) {
    $validation=Read-Json (Join-Path $folder 'Validation.json')
    if($validation.Count -ne 27 -or @($validation | Where-Object {-not $_.Passed}).Count) {throw "Validation failed: $folder"}
    $ui=Get-Content -LiteralPath (Join-Path $folder 'ConfigPresentation.txt') -Raw
    if($ui -notmatch 'TOTAL_CONFIG_PRESENTATION: passed=1583; failed=0' -or $ui -match '(?m)^FAIL:' -or $ui -notmatch 'SOURCE_UNCHANGED: True') {throw "Presentation failed: $folder"}
}
$directed=Get-Content -LiteralPath (Join-Path $evidence 'Directed.txt') -Raw
foreach($total in @('TOTAL_POSTAUDIT03_CONTOURS: passed=231; failed=0','TOTAL_AUTOCAD_CONTOURS: passed=283; failed=0','TOTAL: passed=668; failed=0','TOTAL_AUDIT03_CONTOUR_ARC: passed=393; failed=0','TOTAL_AUDIT03_READ_LIFECYCLE: passed=3929; failed=0')) {
    if(-not $directed.Contains($total)) {throw "Directed gate missing: $total"}
}
if($directed -match '(?m)^FAIL:' -or $directed -notmatch 'SOURCE_UNCHANGED: True') {throw 'Directed gate failed.'}
$outputBook=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$outputExport=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$outputReport=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
Assert-Hash $outputBook $expectedBook;Assert-Hash $outputExport $expectedExport;Assert-Hash $outputReport $expectedReport
$excel=$null;$book=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    $book=$excel.Workbooks.Open($outputBook,0,$true);$before=Book-Fingerprint $book;$book.Close($false);$book=$null
    $book=$excel.Workbooks.Open($candidate,0,$true);$after=Book-Fingerprint $book
    $anchor=$book.Names.Item('rngNDMSectionContours').RefersToRange
    if([string]$anchor.Value2 -cne 'RunID') {throw 'Noncanonical contour header.'}
    $book.Close($false);$book=$null
    if(($before | ConvertTo-Json -Compress) -cne ($after | ConvertTo-Json -Compress)) {throw 'User state changed beyond the contour header.'}
}
finally {
    if($book) {try {$book.Close($false)} catch {}}
    if($excel) {try {$excel.Quit()} catch {};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
Assert-Hash $outputBook $expectedBook;Assert-Hash $outputExport $expectedExport;Assert-Hash $outputReport $expectedReport
$export=Join-Path $evidence 'VBA_All_Code.txt';$exportHash=(Get-FileHash -LiteralPath $export).Hash
Copy-Item -LiteralPath $candidate -Destination $outputBook
Copy-Item -LiteralPath $export -Destination $outputExport
Assert-Hash $outputBook $candidateHash;Assert-Hash $outputExport $exportHash;Assert-Hash $outputReport $expectedReport
[ordered]@{Published=$true;UTC=[DateTime]::UtcNow.ToString('o');WorkbookSHA256=$candidateHash;ExportSHA256=$exportHash;
    NativeABExact=$true;MeasuredWorkbookSHA256=$nativeEnvironment.SourceHashes.B;MeasuredCodeMatchesFinal=$true;SourceComponents=121;GeneratedComponents=5;UserReportPreserved=$true;UserReportSHA256=$expectedReport;
    InputsNamesWidthsResultsPreservedExceptHeader=$true;Fingerprint=$after;FullOnOffCompleted=$false} |
    ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $evidence 'Publication.json') -Encoding UTF8
Write-Output "PUBLISHED: $candidateHash; export=$exportHash; measured code and user state preserved."
