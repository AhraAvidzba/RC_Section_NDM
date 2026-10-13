$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$relative='docs/regression/ExcelGuardGuide_2026-10-13/RecipeArguments'
$directory=Join-Path $root $relative
if(Test-Path -LiteralPath (Join-Path $directory 'Before.xlsm')){throw 'Workbook baseline exists; inspect before replacing.'}
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$lock=$null
try {$lock=[IO.File]::Open($production,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::None)} catch [IO.IOException] {}
if($lock){$lock.Dispose()} else {
    # User authorized saving and closing this particular workbook earlier.
    # Keep every other workbook and application setting as it is.
    $book=[Runtime.InteropServices.Marshal]::BindToMoniker($production)
    try {
        if([string]$book.FullName -cne $production){throw 'Unexpected workbook binding.'}
        if(-not $book.Application.Ready -or $book.ReadOnly){throw 'Leave the unavailable/read-only workbook alone.'}
        [ordered]@{FullName=[string]$book.FullName;Saved=[bool]$book.Saved;OtherWorkbooks=$book.Application.Workbooks.Count-1} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'LiveStateBeforeClose.json') -Encoding UTF8
        if(-not $book.Saved){$book.Save()}
        $book.Close($false)
    } finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
}
$beforeHash=(Get-FileHash -LiteralPath $production).Hash
foreach($name in @('Before.xlsm','LiveBefore.xlsm','Candidate.xlsm')){
    Copy-Item -LiteralPath $production -Destination (Join-Path $directory $name)
}
[ordered]@{BeforeSHA256=$beforeHash;AllCodeSHA256=(Get-FileHash -LiteralPath 'workbook/output/VBA_All_Code.txt').Hash;ExecutionReportSHA256=(Get-FileHash -LiteralPath 'workbook/output/RC_Section_NDM_execution_report.txt').Hash;PreparedAt=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Before.json') -Encoding UTF8
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Export-Audit03VbaSnapshot.ps1 -WorkbookPath "$relative/LiveBefore.xlsm" -OutputPath "$relative/LiveNative.txt"
if($LASTEXITCODE -ne 0){throw 'Initial native snapshot failed.'}
& "$root/tools/build_workbook/Update-PrivateWorkbookCode.ps1" -WorkbookPath "$relative/Candidate.xlsm" -ModuleNames @('CSectionRecipeCircle','CSectionRecipeRectSet','CSectionRecipeRoundedRectangle','CSectionRecipeHollowRectangle','modSectionRecipeApiGuide')
& "$root/tools/build_workbook/Update-RecipeAuthorGuide.ps1" -WorkbookPath "$relative/Candidate.xlsm"
& 'C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe' -X utf8 tools/build_workbook/Restore-ExcelGuardGuideView.py $relative
if($LASTEXITCODE -ne 0){throw 'View preservation failed.'}
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Export-Audit03VbaSnapshot.ps1 -WorkbookPath "$relative/Candidate.xlsm" -OutputPath "$relative/NativeVBA.txt"
if($LASTEXITCODE -ne 0){throw 'Candidate native snapshot failed.'}
$exampleDirectory=Join-Path $directory 'ExampleFixture'
New-Item -ItemType Directory -Path $exampleDirectory | Out-Null
Copy-Item -LiteralPath (Join-Path $directory 'Candidate.xlsm') -Destination (Join-Path $exampleDirectory 'ExampleCompile.xlsm')
& "$root/tools/build_workbook/Update-PrivateWorkbookCode.ps1" -WorkbookPath "$relative/ExampleFixture/ExampleCompile.xlsm" -ModuleNames @('CSectionRecipeRoundedTrapezoid') -ExampleFiles @('docs/geometry_refactor/examples/CSectionRecipeRoundedTrapezoid.cls')
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools/build_workbook/Export-Audit03VbaSnapshot.ps1 -WorkbookPath "$relative/ExampleFixture/ExampleCompile.xlsm" -OutputPath "$relative/ExampleFixture/NativeVBA.txt"
if($LASTEXITCODE -ne 0){throw 'Example native snapshot failed.'}
if((Get-FileHash -LiteralPath $production).Hash -cne $beforeHash){throw 'Production changed during private preparation.'}
Write-Output 'ALL_RECIPE_COMMENTS_READY: four production recipes and external example compiled; production unchanged.'
