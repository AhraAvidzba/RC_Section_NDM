# ДЛЯ ТЕСТОВ: синхронизирует только один проверенный тестовый модуль в отдельной полной сборке.
# Пользовательская книга, значения, раскладка и рабочие методы VBA не редактируются.
# Внешний supervisor обязан ограничивать зависший COM/диалог компиляции.
[CmdletBinding()]
param(
    [string]$SourceWorkbook='docs/regression/ContourCellGeometry_2026-10-26_Run01/FreshBuild03/RC_Section_NDM.xlsm',
    [string]$EvidenceDirectory='docs/regression/ContourCellGeometry_2026-10-26_Run01/FreshBuild04_TestSync',
    [string]$TestSource='tests/modTestAutoCADContours.bas',
    [string]$OldTestSource='docs/regression/ContourCellGeometry_2026-10-26_Run01/DeferredContourSnapshotFix01/Candidate/tests/modTestAutoCADContours.bas',
    [string]$ExpectedNewTestSHA256='B0A8CF25094AF6190B9EB5C755D4D34C21879E3A795ED36537028FA632377977',
    [string]$ExpectedOldTestSHA256='C809FFB25187A01E4A96A61F393E5E64C558284043BB13BFA75C0B7A7117BBAA',
    [double]$MinimumFreeRAMGiB=4,
    [double]$MinimumFreeCommitGiB=8
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne 5){throw 'Use Windows PowerShell 5.1.'}
if($MinimumFreeRAMGiB -lt 4 -or $MinimumFreeCommitGiB -lt 8){throw 'Minimum reserve is 4 GiB RAM / 8 GiB commit.'}
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
function Resolve-Input([string]$path){if(-not [IO.Path]::IsPathRooted($path)){$path=Join-Path $root $path};return (Resolve-Path -LiteralPath $path).Path}
$source=Resolve-Input $SourceWorkbook;$newSource=Resolve-Input $TestSource;$oldSource=Resolve-Input $OldTestSource
if(-not [IO.Path]::IsPathRooted($EvidenceDirectory)){$EvidenceDirectory=Join-Path $root $EvidenceDirectory}
$directory=[IO.Path]::GetFullPath($EvidenceDirectory)
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Output must stay in docs/regression.'}
if(Test-Path -LiteralPath $directory){throw 'Use a fresh isolated output directory.'}
$sourceHash=(Get-FileHash -LiteralPath $source).Hash;$oldHash=(Get-FileHash -LiteralPath $oldSource).Hash;$newHash=(Get-FileHash -LiteralPath $newSource).Hash
if($oldHash -cne $ExpectedOldTestSHA256 -or $newHash -cne $ExpectedNewTestSHA256){throw 'Reviewed old/new test source SHA mismatch.'}
$oldBody=[regex]::Match([IO.File]::ReadAllText($oldSource,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
$newBody=[regex]::Match([IO.File]::ReadAllText($newSource,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
if(-not $oldBody -or -not $newBody){throw 'Test bodies must begin with Option Explicit.'}
$oldBody=[regex]::Replace($oldBody,'\r\n?',"`n").TrimEnd();$newBody=[regex]::Replace($newBody,'\r\n?',"`n").TrimEnd()
$reopen='    Set fixture = Application.Workbooks.Open(path): Set sheet = fixture.Worksheets("Results")'
$refresh='    Set geometryAnchor = fixture.Names.Item("rngNDMSectionGeometry").RefersToRange'
if($newBody -cne $oldBody.Replace($reopen+"`n",$reopen+"`n"+$refresh+"`n")){throw 'Only the reviewed one-line Range refresh is allowed.'}
[void][IO.Directory]::CreateDirectory($directory)
$output=Join-Path $directory 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $source -Destination $output
if((Get-FileHash -LiteralPath $output).Hash -cne $sourceHash){throw 'Initial isolated copy mismatch.'}
function Write-Proof([string]$name,$value){$value | ConvertTo-Json -Depth 9 | Set-Content -LiteralPath (Join-Path $directory $name) -Encoding UTF8}
function Exact-Process([int]$processID,[long]$ticks){
    if($processID -le 0 -or $ticks -le 0){return $false}
    $process=Get-Process -Id $processID -ErrorAction SilentlyContinue
    if(-not $process){return $false}
    try{return $process.ProcessName -ieq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks}catch{return $false}
}
function Inventory($book){return @($book.VBProject.VBComponents | ForEach-Object {([string]$_.Name)+'|'+([string]$_.Type)} | Sort-Object)}
function Body($component){$count=[int]$component.CodeModule.CountOfLines;if($count -eq 0){return ''};return [string]$component.CodeModule.Lines(1,$count)}
# Берёт только существующий чистый normalizer из проверенного source-equality helper.
# Его основная часть не выполняется: никаких других компонентов/книг не читается.
$equalityHelper=Join-Path $PSScriptRoot 'TestVerifySourceEquality.ps1'
$equalityText=[IO.File]::ReadAllText($equalityHelper,[Text.Encoding]::UTF8)
$canonicalCS=[regex]::Match($equalityText,"(?s)using System;.*?(?=\r?\n'@)").Value.Replace('PostAuditVbaCanonicalText','FinalTestSyncCanonicalText')
if(-not $canonicalCS){throw 'Existing VBA comparison normalizer unavailable.'}
if(-not ('FinalTestSyncCanonicalText' -as [type])){Add-Type -TypeDefinition $canonicalCS}
function Canonical([string]$body){return [FinalTestSyncCanonicalText]::Normalize([regex]::Replace($body,'\r\n?',"`n").TrimEnd())}
if(-not ('NdmFinalTestSyncOwner' -as [type])){Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class NdmFinalTestSyncOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'}
$existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {[ordered]@{PID=$_.Id;StartTicksUTC=[string]$_.StartTime.ToUniversalTime().Ticks}})
$priorIDs=@($existing | ForEach-Object {$_.PID})
$app=$null;$book=$null;$component=$null;$compile=$null;$pane=$null;$excelID=0;$ticks=0L;$owned=$false;$failure=$null;$saved=$false;$beforeMatches=$false;$afterMatches=$false;$inventorySame=$false;$beforeInventory=@();$afterInventory=@();$closed=$false
Write-Proof 'Input.json' ([ordered]@{Source=$source;SourceSHA256=$sourceHash;OldTestSource=$oldSource;OldTestSHA256=$oldHash;NewTestSource=$newSource;NewTestSHA256=$newHash;EqualityNormalizerSHA256=(Get-FileHash -LiteralPath $equalityHelper).Hash;OneLineSourceDeltaVerified=$true;ExistingExcel=$existing;OnlyComponent='modTestAutoCADContours';MacrosRun=$false;CellWritesPerformed=$false;ProductionSourceEdited=$false})
try{
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $ram=[double]$memory.AvailableMBytes/1024;$commit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
    $ready=$ram -ge $MinimumFreeRAMGiB -and $commit -ge $MinimumFreeCommitGiB
    Write-Proof 'MemoryBeforeExcel.json' ([ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');FreeRAMGiB=$ram;FreeCommitGiB=$commit;RequiredRAMGiB=$MinimumFreeRAMGiB;RequiredCommitGiB=$MinimumFreeCommitGiB;Passed=$ready})
    if(-not $ready){throw 'Insufficient memory; no Excel started.'}
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeID=0;[void][NdmFinalTestSyncOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeID)
    $excelID=[int]$nativeID;$ticks=(Get-Process -Id $excelID).StartTime.ToUniversalTime().Ticks
    if($excelID -in $priorIDs){throw 'Pre-existing Excel must not be used.'}
    $owned=$true
    Write-Proof 'Ownership.json' ([ordered]@{ExcelPID=$excelID;ExcelStartTicksUTC=[string]$ticks;SourceSHA256=$sourceHash;OwnOutput=$output;MacrosRun=$false;UserDocumentsUsed=$false})
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AskToUpdateLinks=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($output,0,$false)
    if($book.ReadOnly){throw 'Only the isolated output copy must be writable.'}
    $beforeInventory=Inventory $book
    if($beforeInventory.Count -ne 181){throw 'Expected exactly181 existing components; no add/remove operation allowed.'}
    $component=$book.VBProject.VBComponents.Item('modTestAutoCADContours')
    if([int]$component.Type -ne 1){throw 'Expected existing standard test module.'}
    $beforeBody=Body $component
    [IO.File]::WriteAllText((Join-Path $directory 'TargetModule.Before.bas'),$beforeBody,(New-Object Text.UTF8Encoding($true)))
    $beforeMatches=(Canonical $beforeBody) -ceq (Canonical $oldBody)
    if(-not $beforeMatches){throw 'Target module is not the reviewed C809 source before patch.'}
    $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
    $component.CodeModule.AddFromString(($newBody -split "`n") -join "`r`n")
    $app.VBE.ActiveVBProject=$book.VBProject;$compile=$app.VBE.CommandBars.FindControl(1,578)
    if(-not $compile){throw 'VBE compile command unavailable.'}
    if($compile.Enabled){$compile.Execute()}
    if($compile.Enabled){
        $pane=$app.VBE.ActiveCodePane;$sl=0;$sc=0;$el=0;$ec=0;$pane.GetSelection([ref]$sl,[ref]$sc,[ref]$el,[ref]$ec)
        Write-Proof 'CompileFailure.json' ([ordered]@{Component=$pane.CodeModule.Parent.Name;Line=$sl;Text=$pane.CodeModule.Lines([math]::Max(1,$sl-2),5)})
        throw 'Test-only sync did not compile.'
    }
    $afterBody=Body $component
    [IO.File]::WriteAllText((Join-Path $directory 'TargetModule.After.bas'),$afterBody,(New-Object Text.UTF8Encoding($true)))
    $afterMatches=(Canonical $afterBody) -ceq (Canonical $newBody)
    if(-not $afterMatches){throw 'Target module does not match reviewed B0A source after compilation.'}
    $afterInventory=Inventory $book;$inventorySame=($beforeInventory -join "`n") -ceq ($afterInventory -join "`n")
    if(-not $inventorySame){throw 'Component inventory changed.'}
    $book.Save();$saved=$true;$book.Close($false);$book=$null
}catch{$failure=$_.Exception.ToString();Write-Proof 'Failure.json' ([ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');Reason=$failure})}
finally{
    if($owned -and (Exact-Process $excelID $ticks)){
        if($book){try{$book.Close($false)}catch{}}
        if($app){try{$app.Quit()}catch{}}
    }
    $component=$null;$compile=$null;$pane=$null;$book=$null
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $forced=$false
    if($owned -and (Exact-Process $excelID $ticks)){
        $process=Get-Process -Id $excelID;[void]$process.WaitForExit(3000)
        if(Exact-Process $excelID $ticks){Stop-Process -Id $excelID -Force;$forced=$true}
    }
    $closed=-not (Exact-Process $excelID $ticks)
    Write-Proof 'ProcessClosure.json' ([ordered]@{PID=$excelID;StartTicksUTC=[string]$ticks;Owned=$owned;ResidualForceKilled=$forced;ExactIdentityAbsent=$closed;MacrosRun=$false})
}
$unchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash
$newUnchanged=(Get-FileHash -LiteralPath $newSource).Hash -ceq $newHash;$oldUnchanged=(Get-FileHash -LiteralPath $oldSource).Hash -ceq $oldHash
$outputHash=(Get-FileHash -LiteralPath $output).Hash
$preservationPassed=$false;$nonVbaDifferences=@()
if($saved -and -not $failure){
    $python='C:/Users/avidzba/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe'
    & $python (Join-Path $PSScriptRoot 'TestVerifyOpenXmlPreservation.py') --before $source --after $output --report (Join-Path $directory 'DataPreservation.json') | Set-Content -LiteralPath (Join-Path $directory 'DataPreservation.stdout.txt') -Encoding UTF8
    $preservationExit=$LASTEXITCODE
    $preservation=Get-Content -LiteralPath (Join-Path $directory 'DataPreservation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $preservationPassed=$preservationExit -eq 0 -and $preservation.Status -ceq 'PASS'
    $nonVbaDifferences=@($preservation.ZIPParts.Differences | Where-Object Part -cne 'xl/vbaProject.bin')
    Write-Proof 'ZIPNonVBAReview.json' ([ordered]@{SourceSHA256=$sourceHash;OutputSHA256=$outputHash;PartsBefore=$preservation.ZIPParts.BeforeCount;PartsAfter=$preservation.ZIPParts.AfterCount;AllDifferences=$preservation.ZIPParts.Differences;NonVBADifferences=$nonVbaDifferences;AllNonVBAPartsByteUnchanged=($nonVbaDifferences.Count -eq 0);SemanticDataPreservation=$preservationPassed;ExcelSaveMayNormalizeMetadata=$true;NoClaimOfByteIdentityForChangedParts=$true})
}
$passed=$saved -and $beforeMatches -and $afterMatches -and $inventorySame -and $unchanged -and $oldUnchanged -and $newUnchanged -and $closed -and $preservationPassed -and -not $failure
Write-Proof 'Preparation.json' ([ordered]@{Status=$(if($passed){'PREPARED_TEST_ONLY_SYNC_COMPILED_REQUIRES_FINAL_ACCEPTANCE'}else{'REVIEW_REQUIRED'});Passed=$passed;SourceSHA256=$sourceHash;CandidateSHA256=$outputHash;SourceUnchanged=$unchanged;OldTestSourceUnchanged=$oldUnchanged;NewTestSourceUnchanged=$newUnchanged;TargetBeforeMatchesC809=$beforeMatches;TargetAfterMatchesB0A=$afterMatches;ComponentInventoryUnchanged=$inventorySame;ComponentsBefore=$beforeInventory;ComponentsAfter=$afterInventory;OwnProcessesClosed=$closed;OnlyComponentEdited='modTestAutoCADContours';MacrosRun=$false;SavePerformed=$saved;CellWritesPerformed=$false;DataPreservationPassed=$preservationPassed;NonVBADifferences=$nonVbaDifferences;ProductionBodyEqualityDeferredToFinalNativeExport=$true;FullRegressionPass=$false;Failure=$failure})
if(-not $passed){throw 'Test-only sync requires review; inspect primary evidence.'}
Get-Content -LiteralPath (Join-Path $directory 'Preparation.json') -Raw
