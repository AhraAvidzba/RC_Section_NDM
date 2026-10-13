$ErrorActionPreference='Stop'
$root=(Get-Location).Path
$directory=Join-Path $root 'docs/regression/ExcelGuardGuide_2026-10-13/RecipeArguments'
$before=Get-Content -LiteralPath (Join-Path $directory 'Before.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$preserve=Get-Content -LiteralPath (Join-Path $directory 'Preservation.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$equality=Get-Content -LiteralPath (Join-Path $directory 'SourceEquality.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$help=Get-Content -LiteralPath (Join-Path $directory 'StandaloneHelp.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$update=Get-Content -LiteralPath (Join-Path $directory 'CodeUpdate.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if(-not $preserve.Passed -or -not $preserve.ExecutableCodeUnchanged -or $preserve.GuideExampleInlineComments -ne 13 -or $preserve.RecipeClasses -ne 5 -or $preserve.CommentedCommands -ne 229 -or $preserve.ArgumentsExplained -ne 687 -or -not $preserve.ExampleCompiled -or -not $equality.Passed -or $help.status -ne 'STANDALONE_HELP_PASS' -or -not $update.Compiled){throw 'Comment-only acceptance incomplete.'}
$candidate=Join-Path $directory 'Candidate.xlsm'
$nativePath=Join-Path $directory 'NativeVBA.txt'
if((Get-FileHash -LiteralPath $candidate).Hash.ToLowerInvariant() -cne $preserve.CandidateSHA256){throw 'Candidate changed after acceptance.'}
if((Get-FileHash -LiteralPath $nativePath).Hash.ToLowerInvariant() -cne $preserve.NativeSHA256){throw 'Native export changed after acceptance.'}
$published=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$allCode=Join-Path $root 'workbook/output/VBA_All_Code.txt'
$execution=Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if((Get-FileHash -LiteralPath $execution).Hash -cne $before.ExecutionReportSHA256){throw 'User execution report changed since snapshot; leave it alone.'}
$native=[IO.File]::ReadAllText($nativePath,[Text.Encoding]::UTF8)
if($native -notmatch '(?m)^Workbook: Candidate.xlsm\r?$'){throw 'Unexpected export identity.'}
$native=$native.Replace('Workbook: Candidate.xlsm','Workbook: RC_Section_NDM.xlsm')
$utf8=New-Object Text.UTF8Encoding($true)
[byte[]]$bookBytes=[IO.File]::ReadAllBytes($candidate)
[byte[]]$codeBytes=$utf8.GetPreamble()+$utf8.GetBytes($native)
$bookStream=$null;$codeStream=$null
$digest=[Security.Cryptography.SHA256]::Create()
try {
    $bookStream=[IO.File]::Open($published,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $codeStream=[IO.File]::Open($allCode,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $bookHash=[BitConverter]::ToString($digest.ComputeHash($bookStream)).Replace('-','')
    $codeHash=[BitConverter]::ToString($digest.ComputeHash($codeStream)).Replace('-','')
    if($bookHash -cne $before.BeforeSHA256 -or $codeHash -cne $before.AllCodeSHA256){throw 'Production inputs changed; preserve current data.'}
    $codeStream.Position=0
    $backup=[IO.File]::Create((Join-Path $directory 'AllCodeBeforePublication.txt'))
    try {$codeStream.CopyTo($backup)} finally {$backup.Dispose()}
    $bookStream.Position=0;$bookStream.Write($bookBytes,0,$bookBytes.Length);$bookStream.SetLength($bookBytes.Length);$bookStream.Flush($true)
    $codeStream.Position=0;$codeStream.Write($codeBytes,0,$codeBytes.Length);$codeStream.SetLength($codeBytes.Length);$codeStream.Flush($true)
} finally {
    if($codeStream){$codeStream.Dispose()}
    if($bookStream){$bookStream.Dispose()}
    $digest.Dispose()
}
if((Get-FileHash -LiteralPath $published).Hash.ToLowerInvariant() -cne $preserve.CandidateSHA256){throw 'Published book differs from accepted candidate.'}
if((Get-FileHash -LiteralPath $execution).Hash -cne $before.ExecutionReportSHA256){throw 'Execution report changed during publication.'}
[ordered]@{PublishedAt=(Get-Date).ToString('o');WorkbookSHA256=(Get-FileHash -LiteralPath $published).Hash;AllCodeSHA256=(Get-FileHash -LiteralPath $allCode).Hash;UserExecutionReportUnchanged=$true;ExecutableCodeUnchanged=$true;GuideExampleInlineComments=13;RecipeClasses=5;CommentedCommands=229;ArgumentsExplained=687;ReadbackPending=$true} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Publication.json') -Encoding UTF8
Write-Output 'COMMENT_ONLY_PAIR_PUBLISHED: native readback remains.'
