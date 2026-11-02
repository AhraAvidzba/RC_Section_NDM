# ==================== ДЛЯ ТЕСТОВ ====================
# Возвращает JSON фактически проверенного полного PASS; никаких изменений или процессов Excel.
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'TestAssertCurrentRun02Pass60.ps1')
$proof=Test-AssertCurrentRun02Pass60 -RequireTaskRemoved -RequireAdditionalIO
$proof|Select-Object Root,Launch,Evidence,Source,SourceSHA256,SourceCommit,InputsChecked,CoreSuites,FocusedSingles,FocusedRecords,OwnCount,AllOwnIdentitiesTerminal,ScopeReviewSHA256,OwnProofSHA256,OfficialPASS|ConvertTo-Json -Depth 4
