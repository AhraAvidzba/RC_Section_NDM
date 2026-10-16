# Пользователь разрешил закрывать основную книгу. Перед закрытием сохраняем всё её текущее состояние отдельно.
$ErrorActionPreference='Stop'
$path=Join-Path (Get-Location).Path 'workbook/output/RC_Section_NDM.xlsm'
$copy=Join-Path $PSScriptRoot 'LiveFinalBeforeClose.xlsm'
if(Test-Path -LiteralPath $copy){throw 'Final live snapshot already exists'}
$book=[Runtime.InteropServices.Marshal]::BindToMoniker($path)
try {
    if([string]$book.FullName -cne $path -or $book.ReadOnly -or -not $book.Application.Ready){throw 'Workbook unavailable'}
    $state=[ordered]@{FullName=$book.FullName;WasSaved=[bool]$book.Saved;OtherWorkbooks=[int]$book.Application.Workbooks.Count-1;Snapshot=$copy}
    $book.SaveCopyAs($copy)
    if(-not (Test-Path -LiteralPath $copy) -or (Get-Item -LiteralPath $copy).Length -lt 1000000){throw 'Snapshot was not saved'}
    # Закрываем только эту книгу. Её актуальные данные сохранены в snapshot и войдут в итоговый файл.
    $book.Close($false)
    $state.Closed=$true
    $state | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'MainBookClose.json') -Encoding UTF8
    $state | ConvertTo-Json
} finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
