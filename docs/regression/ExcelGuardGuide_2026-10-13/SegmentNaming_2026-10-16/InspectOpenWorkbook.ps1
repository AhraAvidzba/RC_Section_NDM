# Только чтение состояния основной книги перед безопасной публикацией.
param([switch]$SaveCopy)
$ErrorActionPreference='Stop'
$path=Join-Path (Get-Location).Path 'workbook/output/RC_Section_NDM.xlsm'
$book=[Runtime.InteropServices.Marshal]::BindToMoniker($path)
try {
    if([string]$book.FullName -cne $path){throw 'Unexpected workbook'}
    if($SaveCopy){
        $copy=Join-Path $PSScriptRoot 'LiveBeforePublication.xlsm'
        if(Test-Path -LiteralPath $copy){throw 'Live snapshot already exists'}
        if($book.ReadOnly -or -not $book.Application.Ready){throw 'Workbook unavailable'}
        $book.SaveCopyAs($copy)
    }
    [ordered]@{Path=$book.FullName;Saved=[bool]$book.Saved;ReadOnly=[bool]$book.ReadOnly;Ready=[bool]$book.Application.Ready;Workbooks=[int]$book.Application.Workbooks.Count} | ConvertTo-Json
} finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}
