# Публикует проверенные артефакты только при неизменной пользовательской книге.
# Не пересчитывает Results и не заменяет последний пользовательский отчет запуска.
param([string]$ReportDirectory = '')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
$directory = Join-Path $PSScriptRoot 'FinalPublication'
if ($ReportDirectory) { $directory = Join-Path $root $ReportDirectory }
$manifest = Get-Content -LiteralPath (Join-Path $directory 'Manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$target = Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$source = Join-Path $directory 'RC_Section_NDM.xlsm'
$sourceCode = Join-Path $directory 'VBA_All_Code.txt'
$targetCode = Join-Path $root 'workbook/output/VBA_All_Code.txt'
$report = Join-Path $root 'workbook/output/RC_Section_NDM_execution_report.txt'
if ((Get-FileHash -LiteralPath $target).Hash -ne $manifest.SourceSHA256) { throw 'User workbook changed; publication must rebase.' }
if ((Get-FileHash -LiteralPath $source).Hash -ne $manifest.PreparedSHA256) { throw 'Prepared workbook changed after verification.' }
$reportHash = (Get-FileHash -LiteralPath $report).Hash
$codeHash = (Get-FileHash -LiteralPath $sourceCode).Hash
Copy-Item -LiteralPath $source -Destination $target -Force
Copy-Item -LiteralPath $sourceCode -Destination $targetCode -Force
if ((Get-FileHash -LiteralPath $target).Hash -ne $manifest.PreparedSHA256) { throw 'Published workbook differs.' }
if ((Get-FileHash -LiteralPath $targetCode).Hash -ne $codeHash) { throw 'Published code differs.' }
if ((Get-FileHash -LiteralPath $report).Hash -ne $reportHash) { throw 'User execution report changed.' }
[pscustomobject]@{
    PublishedWorkbookSHA256=$manifest.PreparedSHA256
    PublishedCodeSHA256=$codeHash
    UserExecutionReportSHA256=$reportHash
    UserExecutionReportPreserved=$true
    CalculationExecuted=$false
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directory 'Published.json') -Encoding UTF8
Write-Output 'PUBLICATION_OK'
