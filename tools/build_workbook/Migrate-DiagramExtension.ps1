# Точечно обновляет глобальную настройку Extension, не пересобирая ввод книги.
param([string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm")
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "SettingsCatalog.ps1")
$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$path = Join-Path $root $WorkbookPath
$printAreas = @(Get-WorkbookPrintAreas $path)
$excel = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($path)
    if ($workbook.ReadOnly) { throw "Книга открылась только для чтения. Закройте ее перед миграцией." }
    $result = Invoke-DiagramExtensionMigration $workbook
    $workbook.Save()
    Write-Output "DiagramExtension migrated: value=$($result.Value); row=$($result.Row); workbook=$path"
} finally {
    if ($workbook) {
        $workbook.Close($false)
        [Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
Restore-WorkbookPrintAreas $path $printAreas
