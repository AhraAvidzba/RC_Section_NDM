# Обновляет VBA существующей книги через единый проверенный importer.
# Не пишет устаревшую семиколоночную Config-схему и не назначает defaults
# поверх пользовательского ввода. Схема и справка обновляются отдельно
# в явно запущенном сборочном сценарии.
param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Refresh-VbaModules.ps1") -WorkbookPath $WorkbookPath
