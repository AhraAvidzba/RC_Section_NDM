# Обновление только списка типов сечений по реестру; исходники импортируются отдельно.
param([Parameter(Mandatory=$true)][string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$path = (Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$beforeHash = (Get-FileHash -LiteralPath $path).Hash
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class RecipeSetupWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null; $book=$null; $ownedId=0; $ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    [uint32]$id=0; [void][RecipeSetupWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedId=[int]$id; $ownedTicks=(Get-Process -Id $ownedId).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedId+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Target workbook is read-only.'}
    $namesBefore=@($book.Names | ForEach-Object {$_.Name})
    $inputNames=@('rngLoadCombinations','rngCalculationProfiles','rngUnitSettings','rngSignConventionSettings')
    $inputsBefore=@{}
    foreach ($name in $inputNames) {$inputsBefore[$name]=$book.Names.Item($name).RefersToRange.Formula}
    $macro="'"+$book.Name+"'!modSectionRecipeSetup.UpdateSectionRecipeTypes"
    $excel.Run($macro)
    foreach ($name in $inputNames) {
        $after=$book.Names.Item($name).RefersToRange.Formula
        if ((ConvertTo-Json -InputObject $inputsBefore[$name] -Compress -Depth 4) -cne (ConvertTo-Json -InputObject $after -Compress -Depth 4)) {throw ('Unrelated input changed: '+$name)}
    }
    # Повторное обновление сохраняет геометрические таблицы и выбранные параметры.
    $first=@{}
    foreach ($name in @('rngRoundedRectangleGeometry','rngHollowRectangleGeometry','rngRectSetGeometry','rngCircleGeometry')) {$first[$name]=$book.Names.Item($name).RefersToRange.Formula}
    $excel.Run($macro)
    foreach ($name in $first.Keys) {
        if ((ConvertTo-Json -InputObject $first[$name] -Compress -Depth 4) -cne (ConvertTo-Json -InputObject $book.Names.Item($name).RefersToRange.Formula -Compress -Depth 4)) {throw ('Update is not idempotent: '+$name)}
    }
    $book.Save()
    $namesAfter=@($book.Names | ForEach-Object {$_.Name})
    $book.Close($false); $book=$null
    [ordered]@{Status='PASS';BeforeSHA256=$beforeHash;AfterSHA256=(Get-FileHash -LiteralPath $path).Hash;NamesBefore=$namesBefore;NamesAfter=$namesAfter;UnrelatedInputsUnchanged=$true;Idempotent=$true} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $path) 'RecipeTypesUpdate.json') -Encoding UTF8
    Write-Output 'RECIPE_TYPES_READY: preserved inputs; idempotent=True'
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit(); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($ownedId) {
        $owned=Get-Process -Id $ownedId -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedId -Force}
    }
}
