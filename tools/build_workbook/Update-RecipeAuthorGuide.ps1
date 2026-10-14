# Обновляет текст существующей справки автора, сохраняя её оформление и прочие данные.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,[string]$PreviousLayoutPath='')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
. (Join-Path $PSScriptRoot 'SectionRecipeGuide.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only a private verification workbook may be updated.'}
$areas=@(Get-WorkbookPrintAreas $path)
$layout=Get-Content -LiteralPath (Join-Path $root 'docs/geometry_refactor/SECTION_RECIPE_GUIDE_LAYOUT.json') -Raw -Encoding UTF8 | ConvertFrom-Json
foreach($property in $layout.sourceHashes.PSObject.Properties) {
    if((Get-FileHash -LiteralPath (Join-Path $root $property.Name)).Hash.ToLowerInvariant() -cne $property.Value) {throw 'Guide sources changed after generation.'}
}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class AuthorGuideWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.EnableEvents=$false;$excel.DisplayAlerts=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][AuthorGuideWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if($book.ReadOnly){throw 'Candidate is read-only.'}
    $range=$book.Names.Item('rngSectionRecipeGuide').RefersToRange
    if($range.Columns.Count -ne 6 -or $range.Rows.Count -ne $layout.rows.Count){throw 'Unexpected guide layout.'}
    $old=@(Get-RecipeGuideNativeLines $range.Worksheet $range.Row ($range.Row+$range.Rows.Count-1))
    Assert-RecipeGuidePreviousContent $old $layout $PreviousLayoutPath
    $changed=@()
    for($i=0;$i -lt $layout.rows.Count;$i++) {
        $text=[string]$layout.rows[$i].text
        if($old[$i] -ceq $text){continue}
        $cell=$range.Cells.Item($i+1,1)
        $cell.Value2=$text
        $changed+=,[string]$cell.Address($false,$false)
    }
    $heights=@(Set-RecipeGuideRowHeights $range.Worksheet $range.Row $layout)
    $actual=@(Get-RecipeGuideNativeLines $range.Worksheet $range.Row ($range.Row+$range.Rows.Count-1))
    if((Get-RecipeGuideTextHash $actual) -cne $layout.bodySHA256){throw 'Saved guide text differs from external layout.'}
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $path $areas
    [ordered]@{ChangedCells=$changed;RowHeights=$heights;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path (Split-Path $path -Parent) 'GuideUpdate.json') -Encoding UTF8
    Write-Output ('GUIDE_UPDATED: '+($changed -join ','))
} finally {
    if($book){$book.Close($false)}
    if($excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($ownedPid){
        $p=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if($p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ownedTicks){Stop-Process -Id $ownedPid -Force}
    }
}
