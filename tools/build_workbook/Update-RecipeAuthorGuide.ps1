# Обновляет текст существующей справки автора, сохраняя её оформление и прочие данные.
param([Parameter(Mandatory=$true)][string]$WorkbookPath)
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
    if($old[-1] -notmatch '^NDM.RECIPE.GUIDE/2\|([a-f0-9]{64})\|[a-f0-9]{64}$'){throw 'Missing guide footer.'}
    if((Get-RecipeGuideTextHash $old[0..($old.Count-2)]) -cne $Matches[1]){throw 'Guide was edited; reconcile before replacing.'}
    $changed=@()
    for($i=0;$i -lt $layout.rows.Count;$i++) {
        $text=[string]$layout.rows[$i].text
        if($old[$i] -ceq $text){continue}
        $cell=$range.Cells.Item($i+1,1)
        $hidden=[bool]$cell.EntireRow.Hidden
        $cell.Value2=$text
        if($layout.rows[$i].style -eq 'paragraph') {
            $chars=[Math]::Max(20,[Math]::Floor(($cell.MergeArea.Width-16)/6.7))
            $height=18*[Math]::Max(1,[Math]::Ceiling($text.Length/$chars))+3
            $cell.EntireRow.RowHeight=[Math]::Max($cell.RowHeight,$height)
            $cell.EntireRow.Hidden=$hidden
        }
        $changed+=,[string]$cell.Address($false,$false)
    }
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $path $areas
    [ordered]@{ChangedCells=$changed;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path (Split-Path $path -Parent) 'GuideUpdate.json') -Encoding UTF8
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
