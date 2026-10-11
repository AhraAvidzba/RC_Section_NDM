# Saves an immutable baseline and reads native Excel metadata without saving it.
param([string]$Directory='docs/regression/GeometryDeclarative_2026-10-11/Baseline')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directoryPath=[IO.Path]::GetFullPath((Join-Path $root $Directory))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/GeometryDeclarative_2026-10-11'))+[IO.Path]::DirectorySeparatorChar
if (-not $directoryPath.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Unexpected baseline path.'}
if (Test-Path -LiteralPath $directoryPath) {throw 'Baseline already exists; never overwrite it.'}
New-Item -ItemType Directory -Path $directoryPath | Out-Null
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
$productionHash=(Get-FileHash -LiteralPath $production).Hash
$bookPath=Join-Path $directoryPath 'RC_Section_NDM.xlsm'
Copy-Item -LiteralPath $production -Destination $bookPath
Copy-Item -LiteralPath (Join-Path $root 'workbook/output/VBA_All_Code.txt') -Destination $directoryPath
Compress-Archive -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests'),(Join-Path $root 'tools'),(Join-Path $root 'AGENTS.md'),(Join-Path $root 'README.md') -DestinationPath (Join-Path $directoryPath 'SourceBaseline.zip') -CompressionLevel Optimal
git -C $root status --porcelain=v1 | Set-Content -LiteralPath (Join-Path $directoryPath 'GitStatus.txt') -Encoding UTF8
git -C $root diff --binary HEAD | Set-Content -LiteralPath (Join-Path $directoryPath 'DirtyInput.patch') -Encoding UTF8
$head=git -C $root rev-parse HEAD
$files=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | ForEach-Object {
    [ordered]@{Path=$_.FullName.Substring($root.Length+1).Replace('\','/');SHA256=(Get-FileHash -LiteralPath $_.FullName).Hash;Length=$_.Length}
})
$spec=Join-Path $root 'docs/geometry_refactor/NDM_Geometry_Declarative_Refactor_TZ_CodexGoal.md'
$manifest=[ordered]@{CapturedAt=(Get-Date -Format o);Head=$head;SpecVersion='1.4';SpecSHA256=(Get-FileHash -LiteralPath $spec).Hash;WorkbookSHA256=$productionHash;DumpSHA256=(Get-FileHash -LiteralPath (Join-Path $directoryPath 'VBA_All_Code.txt')).Hash;ArchiveSHA256=(Get-FileHash -LiteralPath (Join-Path $directoryPath 'SourceBaseline.zip')).Hash;DirtyDiffSHA256=(Get-FileHash -LiteralPath (Join-Path $directoryPath 'DirtyInput.patch')).Hash;SourceFiles=$files;InventoryStatus='NOT_RUN'}
$manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'Manifest.json') -Encoding UTF8
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class GeometryBaselineWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
function TextHash([string]$value) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try {return [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value))).Replace('-','')}
    finally {$sha.Dispose()}
}
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][GeometryBaselineWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$process=Get-Process -Id $ownedPid;$ownedTicks=$process.StartTime.ToUniversalTime().Ticks
    [ordered]@{Pid=$ownedPid;UtcStartTicks=$ownedTicks;Workbook=$bookPath} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $directoryPath 'OwnedExcel.json') -Encoding UTF8
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($bookPath,0,$true)
    $names=@();foreach ($name in $book.Names) {$names += [ordered]@{Name=[string]$name.Name;RefersTo=[string]$name.RefersTo;Visible=[bool]$name.Visible}}
    $components=@();foreach ($component in $book.VBProject.VBComponents) {
        $module=$component.CodeModule;$body='';if ($module.CountOfLines) {$body=[string]$module.Lines(1,$module.CountOfLines)}
        $components += [ordered]@{Name=[string]$component.Name;Type=[int]$component.Type;Lines=[int]$module.CountOfLines;CodeSHA256=(TextHash $body)}
    }
    $references=@();foreach ($reference in $book.VBProject.References) {$references += [ordered]@{Name=[string]$reference.Name;GUID=[string]$reference.GUID;Major=[int]$reference.Major;Minor=[int]$reference.Minor;FullPath=[string]$reference.FullPath;Broken=[bool]$reference.IsBroken}}
    $sheets=@();foreach ($sheet in $book.Worksheets) {
        $objects=@();foreach ($shape in $sheet.Shapes) {
            $action='';try {$action=[string]$shape.OnAction} catch {}
            $objects += [ordered]@{Name=[string]$shape.Name;Type=[int]$shape.Type;Left=[double]$shape.Left;Top=[double]$shape.Top;Width=[double]$shape.Width;Height=[double]$shape.Height;OnAction=$action}
        }
        $sheets += [ordered]@{Name=[string]$sheet.Name;CodeName=[string]$sheet.CodeName;UsedRange=[string]$sheet.UsedRange.Address();Visible=[int]$sheet.Visible;Shapes=$objects}
        if ([string]$sheet.Name -in @('Config','Results')) {
            [ordered]@{Name=[string]$sheet.Name;Address=[string]$sheet.UsedRange.Address();Formula=$sheet.UsedRange.Formula;Value2=$sheet.UsedRange.Value2} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directoryPath ($sheet.Name+'.json')) -Encoding UTF8
        }
    }
    $validation=@();$config=$book.Worksheets.Item('Config');$cells=$null
    try {$cells=$config.UsedRange.SpecialCells(-4174)} catch {}
    if ($cells) {foreach ($cell in $cells.Cells) {$validation += [ordered]@{Address=[string]$cell.Address();Type=[int]$cell.Validation.Type;Formula1=[string]$cell.Validation.Formula1;Formula2=[string]$cell.Validation.Formula2}}}
    $inventory=[ordered]@{CapturedAt=(Get-Date -Format o);ExcelVersion=[string]$excel.Version;ExcelBuild=[string]$excel.Build;ExcelExecutable=$process.Path;Calculation=[int]$excel.Calculation;Names=$names;Components=$components;References=$references;Sheets=$sheets;ConfigValidation=$validation}
    $inventory | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $directoryPath 'NativeInventory.json') -Encoding UTF8
    $manifest.InventoryStatus='PASS';$manifest.InventorySHA256=(Get-FileHash -LiteralPath (Join-Path $directoryPath 'NativeInventory.json')).Hash
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'Manifest.json') -Encoding UTF8
    Write-Output ('BASELINE_CAPTURED: components='+$components.Count+'; names='+$names.Count+'; validation='+$validation.Count+'; source='+$productionHash)
} catch {
    $manifest.InventoryStatus='FAIL';$manifest.InventoryError=$_.Exception.Message
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $directoryPath 'Manifest.json') -Encoding UTF8
    throw
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
    if ((Get-FileHash -LiteralPath $production).Hash -cne $productionHash -or (Get-FileHash -LiteralPath $bookPath).Hash -cne $productionHash) {throw 'Read-only baseline capture changed a workbook.'}
}
