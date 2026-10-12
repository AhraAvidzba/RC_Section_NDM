# Updates source modules in an isolated workbook; never writes the production file.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [string[]]$ModuleNames = @(),
    [string[]]$RetiredModules = @(),
    [string[]]$EntryPointNames = @(),
    [string[]]$ExampleFiles = @(),
    [string[]]$RecipeRegistration = @(),
    [switch]$TraceGeometry
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'GeometryExampleSupport.ps1')
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only isolated evidence workbooks may be updated.'}
$printAreas=@(Get-WorkbookPrintAreas $path)
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
function Get-OpenWorkbookHash([string]$Path) {
    # Excel держит исходную книгу открытой. Разрешение совместного чтения
    # не меняет файл; сравнение до/после всё равно обнаруживает его сохранение.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $digest=[Security.Cryptography.SHA256]::Create()
    try {[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}
    finally {$digest.Dispose();$stream.Dispose()}
}
$productionHash=Get-OpenWorkbookHash $production
if ($EntryPointNames.Count -and ($ModuleNames.Count -ne 1 -or $RetiredModules.Count -or $TraceGeometry)) {throw 'Entry-point instrumentation requires exactly one module and no retirement/trace.'}
foreach ($name in $EntryPointNames) {if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {throw ('Invalid entry-point name: '+$name)}}
$sourceFiles=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
$sourceFiles+=@(Get-GeometryExampleFiles $root $ExampleFiles)
if ($RecipeRegistration.Count -and ('CSectionTypeRegistry' -notin $ModuleNames -or $RecipeRegistration[1] -notin $ModuleNames)) {throw 'Example registration must import both registry and recipe.'}
foreach ($name in $RetiredModules) {
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$' -or $name -in $ModuleNames) {throw ('Invalid retirement: '+$name)}
    if (@($sourceFiles | Where-Object BaseName -eq $name).Count) {throw ('Retired source still exists: '+$name)}
}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PrivateCodeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    [uint32]$id=0;[void][PrivateCodeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    $ownedPid=[int]$id;$ownedTicks=(Get-Process -Id $ownedPid).StartTime.ToUniversalTime().Ticks
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Private workbook is read-only.'}
    $updated=@()
    $removed=@()
    $componentNames=@($book.VBProject.VBComponents | ForEach-Object {$_.Name})
    foreach ($name in $RetiredModules) {
        if ($name -notin $componentNames) {throw ('Retirement component not found: '+$name)}
        $component=$book.VBProject.VBComponents.Item($name)
        if ($component.Type -notin 1,2) {throw ('Cannot retire a document/designer component: '+$name)}
        $book.VBProject.VBComponents.Remove($component)
        $removed += $name
    }
    $componentNames=@($book.VBProject.VBComponents | ForEach-Object {$_.Name})
    foreach ($file in $sourceFiles) {
        if ($ModuleNames.Count -and $file.BaseName -notin $ModuleNames) {continue}
        $body=[regex]::Match([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if (-not $body -and $file.BaseName -eq 'modSectionRecipeApiGuide') {
            $body=([IO.File]::ReadAllLines($file.FullName,[Text.Encoding]::UTF8) | Where-Object {$_ -notmatch '^Attribute VB_Name = '}) -join "`r`n"
            if (@($body -split "`r?`n" | Where-Object {$_.Trim() -and $_ -notmatch "^\s*'"}).Count) {throw 'API guide must contain comments only.'}
        }
        if (-not $body) {throw ('Source body missing: '+$file.FullName)}
        if ($file.BaseName -eq 'CSectionTypeRegistry') {$body=Add-GeometryExampleRegistration $body $RecipeRegistration}
        if ($TraceGeometry -and $file.BaseName -eq 'CSectionGeometryQuery') {
            $body=$body.Replace('Option Explicit',"Option Explicit`r`nPrivate mCutTrace As String")
            $body=$body.Replace('    Set ClipHalfPlane = MakeRegion',"    mCutTrace = CutTrace(raw, rawCount, nx, ny, bound, lineCuts, lineCount)`r`n    Set ClipHalfPlane = MakeRegion")
            $needle='If candidate = 0 Then GeometryError'
            if ([regex]::Matches($body,[regex]::Escape($needle)).Count -ne 1) {throw 'Geometry trace target mismatch.'}
            $body=$body.Replace($needle,'If candidate = 0 Then GeometryError FailedRegionTrace(edges, count, used) & ')
            $body += @'

' ДЛЯ ТЕСТОВ: диагностика только в отдельной книге, не в исходном ядре.
Private Function FailedRegionTrace(ByRef edges() As TRegionEdge, ByVal count As Long, ByRef used() As Boolean) As String
    Dim i As Long, result As String
    result = mCutTrace & "REGION_TRACE; count=" & CStr(count) & vbCrLf
    For i = 1 To count
        result = result & "EDGE|" & CStr(i) & "|used=" & CStr(used(i)) & "|" & _
            CStr(edges(i).X1) & "|" & CStr(edges(i).Y1) & "|" & _
            CStr(edges(i).X2) & "|" & CStr(edges(i).Y2) & "|sweep=" & _
            CStr(edges(i).Sweep) & vbCrLf
    Next i
    FailedRegionTrace = result
End Function

Private Function CutTrace(ByRef raw() As TRegionEdge, ByVal rawCount As Long, ByVal nx As Double, ByVal ny As Double, ByVal bound As Double, ByRef cuts() As Double, ByVal count As Long) As String
    Dim i As Long, s As Double, x As Double, y As Double, result As String
    result = "CUT_TRACE|nx=" & CStr(nx) & "|ny=" & CStr(ny) & "|bound=" & CStr(bound) & "|cuts=" & CStr(count) & vbCrLf
    For i = 1 To count
        result = result & "CUT|" & CStr(i) & "|s=" & CStr(cuts(i)) & vbCrLf
        If i < count Then
            s = (cuts(i) + cuts(i + 1)) * 0.5
            x = nx * bound + ny * s: y = ny * bound - nx * s
            result = result & "MID|" & CStr(x) & "|" & CStr(y) & "|inside=" & CStr(ContainsRaw(raw, rawCount, x + nx * QUERY_TOLERANCE * 8#, y + ny * QUERY_TOLERANCE * 8#, False)) & vbCrLf
        End If
    Next i
    CutTrace = result
End Function
'@
        }
        if ($file.BaseName -in $componentNames) {
            $component=$book.VBProject.VBComponents.Item($file.BaseName)
        } else {
            $component=$book.VBProject.VBComponents.Add($(if ($file.Extension -eq '.cls') {2} else {1}))
            $component.Name=$file.BaseName
            $componentNames+=,$file.BaseName
        }
        $module=$component.CodeModule
        if ($EntryPointNames.Count) {
            $original=[string]$module.Lines(1,$module.CountOfLines)
            $entries=@()
            foreach ($name in $EntryPointNames) {
                $declaration='(?m)^Public (?:Function|Sub) '+[regex]::Escape($name)+'\('
                if ([regex]::IsMatch($original,$declaration)) {throw ('Entry point already exists in baseline: '+$name)}
                $matches=[regex]::Matches($body,'(?ms)^Public (Function|Sub) '+[regex]::Escape($name)+'\(.*?^End \1\b')
                if ($matches.Count -ne 1) {throw ('Source entry point is not unique: '+$name)}
                $entries+=,"' FOR TESTS: baseline entry point only; existing methods retained.`r`n"+$matches[0].Value
            }
            $body=$entries -join "`r`n`r`n"
        } elseif ($module.CountOfLines) {$module.DeleteLines(1,$module.CountOfLines)}
        $module.AddFromString(($body -split '\r?\n') -join "`r`n")
        $updated += $file.BaseName
    }
    if ($ModuleNames.Count -and @($ModuleNames | Where-Object {$_ -notin $updated}).Count) {throw 'Requested source module was not found.'}
    $excel.VBE.ActiveVBProject=$book.VBProject
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    if ($compile.Enabled) {
        $pane=$excel.VBE.ActiveCodePane
        [int]$firstLine=0;[int]$firstColumn=0;[int]$lastLine=0;[int]$lastColumn=0
        $pane.GetSelection([ref]$firstLine,[ref]$firstColumn,[ref]$lastLine,[ref]$lastColumn)
        Write-Output ('COMPILE_LOCATION: module='+$pane.CodeModule.Parent.Name+'; line='+$firstLine+'; column='+$firstColumn)
        Write-Output $pane.CodeModule.Lines([Math]::Max(1,$firstLine-2),5)
        throw 'Private VBA compilation failed.'
    }
    $book.Save();$book.Close($false);$book=$null
    Restore-WorkbookPrintAreas $path $printAreas
    if ((Get-OpenWorkbookHash $production) -cne $productionHash) {throw 'Production changed during preparation.'}
    [ordered]@{Workbook=$path;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash;ProductionSHA256=$productionHash;Compiled=$true;Modules=$updated;RetiredModules=$removed;EntryPointNames=$EntryPointNames;TraceGeometry=[bool]$TraceGeometry;ProductionChanged=$false} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $path) 'CodeUpdate.json') -Encoding UTF8
    Write-Output ('PRIVATE_CODE_READY: modules='+$updated.Count+'; production unchanged=True')
} finally {
    if ($book) {$book.Close($false)}
    if ($excel) {$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
}
