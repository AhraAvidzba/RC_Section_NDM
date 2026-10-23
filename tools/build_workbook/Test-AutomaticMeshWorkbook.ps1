# ==========================================================================
# ДЛЯ ТЕСТОВ: Auto-сетка, справка и сохранность книги
# ==========================================================================
# Направленная проверка новой настройки в настоящем Excel: вид, ссылки, чтение
# и четыре параметрических типа. Все изменения остаются в read-only памяти копии.
param([Parameter(Mandatory=$true)][string]$WorkbookPath,
      [string]$EvidenceDirectory='docs/regression/AutomaticMesh_2026-10-23_Run03')
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'WorkbookLayout.ps1')
$script:WorkbookLayout=Read-WorkbookLayout
$path=(Resolve-Path -LiteralPath $WorkbookPath).Path
$directory=(Resolve-Path -LiteralPath $EvidenceDirectory).Path
$hash=(Get-FileHash -LiteralPath $path).Hash
$xml=Get-WorkbookLayoutXmlSheets $path
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class AutoMeshTestWindow{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'
$app=$null;$book=$null;$owned=0;$ticks=0;$beforePids=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Повторно проверяет PID и время старта перед закрытием собственных COM-объектов.
function Test-OwnedExcel {
    if(-not $owned -or $owned -in $beforePids){return $false}
    $p=Get-Process -Id $owned -ErrorAction SilentlyContinue
    return $null -ne $p -and $p.ProcessName -eq 'EXCEL' -and $p.StartTime.ToUniversalTime().Ticks -eq $ticks
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Читает колонку ключей одним вызовом COM, включая объединённые заголовки Справки.
function Test-FindKeyRow($Sheet,[string]$Key) {
    $last=[int]$Sheet.UsedRange.Row+[int]$Sheet.UsedRange.Rows.Count-1
    $values=$Sheet.Range('A1:A'+$last).Value2
    for($r=1;$r -le $last;$r++){if([string]$values[$r,1] -ceq $Key){return $r}}
    throw ('Missing key: '+$Key)
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Сопоставляет существенное оформление с соседним блоком, без сравнения текста.
function Test-StyleSignature($Range) {
    return @($Range.Font.Name,$Range.Font.Size,$Range.Font.Bold,$Range.Font.Italic,$Range.Font.Color,
        $Range.Interior.Color,$Range.HorizontalAlignment,$Range.VerticalAlignment,$Range.WrapText,$Range.IndentLevel) -join '|'
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# VBE меняет регистр имён, пробелы и запись чисел (до 15 значащих цифр).
# Сравнивает токены построчно; строки, комментарии и суффиксы типов сохраняются точно.
function Test-NormalizeCode([string]$Body) {
    $pattern=@'
"(?:[^"]|"")*"|'[^\r\n]*|(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?[#&!@%]?|[A-Za-z_]\w*|[^\s]
'@
    $lines=@(foreach($line in ($Body.TrimEnd() -split '\r?\n')){
        $parts=@(foreach($match in [regex]::Matches($line,$pattern)){
            $token=$match.Value
            if($token.StartsWith('"') -or $token.StartsWith("'")){$token}
            elseif($token -match '^\d|^\.\d'){
                $suffix=if($token -match '[#&!@%]$'){$token.Substring($token.Length-1)}else{''}
                if(-not $suffix -and $token -match '[.Ee]'){$suffix='#'} # Дробный литерал без суффикса уже имеет тип Double; VBE может убрать избыточный #.
                [double]::Parse($token.TrimEnd([char[]]'#&!@%'),[Globalization.CultureInfo]::InvariantCulture).ToString('G15',[Globalization.CultureInfo]::InvariantCulture)+$suffix
            }else{$token.ToLowerInvariant()}
        })
        $parts -join ' '
    })
    return $lines -join "`n"
}

# -------------------- ДЛЯ ТЕСТОВ --------------------
# Создаёт нативный PNG только для локальной визуальной проверки; файл не считается доставленным пользователю.
function Export-TestView($Sheet,[string]$Address,[string]$Output) {
    $Sheet.Activate();$range=$Sheet.Range($Address);$range.CopyPicture(1,-4147)
    $scratch=$app.Workbooks.Add()
    try {
        $chart=$scratch.Worksheets.Item(1).ChartObjects().Add(0,0,[double]$range.Width,[double]$range.Height)
        $chart.Activate();$chart.Chart.Paste()
        if(-not $chart.Chart.Export($Output,'PNG')){throw 'Native render failed.'}
    } finally {if(Test-OwnedExcel){$scratch.Close($false)}}
}

try {
    $app=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][AutoMeshTestWindow]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$id)
    $owned=[int]$id;$ticks=(Get-Process -Id $owned).StartTime.ToUniversalTime().Ticks
    if(-not(Test-OwnedExcel)){throw 'Excel ownership was not established.'}
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AutomationSecurity=1
    $book=$app.Workbooks.Open($path,0,$true)
    if(-not $book.ReadOnly){throw 'Verification workbook must be read-only.'}
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка')
    $help.Outline.ShowLevels(8)
    $r=Test-FindKeyRow $config 'Mesh.StepMode';$x=Test-FindKeyRow $config 'Mesh.StepX';$boundary=Test-FindKeyRow $config 'Mesh.BoundarySubdivisions'
    if($r+1 -ne $x -or [string]$config.Cells.Item($r,2).Value2 -cne 'Auto 0.1' -or $config.Cells.Item($boundary,2).Value2 -ne 2){throw 'Defaults or position are wrong.'}
    $cell=$config.Cells.Item($r,2)
    $options=@('Auto 0.05','Auto 0.1','Auto 0.2','User')
    $localList=$options -join [string]$app.International(5)
    if($cell.Validation.Type -ne 3 -or $cell.Validation.Formula1 -cne $localList -or -not $cell.Validation.InCellDropdown){throw 'Wrong mode dropdown.'}
    $savedMode=$cell.Value2
    foreach($option in $options){$cell.Value2=$option;if(-not $cell.Validation.Value){throw ('Mode rejected: '+$option)}}
    $cell.Value2=$savedMode
    if($cell.Interior.Color -ne 15135487 -or $cell.HorizontalAlignment -ne -4108){throw 'Wrong input style.'}
    if([string]$config.Cells.Item($x,3).Formula -notmatch '^='){throw 'Manual step unit must remain a formula.'}
    $h=Test-FindKeyRow $help 'Mesh.StepMode';$hx=Test-FindKeyRow $help 'Mesh.StepX'
    if((Test-StyleSignature $help.Cells.Item($h,1)) -cne (Test-StyleSignature $help.Cells.Item($hx,1))){throw 'Help header style differs.'}
    if((Test-StyleSignature $help.Cells.Item($h+1,2)) -cne (Test-StyleSignature $help.Cells.Item($hx+1,2))){throw 'Help paragraph style differs.'}
    if($help.Range('A'+$h+':F'+($hx-1)).Borders.Item(10).LineStyle -eq -4142){throw 'Help right border is absent.'}
    $link=$config.Cells.Item($r,5).Hyperlinks.Item(1)
    if($link.SubAddress -cne ("'Справка'!A"+$h)){throw 'Wrong Help link.'}
    $pictures=@($help.Shapes | Where-Object Type -eq 13)
    if($pictures.Count -ne 76 -or @($pictures | Where-Object {$_.Height -le 0 -or $_.Width -le 0}).Count){throw 'Help pictures were lost.'}
    $sourceFiles=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
    $verified=0
    foreach($file in $sourceFiles){
        $body=[regex]::Match([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if(-not $body){continue}
        $module=$book.VBProject.VBComponents.Item($file.BaseName).CodeModule
        $actual=[string]$module.Lines(1,$module.CountOfLines)
        if((Test-NormalizeCode $actual) -cne (Test-NormalizeCode $body)){throw ('VBA differs from source: '+$file.BaseName)}
        $verified++
    }
    Write-Output ('VBA_SOURCE_VERIFIED: '+$verified)
    $code=New-Object 'System.Collections.Generic.List[string]'
    $code.Add('VBA PROJECT EXPORT');$code.Add('Workbook: RC_Section_NDM.xlsm');$code.Add('Date: '+(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    foreach($component in $book.VBProject.VBComponents){
        $module=$component.CodeModule;$code.Add(('='*100));$code.Add('COMPONENT: '+$component.Name);$code.Add('TYPE: '+$component.Type);$code.Add('LINES: '+$module.CountOfLines);$code.Add(('='*100))
        if($module.CountOfLines){$code.Add([string]$module.Lines(1,$module.CountOfLines))}
    }
    [IO.File]::WriteAllText((Join-Path $directory 'VBA_All_Code.txt'),($code -join "`r`n"),[Text.UTF8Encoding]::new($true))
    $probe=$book.VBProject.VBComponents.Add(1);$probe.Name='modTestAutomaticWorkbookProbe'
    $probe.CodeModule.AddFromString(@'
Option Explicit
' ==========================================================================
' ДЛЯ ТЕСТОВ: разовая проверка четырёх типов в несохраняемой копии
' ==========================================================================
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет Auto 0.1 по реальным General/Main/Extra, сравнивая все волокна
' с ручным эталоном того же шага; исходный Config восстанавливается при любом исходе.
Public Function TestAutomaticWorkbookModels() As String
    Dim table As Object, saved As Variant, settings As CSystemSettingsReader, units As CUnitSystem
    Dim definition As CSectionDefinition, geometry As CGeometryRegion, builder As CSectionBuilder, model As CSectionModel
    Dim expected As CFiberMeshBuilder, actual As CFiberMeshBuilder, options As TSectionMeshOptions
    Dim kind As Variant, r As Long, i As Long, step As Double, result As String, code As Long, reason As String, solves As Long
    Set table = ThisWorkbook.Names("rngSystemSettings").RefersToRange: saved = table.Formula
    solves = SectionEquilibriumSolveCount()
    On Error GoTo Failed
    For Each kind In RegisteredSectionTypes()
        For r = 2 To table.Rows.Count
            Select Case CStr(table.Cells(r, 1).Value2)
                Case "Geometry.Type": table.Cells(r, 2).Value2 = CStr(kind)
                Case "Geometry.Source": table.Cells(r, 2).Value2 = "Generated"
            End Select
        Next r
        Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook ThisWorkbook
        Set units = New CUnitSystem: units.LoadFromSettings settings
        Set definition = ReadWorkbookSectionDefinition(ThisWorkbook, settings, units)
        Set geometry = New CGeometryRegion: geometry.Initialize definition.ConcreteDomain
        step = geometry.Width: If geometry.Height < step Then step = geometry.Height
        step = step * 0.1
        options = ReadSectionMeshOptions(settings, units)
        If options.AutoStepFactor <> 0.1 Or options.BoundarySubdivisions <> 2 Then Err.Raise 5, , "Wrong default mesh options."
        Set builder = New CSectionBuilder: Set model = builder.Build(definition, options): Set actual = builder.Mesh
        Set expected = New CFiberMeshBuilder: expected.BuildMesh geometry, step, step, 1, 2
        If expected.FiberCount <> actual.FiberCount Then Err.Raise 5, , "Different generated mesh count."
        For i = 1 To expected.FiberCount
            If expected.FiberX(i) <> actual.FiberX(i) Or expected.FiberY(i) <> actual.FiberY(i) Or _
                    expected.FiberWidth(i) <> actual.FiberWidth(i) Or expected.FiberHeight(i) <> actual.FiberHeight(i) Then Err.Raise 5, , "Different generated fiber."
        Next i
        result = result & CStr(kind) & ": step=" & CStr(step) & " mm; concrete=" & CStr(model.ConcreteCount) & "; rebar=" & CStr(model.RebarCount) & vbCrLf
    Next kind
    If SectionEquilibriumSolveCount() <> solves Then Err.Raise 5, , "Unexpected equilibrium solve."
    table.Formula = saved: TestAutomaticWorkbookModels = result: Exit Function
Failed:
    code = Err.Number: reason = Err.Description: table.Formula = saved: Err.Raise code, , reason
End Function
'@
    )
    $models=[string]$app.Run(("'"+$book.Name+"'!modTestAutomaticWorkbookProbe.TestAutomaticWorkbookModels"))
    $book.VBProject.VBComponents.Remove($probe)
    Write-Output 'AUTO_MODELS_BUILT'
    $loads=$book.Names.Item('rngLoadCombinations').RefersToRange
    $savedLoads=$loads.Formula;$count=[int]$loads.Rows.Count-1
    $ids=[object[,]]::new($count,1)
    for($i=0;$i -lt $count;$i++){$ids[$i,0]='TEST_'+($i+1).ToString('00')}
    try {
        # COM-адаптер PowerShell может запомнить Value2 как String после чтения
        # пустого шаблона. Здесь всего 30 тестовых ID; скаляры обходят его ошибку
        # приведения массива. Рабочая блочная запись книги не изменяется.
        for($i=0;$i -lt $count;$i++){$loads.Cells.Item($i+2,1).Value2=[string]$ids[$i,0]}
        $app.Calculate()
        $helperColumn=Get-WorkbookLayoutHelperColumn 'LoadCases'
        $actualIds=$config.Cells.Item(2,$helperColumn).Resize($count,1).Value2
        for($i=1;$i -le $count;$i++){
            $actual=if($count -eq 1){[string]$actualIds}else{[string]$actualIds[$i,1]}
            if($actual -cne [string]$ids[($i-1),0]){throw ('Load picker lost combination '+$i)}
        }
    } finally {$loads.Formula=$savedLoads;$app.Calculate()}
    Write-Output 'LOAD_PICKER_VERIFIED'
    $help.Outline.ShowLevels(1);$app.EnableEvents=$true;$config.Activate();$link.Follow()
    $section=Test-FindKeyRow $help 'Справка по настройкам Config и листа «Сечения»'
    $end=[int]$help.UsedRange.Row+[int]$help.UsedRange.Rows.Count-1
    $hidden=$help.Range('A'+($section+1)+':A'+$end).EntireRow.Hidden
    if($null -eq $hidden -or $hidden -or $app.ActiveWindow.ScrollRow -ne $h){throw 'Help navigation did not reveal the entire section at the target row.'}
    $help.Rows.Item($section).ShowDetail=$false
    if(-not $help.Range('A'+($section+1)+':A'+$end).EntireRow.Hidden){throw 'Help section did not collapse completely.'}
    $app.EnableEvents=$false;$help.Outline.ShowLevels(8)
    Export-TestView $help ('A'+$h+':F'+($hx-1)) (Join-Path $directory 'Help.png')
    Export-TestView $config ('A'+($r-2)+':E'+($boundary+2)) (Join-Path $directory 'Config.png')
    $layout=Get-WorkbookLayoutSnapshot $book $xml
    $layout | Add-Member Provenance ([ordered]@{Workbook=[IO.Path]::GetFileName($path);SHA256=$hash;Captured=(Get-Date).ToString('s')})
    $layout | ConvertTo-Json -Depth 15 | Set-Content -LiteralPath (Join-Path $directory 'WorkbookLayout.json') -Encoding UTF8
    [ordered]@{Status='PASS';WorkbookSHA256=$hash;Default='Auto 0.1';BoundarySubdivisions=2;ConfigRow=$r;HelpRow=$h;VbaSourceModulesVerified=$verified;HelpPictures=$pictures.Count;HelpStyleMatches=$true;HelpNavigation=$true;HelpWholeSectionExpandCollapse=$true;GeneratedModels=$models;LoadCasePickerRows=$count;NumericalSolves=0} |
        ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $directory 'WorkbookVerification.json') -Encoding UTF8
    Write-Output $models
    Write-Output 'AUTOMATIC_MESH_WORKBOOK_PASS'
} catch {
    Write-Output $_.ScriptStackTrace
    Write-Output $_.InvocationInfo.PositionMessage
    throw
} finally {
    if(Test-OwnedExcel){
        if($book){$book.Close($false)}
        if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if(Test-OwnedExcel){Stop-Process -Id $owned -Force -ErrorAction SilentlyContinue}
    }
}
if((Get-FileHash -LiteralPath $path).Hash -cne $hash){throw 'Verification changed the workbook.'}
