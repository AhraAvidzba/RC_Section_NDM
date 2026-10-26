# ДЛЯ ТЕСТОВ: только чтение структуры точной конечной XLSM в собственном Excel.
# Не импортирует VBA, не вызывает макросы, не сохраняет и не меняет данные книги.
[CmdletBinding()]
param(
    [string]$SourceWorkbook='workbook/output/RC_Section_NDM.xlsm',
    [string]$EvidenceDirectory='docs/regression/ContourCellGeometry_2026-10-26_Run01/FinalStructure01',
    [string]$PreservationReferenceWorkbook='',
    [int]$ExpectedComponentCount=181,
    [double]$MinimumFreeRAMGiB=6,
    [double]$MinimumFreeCommitGiB=10
)
$ErrorActionPreference='Stop'
if($PSVersionTable.PSVersion.Major -ne 5){throw 'Use Windows PowerShell 5.1.'}
if($MinimumFreeRAMGiB -lt 4 -or $MinimumFreeCommitGiB -lt 8){throw 'Minimum reserve is 4 GiB RAM / 8 GiB commit.'}
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../../..')).Path
if(-not [IO.Path]::IsPathRooted($SourceWorkbook)){$SourceWorkbook=Join-Path $root $SourceWorkbook}
$source=(Resolve-Path -LiteralPath $SourceWorkbook).Path
if(-not [IO.Path]::IsPathRooted($EvidenceDirectory)){$EvidenceDirectory=Join-Path $root $EvidenceDirectory}
$directory=[IO.Path]::GetFullPath($EvidenceDirectory)
$allowed=(Join-Path $root 'docs/regression')+[IO.Path]::DirectorySeparatorChar
if(-not $directory.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence must stay within docs/regression.'}
if(Test-Path -LiteralPath $directory){throw 'Use a fresh evidence directory.'}
[void][IO.Directory]::CreateDirectory($directory)
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$templatePath=Join-Path $root 'tools/build_workbook/WorkbookHelpEvents.vba'
$layoutPath=Join-Path $root 'tools/build_workbook/WorkbookLayout.json'
$template=[IO.File]::ReadAllText($templatePath,[Text.Encoding]::UTF8)
$layout=Get-Content -LiteralPath $layoutPath -Raw -Encoding UTF8 | ConvertFrom-Json
$templateHash=(Get-FileHash -LiteralPath $templatePath).Hash
$layoutHash=(Get-FileHash -LiteralPath $layoutPath).Hash
$reference=$null;$referenceHash=$null
if($PreservationReferenceWorkbook){
    if(-not [IO.Path]::IsPathRooted($PreservationReferenceWorkbook)){$PreservationReferenceWorkbook=Join-Path $root $PreservationReferenceWorkbook}
    $reference=(Resolve-Path -LiteralPath $PreservationReferenceWorkbook).Path
    $referenceHash=(Get-FileHash -LiteralPath $reference).Hash
}
$script:checks=New-Object 'System.Collections.Generic.List[object]'
function Write-Proof([string]$name,$value){$value | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $directory $name) -Encoding UTF8}
function Check([string]$name,[bool]$passed,$details){$script:checks.Add([ordered]@{Check=$name;Passed=$passed;Details=$details})}
function Exact-OwnProcess([int]$processID,[long]$ticks){
    if($processID -le 0 -or $ticks -le 0){return $false}
    $process=Get-Process -Id $processID -ErrorAction SilentlyContinue
    if(-not $process){return $false}
    try{return $process.ProcessName -ieq 'EXCEL' -and $process.StartTime.ToUniversalTime().Ticks -eq $ticks}catch{return $false}
}
function Normalize-Vba([string]$body){
    # VBE меняет регистр кода и отступ перед inline-комментарием; строки/комментарии сохраняются.
    $result=New-Object Text.StringBuilder
    foreach($line in ([regex]::Replace($body,'\r\n?',"`n").TrimEnd() -split "`n")){
        $quoted=$false;$chunk=New-Object Text.StringBuilder
        for($index=0;$index -lt $line.Length;$index++){
            $character=$line[$index]
            if(-not $quoted -and $character -eq "'"){
                [void]$result.Append($chunk.ToString().TrimEnd().ToLowerInvariant());[void]$chunk.Clear()
                [void]$result.Append($line.Substring($index));break
            }
            if($character -eq '"'){
                if(-not $quoted){[void]$result.Append($chunk.ToString().ToLowerInvariant());[void]$chunk.Clear()}
                [void]$result.Append($character)
                if($quoted -and $index+1 -lt $line.Length -and $line[$index+1] -eq '"'){[void]$result.Append($line[++$index])}else{$quoted=-not $quoted}
            }elseif($quoted){[void]$result.Append($character)}else{[void]$chunk.Append($character)}
        }
        [void]$result.Append($chunk.ToString().ToLowerInvariant());[void]$result.Append("`n")
    }
    return $result.ToString()
}
function Text-SHA([string]$text){
    $sha=[Security.Cryptography.SHA256]::Create()
    try{return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($text)))).Replace('-','')}finally{$sha.Dispose()}
}
function Read-Headers($anchor){
    # Одно блочное чтение строки; непрерывная шапка заканчивается первой пустой ячейкой.
    $values=$anchor.Resize(1,256).Value2;$items=New-Object 'System.Collections.Generic.List[string]'
    for($column=1;$column -le 256;$column++){
        $value=[string]$values[1,$column]
        if([string]::IsNullOrEmpty($value)){break}
        $items.Add($value)
    }
    return ,$items.ToArray()
}
function Read-PackageGeometryPresence([string]$path){
    # Только ZIP/XML: пустой сохранённый блок подтверждается независимым источником,
    # а не объявляется допустимым по одному текущему наблюдению Excel.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($path)
    try{
        function Read-Part([string]$part){
            $entry=$zip.GetEntry($part);if(-not $entry){throw ('Missing package part: '+$part)}
            $stream=$entry.Open();$reader=New-Object IO.StreamReader($stream,[Text.Encoding]::UTF8)
            try{return [xml]$reader.ReadToEnd()}finally{$reader.Dispose();$stream.Dispose()}
        }
        $workbook=Read-Part 'xl/workbook.xml'
        $ns=New-Object Xml.XmlNamespaceManager($workbook.NameTable);$ns.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
        $definition=$workbook.SelectSingleNode('//s:definedName[@name="rngNDMSectionGeometry"]',$ns)
        if(-not $definition){throw 'Reference lacks rngNDMSectionGeometry.'}
        $address=[string]$definition.InnerText;$parts=$address.TrimStart('=').Split('!')
        if($parts.Count -ne 2){throw 'Simple named geometry anchor required.'}
        $sheetName=$parts[0].Trim("'").Replace("''","'");$cell=$parts[1].Replace('$','')
        if($cell -notmatch '^([A-Z]+)([0-9]+)$'){throw 'Single-cell geometry anchor required.'}
        $firstRow=[int]$Matches[2];$firstColumn=0
        foreach($letter in $Matches[1].ToCharArray()){$firstColumn=$firstColumn*26+([int]$letter-[int][char]'A'+1)}
        $sheetNode=@($workbook.SelectNodes('//s:sheets/s:sheet',$ns) | Where-Object {$_.GetAttribute('name') -ceq $sheetName})[0]
        $relationID=$sheetNode.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $relations=Read-Part 'xl/_rels/workbook.xml.rels'
        $target=[string]@($relations.DocumentElement.ChildNodes | Where-Object {$_.GetAttribute('Id') -ceq $relationID})[0].GetAttribute('Target')
        if($target.StartsWith('/')){$part=$target.TrimStart('/')}else{$part='xl/'+$target.TrimStart('/')}
        $sheetXml=Read-Part $part
        $sns=New-Object Xml.XmlNamespaceManager($sheetXml.NameTable);$sns.AddNamespace('s',$ns.LookupNamespace('s'))
        $occupied=@();$headers=@()
        foreach($node in $sheetXml.SelectNodes('//s:sheetData/s:row/s:c',$sns)){
            $ref=$node.GetAttribute('r');if($ref -notmatch '^([A-Z]+)([0-9]+)$'){continue}
            $rowNumber=[int]$Matches[2];if($rowNumber -lt $firstRow){continue}
            $column=0;foreach($letter in $Matches[1].ToCharArray()){$column=$column*26+([int]$letter-[int][char]'A'+1)}
            if($column -lt $firstColumn -or $column -ge $firstColumn+16){continue}
            $value=$node.SelectSingleNode('s:v|s:is|s:f',$sns)
            if($null -eq $value){continue}
            $occupied+=,$ref;if($rowNumber -eq $firstRow){$headers+=,$ref}
        }
        return [ordered]@{Source=$path;SHA256=(Get-FileHash -LiteralPath $path).Hash;NamedAnchor=$address;PhysicalHeaderCellCount=$headers.Count;PhysicalGeometryValueOrFormulaCells=$occupied.Count;OccupiedCells=$occupied;Empty=($occupied.Count -eq 0)}
    }finally{$zip.Dispose()}
}
function Header-Borders($sheet,[int]$row,[int]$first,[int]$last){
    $records=@();$passed=$true
    for($column=$first;$column -le $last;$column++){
        $cell=$sheet.Cells.Item($row,$column);$edges=@()
        foreach($edge in @(7,8,9,10)){
            $style=[int]$cell.Borders.Item($edge).LineStyle
            $edges+=@{Edge=$edge;LineStyle=$style}
            if($style -eq -4142 -or $style -eq 0){$passed=$false}
        }
        $records+=@{Column=$column;Edges=$edges}
    }
    return [ordered]@{Passed=$passed;Row=$row;First=$first;Last=$last;Cells=$records}
}
function Check-Columns($sheet,$sheetLayout,[int[]]$numbers){
    foreach($column in $numbers){
        $expected=@($sheetLayout.Columns | Where-Object {$_.First -le $column -and $_.Last -ge $column})
        $width=[double]$sheetLayout.StandardWidth;$hidden=$false
        if($expected.Count -eq 1){$width=[double]$expected[0].Width;$hidden=[bool]$expected[0].Hidden}
        $actualWidth=[double]$sheet.Columns.Item($column).ColumnWidth;$actualHidden=[bool]$sheet.Columns.Item($column).Hidden
        Check ('column.'+$sheet.Name+'.'+$column) ([math]::Abs($actualWidth-$width) -le 0.015 -and $actualHidden -eq $hidden) ([ordered]@{ExpectedWidth=$width;ActualWidth=$actualWidth;ExpectedHidden=$hidden;ActualHidden=$actualHidden})
    }
}
if(-not ('NdmFinalStructureOwner' -as [type])){Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class NdmFinalStructureOwner{[DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);}'}
$existing=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object {[ordered]@{PID=$_.Id;StartTicksUTC=[string]$_.StartTime.ToUniversalTime().Ticks}})
$existingIDs=@($existing | ForEach-Object {$_.PID})
$app=$null;$book=$null;$owned=$false;$excelID=0;$ticks=0L;$failure=$null;$documents=@();$components=@();$tables=@();$closure=$null
Write-Proof 'Input.json' ([ordered]@{Source=$source;SourceSHA256=$sourceHash;Template=$templatePath;TemplateSHA256=$templateHash;LayoutSHA256=$layoutHash;PreservationReference=$reference;ReferenceSHA256=$referenceHash;ExistingExcel=$existing;ImportPerformed=$false;MacrosRun=$false;SavePerformed=$false;ExpectedComponents=$ExpectedComponentCount})
try{
    $memory=Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
    $ram=[double]$memory.AvailableMBytes/1024;$commit=([double]$memory.CommitLimit-[double]$memory.CommittedBytes)/1GB
    $ready=$ram -ge $MinimumFreeRAMGiB -and $commit -ge $MinimumFreeCommitGiB
    Write-Proof 'MemoryBeforeExcel.json' ([ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');FreeRAMGiB=$ram;FreeCommitGiB=$commit;RequiredRAMGiB=$MinimumFreeRAMGiB;RequiredCommitGiB=$MinimumFreeCommitGiB;Passed=$ready})
    if(-not $ready){throw 'Insufficient reserve; no Excel started.'}
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativeID=0;[void][NdmFinalStructureOwner]::GetWindowThreadProcessId([IntPtr]$app.Hwnd,[ref]$nativeID)
    $excelID=[int]$nativeID;$ticks=(Get-Process -Id $excelID).StartTime.ToUniversalTime().Ticks
    if($excelID -in $existingIDs){throw 'Refusing pre-existing Excel process.'}
    $owned=$true
    Write-Proof 'Ownership.json' ([ordered]@{ExcelPID=$excelID;ExcelStartTicksUTC=[string]$ticks;ReadOnlyVerification=$true;MacrosRun=$false;UserDocumentsUsed=$false})
    $app.Visible=$false;$app.DisplayAlerts=$false;$app.EnableEvents=$false;$app.AskToUpdateLinks=$false;$app.AutomationSecurity=3
    $book=$app.Workbooks.Open($source,0,$true)
    if(-not $book.ReadOnly){throw 'Final verification must open read-only.'}
    Check 'workbook.readOnly' ([bool]$book.ReadOnly) $source
    $actualSheets=@($book.Worksheets | ForEach-Object Name);$expectedSheets=@($layout.Sheets | ForEach-Object Name)
    Check 'workbook.sheets' (($actualSheets -join '|') -ceq ($expectedSheets -join '|')) ([ordered]@{Actual=$actualSheets;ExpectedFromWorkbookLayout=$expectedSheets})
    $config=$book.Worksheets.Item('Config');$help=$book.Worksheets.Item('Справка');$results=$book.Worksheets.Item('Results')
    $settings=$book.Names.Item('rngSystemSettings').RefersToRange;$data=$settings.Value2;$rows=@()
    for($row=1;$row -le $settings.Rows.Count;$row++){if([string]$data[$row,1] -ceq 'Mesh.CellGeometry'){$rows+=,$row}}
    Check 'setting.uniqueKey' ($rows.Count -eq 1) $rows
    if($rows.Count -ne 1){throw 'Mesh.CellGeometry must have exactly one row.'}
    $settingRow=$settings.Row+$rows[0]-1;$inputCell=$config.Cells.Item($settingRow,$settings.Column+1)
    Check 'setting.defaultAdaptive' ([string]$inputCell.Value2 -ceq 'Adaptive') ([string]$inputCell.Value2)
    $formula=[string]$inputCell.Validation.Formula1;$tokens=@($formula -split '[,;]' | ForEach-Object {$_.Trim()})
    Check 'setting.literalList' ([int]$inputCell.Validation.Type -eq 3 -and -not $formula.StartsWith('=') -and ($tokens -join '|') -ceq 'Adaptive|Contour') ([ordered]@{Type=[int]$inputCell.Validation.Type;Formula=$formula;Items=$tokens;InCellDropdown=[bool]$inputCell.Validation.InCellDropdown})
    Check 'setting.dropdownVisible' ([bool]$inputCell.Validation.InCellDropdown) $null
    Check 'setting.editableFill' ([int]$inputCell.Interior.Color -eq 15135487) ([int]$inputCell.Interior.Color)
    Check 'setting.centeredInput' ([int]$inputCell.HorizontalAlignment -eq -4108) ([int]$inputCell.HorizontalAlignment)
    $borders=Header-Borders $config $settingRow $settings.Column ($settings.Column+4)
    Check 'setting.rowBorders' $borders.Passed $borders
    $linkCell=$config.Cells.Item($settingRow,$settings.Column+4)
    Check 'setting.helpLinkCount' ($linkCell.Hyperlinks.Count -eq 1) $linkCell.Hyperlinks.Count
    if($linkCell.Hyperlinks.Count -ne 1){throw 'Mesh.CellGeometry help link missing.'}
    $subAddress=[string]$linkCell.Hyperlinks.Item(1).SubAddress
    $bang=$subAddress.LastIndexOf('!');if($bang -lt 1){throw 'Invalid help subaddress.'}
    $linkSheet=$subAddress.Substring(0,$bang).Replace("'",'');$helpAddress=$subAddress.Substring($bang+1)
    Check 'setting.helpTargetSheet' ($linkSheet -ceq 'Справка') $subAddress
    $helpTarget=$help.Range($helpAddress)
    Check 'setting.helpHeading' ([string]$helpTarget.Value2 -ceq 'Mesh.CellGeometry') ([string]$helpTarget.Value2)
    $helpBodyRange=$help.Range($help.Cells.Item([int]$helpTarget.Row+1,2),$help.Cells.Item([int]$helpTarget.Row+8,2))
    $helpBody=$helpBodyRange.Value2;$paragraphs=@()
    for($row=1;$row -le 8;$row++){$paragraphs+=,[string]$helpBody[$row,1]}
    $helpText=$paragraphs -join "`n"
    Check 'setting.helpModeAndStorage' ($helpText.Contains('Adaptive') -and $helpText.Contains('Contour') -and $helpText.Contains('rngNDMElementGeometry') -and $helpText.Contains('ElementID')) ([ordered]@{RangeAddress=$helpBodyRange.Address($false,$false);ValueType=$helpBody.GetType().FullName;Paragraphs=$paragraphs})
    $helpBorders=@();$helpBorderOK=$true
    for($row=$helpTarget.Row;$row -le $helpTarget.Row+8;$row++){
        $lineStyle=[int]$help.Cells.Item($row,6).Borders.Item(10).LineStyle;$helpBorders+=@{Row=$row;RightBorder=$lineStyle}
        if($lineStyle -eq -4142 -or $lineStyle -eq 0){$helpBorderOK=$false}
    }
    Check 'help.rightBorder' $helpBorderOK $helpBorders
    $contourHeaders=@('RunID','LoopID','SegmentID','Sequence','LoopRole','SegmentType','StartX','StartY','EndX','EndY','CenterX','CenterY','Radius','SweepAngle, rad','SourceID','Comment','LengthUnit (SectionXY)')
    $blocks=@(@('rngNDMSectionGeometry',16),@('rngNDMSectionContours',17),@('rngNDMElementGeometry',18))
    foreach($block in $blocks){
        $name=[string]$block[0];$columns=[int]$block[1];$anchor=$book.Names.Item($name).RefersToRange;$headers=Read-Headers $anchor
        Check ($name+'.sheet') ($anchor.Worksheet.Name -ceq 'Results') $anchor.Worksheet.Name
        if($name -eq 'rngNDMSectionGeometry' -and $headers.Count -eq 0){
            $inheritedEmpty=$false;$emptyProof=[ordered]@{ReferenceRequired=$true;PhysicalHeaderWasNotFabricated=$true;ContractWidth=16}
            if($reference){
                $beforePresence=Read-PackageGeometryPresence $reference;$afterPresence=Read-PackageGeometryPresence $source
                $inheritedEmpty=$beforePresence.Empty -and $afterPresence.Empty -and $beforePresence.NamedAnchor -ceq $afterPresence.NamedAnchor
                $emptyProof.Reference=$beforePresence;$emptyProof.Source=$afterPresence
            }
            Check ($name+'.emptyPreservedFromReference') $inheritedEmpty $emptyProof
            $tables+=@{Name=$name;Address=$anchor.Address($false,$false);Row=$anchor.Row;Column=$anchor.Column;Columns=$columns;LastColumn=$anchor.Column+$columns-1;PhysicalHeaderPresent=$false;InheritedEmptyProven=$inheritedEmpty}
            continue # Нет физической матрицы: её шапку/рамки нельзя объявлять проверенными на этой пользовательской книге.
        }
        Check ($name+'.headerWidth') ($headers.Count -eq $columns) $headers
        if($name -eq 'rngNDMSectionGeometry'){
            $common=@('RunID','ElementID','MaterialType','X','Y','Area','GeometryInterpretationStatus','Width','Height','Diameter','Rotation','LocalIx','LocalIy','LocalIxy','Comment','ExtraData')
            $actual=@($headers | ForEach-Object {($_ -split ',',2)[0]})
            Check ($name+'.commonHeaderContract') (($actual -join '|') -ceq ($common -join '|')) $actual
        }else{
            $expected=$contourHeaders;if($name -eq 'rngNDMElementGeometry'){$expected=$contourHeaders+@('ElementID')}
            Check ($name+'.exactHeaders') (($headers -join '|') -ceq ($expected -join '|')) $headers
        }
        $borderInfo=Header-Borders $results $anchor.Row $anchor.Column ($anchor.Column+$columns-1)
        Check ($name+'.allHeaderBorders') $borderInfo.Passed $borderInfo
        $tables+=@{Name=$name;Address=$anchor.Address($false,$false);Row=$anchor.Row;Column=$anchor.Column;Columns=$columns;LastColumn=$anchor.Column+$columns-1}
    }
    $element=@($tables | Where-Object Name -eq 'rngNDMElementGeometry')[0]
    Check 'elementGeometry.CR156' ($element.Row -eq 156 -and $element.Column -eq 96) $element
    $widths=@{'rngNDMElementResults'=9;'rngNDMSectionGeometry'=16;'rngNDMSectionContours'=17;'rngNDMSectionProperties'=6;'rngNDMMaterialDiagrams'=11;'rngNDMSectionAnnotations'=24}
    $priorRight=0;$priorRanges=@()
    foreach($name in $widths.Keys){$anchor=$book.Names.Item($name).RefersToRange;$right=$anchor.Column+$widths[$name]-1;$priorRight=[math]::Max($priorRight,$right);$priorRanges+=@{Name=$name;Address=$anchor.Address($false,$false);RightColumn=$right}}
    Check 'elementGeometry.lastTechnicalTable' ($element.Column -gt $priorRight+2) ([ordered]@{PreviousRightmostColumn=$priorRight;NewFirstColumn=$element.Column;OtherTables=$priorRanges})
    Check 'elementGeometry.titleMerge' ([bool]$results.Range('CR155').MergeCells -and [string]$results.Range('CR155').MergeArea.Address($false,$false) -ceq 'CR155:DI155') ([string]$results.Range('CR155').MergeArea.Address($false,$false))
    Check 'elementGeometry.title' ([string]$results.Range('CR155').Value2 -ceq 'Фактические границы бетонных элементов') ([string]$results.Range('CR155').Value2)
    foreach($sheetName in @('Config','Справка','Results')){
        $sheet=$book.Worksheets.Item($sheetName);$sheetLayout=@($layout.Sheets | Where-Object Name -ceq $sheetName)[0]
        $numbers=1..6;if($sheetName -eq 'Config'){$numbers=1..24}elseif($sheetName -eq 'Results'){$numbers=96..113}
        Check-Columns $sheet $sheetLayout $numbers
    }
    Check 'vba.componentCount' ($book.VBProject.VBComponents.Count -eq $ExpectedComponentCount) $book.VBProject.VBComponents.Count
    $workbookCodeName=[string]$book.CodeName
    $worksheetStub="Option Explicit`r`n`r`n' Лист книги: отображение пользовательских данных и результатов."
    foreach($component in $book.VBProject.VBComponents){
        $lineCount=[int]$component.CodeModule.CountOfLines;$components+=@{Name=[string]$component.Name;Type=[int]$component.Type;Lines=$lineCount}
        if([int]$component.Type -ne 100){continue}
        $body='';if($lineCount -gt 0){$body=[string]$component.CodeModule.Lines(1,$lineCount)}
        $expected=$worksheetStub;if([string]$component.Name -ceq $workbookCodeName){$expected=$template}
        $equal=(Normalize-Vba $body) -ceq (Normalize-Vba $expected)
        $documents+=@{Name=[string]$component.Name;Lines=$lineCount;IsWorkbook=[string]$component.Name -ceq $workbookCodeName;EqualToBuildTemplate=$equal;ActualCodeSHA256=(Text-SHA $body);ExpectedCodeSHA256=(Text-SHA $expected)}
        Check ('documentModule.'+$component.Name) $equal $documents[-1]
    }
    Check 'vba.sixDocumentModules' ($documents.Count -eq 6) @($documents | ForEach-Object Name)
    Check 'vba.oneWorkbookEventModule' (@($documents | Where-Object IsWorkbook).Count -eq 1) $workbookCodeName
    $book.Close($false);$book=$null
}catch{$failure=$_.Exception.ToString();Write-Proof 'Failure.json' ([ordered]@{ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');Reason=$failure})}
finally{
    if($owned -and (Exact-OwnProcess $excelID $ticks)){
        if($book){try{$book.Close($false)}catch{}}
        if($app){try{$app.Quit()}catch{}}
    }
    $book=$null;$config=$null;$help=$null;$results=$null;$settings=$null;$inputCell=$null;$linkCell=$null;$helpTarget=$null;$helpBodyRange=$null;$anchor=$null;$sheet=$null;$component=$null
    if($app){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $forced=$false
    if($owned -and (Exact-OwnProcess $excelID $ticks)){
        $process=Get-Process -Id $excelID;[void]$process.WaitForExit(3000)
        if(Exact-OwnProcess $excelID $ticks){Stop-Process -Id $excelID -Force;$forced=$true}
    }
    $closure=[ordered]@{PID=$excelID;StartTicksUTC=[string]$ticks;Owned=$owned;ResidualForceKilled=$forced;ExactIdentityAbsent=(-not (Exact-OwnProcess $excelID $ticks));MacrosRun=$false;SavePerformed=$false}
    Write-Proof 'ProcessClosure.json' $closure
}
$unchanged=(Get-FileHash -LiteralPath $source).Hash -ceq $sourceHash
$templateUnchanged=(Get-FileHash -LiteralPath $templatePath).Hash -ceq $templateHash
$layoutUnchanged=(Get-FileHash -LiteralPath $layoutPath).Hash -ceq $layoutHash
$referenceUnchanged=-not $reference -or (Get-FileHash -LiteralPath $reference).Hash -ceq $referenceHash
$failed=@($script:checks | Where-Object {-not $_.Passed})
$passed=-not $failure -and $failed.Count -eq 0 -and $unchanged -and $templateUnchanged -and $layoutUnchanged -and $referenceUnchanged -and $closure.ExactIdentityAbsent
Write-Proof 'Acceptance.json' ([ordered]@{Status=$(if($passed){'PASS_NATIVE_READONLY_FINAL_STRUCTURE'}else{'REVIEW_REQUIRED'});Passed=$passed;Source=$source;SourceSHA256=$sourceHash;SourceUnchanged=$unchanged;TemplateUnchanged=$templateUnchanged;LayoutUnchanged=$layoutUnchanged;ReferenceUnchanged=$referenceUnchanged;ImportPerformed=$false;MacrosRun=$false;SavePerformed=$false;FullRegressionPass=$false;FailedChecks=$failed;CheckCount=$script:checks.Count;Checks=$script:checks;DocumentModules=$documents;Components=$components;Tables=$tables;ProcessClosure=$closure;Failure=$failure})
if(-not $passed){throw 'Final structure check requires review; see primary Acceptance.json.'}
Get-Content -LiteralPath (Join-Path $directory 'Acceptance.json') -Raw
