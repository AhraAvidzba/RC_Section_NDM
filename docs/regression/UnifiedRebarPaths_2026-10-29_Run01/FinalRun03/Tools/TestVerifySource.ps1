# Сверяет read-only VBE-экспорт со всеми текущими исходными компонентами.
# Модули листов/книги генерируются сборщиком и перечисляются отдельно.
param([Parameter(Mandatory=$true)][string]$ExportPath,
      [Parameter(Mandatory=$true)][string]$ReportPath,
      [string]$OverridesManifest='',
      [ValidateSet(0,1,3)][int]$ExpectedOverrideCount=0)
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSHOME 'Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1') -Global -ErrorAction Stop
$root='C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM'
$export=(Resolve-Path -LiteralPath $ExportPath).Path
$report=[IO.Path]::GetFullPath($ReportPath)
$allowed=(Resolve-Path -LiteralPath $PSScriptRoot).Path+[IO.Path]::DirectorySeparatorChar
foreach($path in @($export,$report)){if(-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence path outside own directory.'}}
$lines=[IO.File]::ReadAllLines($export,[Text.Encoding]::UTF8)
$components=@{}; $types=@{}
for($i=0;$i -lt $lines.Length;$i++) {
    if (-not $lines[$i].StartsWith('COMPONENT: ')) {continue}
    $name=$lines[$i].Substring(11)
    if ($components.ContainsKey($name)) {throw "Duplicate exported component: $name"}
    $type=[int]$lines[$i+1].Substring(6)
    $count=[int]$lines[$i+2].Substring(7)
    $body=''
    if ($count -gt 0) {$body=($lines[($i+4)..($i+3+$count)] -join "`n")}
    $components[$name]=$body; $types[$name]=$type
    $i=$i+3+$count
}
# VBE печатает Double-константы с 15 значащими цифрами, меняет регистр имен
# и хранит комментарии в системной ANSI-кодировке. Строковые литералы и
# пробелы сравниваются точно; целые константы сохраняют собственный тип.
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
public static class PostAuditVbaCanonicalText {
    private static readonly Regex Number = new Regex(@"(?<![\p{L}\p{N}_&.])(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?#?(?![\p{L}\p{N}_])");
    // Нормализует только регистр кода и печатное представление Double VBE.
    private static string Code(string chunk) {
        return Number.Replace(chunk.ToLowerInvariant(), delegate(Match match) {
            string token=match.Value;
            int next=match.Index+match.Length;
            if(token.IndexOfAny(new char[]{'.','e','#'})<0 ||
               (next<chunk.Length && "!@%&^".IndexOf(chunk[next])>=0)) return token;
            double value=Double.Parse(token.TrimEnd('#'),CultureInfo.InvariantCulture);
            return "<double:"+value.ToString("G15",CultureInfo.InvariantCulture)+">";
        });
    }
    // Сохраняет строковые литералы и отделяет комментарии от исполняемого кода.
    public static string Normalize(string body) {
        StringBuilder result=new StringBuilder(body.Length);
        Encoding ansi=Encoding.GetEncoding(1251);
        foreach(string line in body.Split('\n')) {
            bool quoted=false;
            StringBuilder chunk=new StringBuilder();
            for(int i=0;i<line.Length;i++) {
                char character=line[i];
                if(!quoted && character=='\'') {
                    result.Append(Code(chunk.ToString().TrimEnd(' ','\t'))); chunk.Clear();
                    result.Append(ansi.GetString(ansi.GetBytes(line.Substring(i)))); break;
                }
                if(character=='"') {
                    if(!quoted) {result.Append(Code(chunk.ToString())); chunk.Clear();}
                    result.Append(character);
                    if(quoted && i+1<line.Length && line[i+1]=='"') result.Append(line[++i]);
                    else quoted=!quoted;
                } else if(quoted) result.Append(character);
                else chunk.Append(character);
            }
            result.Append(Code(chunk.ToString())); result.Append('\n');
        }
        return result.ToString();
    }
}
'@
# Разделяет код, строковые литералы и комментарии без удаления содержимого.
# Двойная кавычка внутри строки не закрывает строковый литерал.
function ConvertTo-VbaCodeCase([string]$Body) {
    return [PostAuditVbaCanonicalText]::Normalize($Body)
}
# ДЛЯ ТЕСТОВ: нормализация не скрывает измененные строки, числа, типы и комментарии.
$selfChecks=[ordered]@{
    CodeCase=(ConvertTo-VbaCodeCase 'Dim Target As Double') -ceq (ConvertTo-VbaCodeCase 'Dim target As Double')
    DoubleFormatting=(ConvertTo-VbaCodeCase 'x = 0.45# + 1E-8') -ceq (ConvertTo-VbaCodeCase 'x = 0.45 + 0.00000001')
    StringCasePreserved=(ConvertTo-VbaCodeCase 'x = "Target"') -cne (ConvertTo-VbaCodeCase 'x = "target"')
    StringNumberPreserved=(ConvertTo-VbaCodeCase 'x = "0.45#"') -cne (ConvertTo-VbaCodeCase 'x = "0.45"')
    EscapedQuotePreserved=(ConvertTo-VbaCodeCase 'x = "a""B"') -cne (ConvertTo-VbaCodeCase 'x = "a""b"')
    CommentPreserved=(ConvertTo-VbaCodeCase "x = 1 ' Target") -cne (ConvertTo-VbaCodeCase "x = 1 ' target")
    NumberChangeRejected=(ConvertTo-VbaCodeCase 'x = 0.45') -cne (ConvertTo-VbaCodeCase 'x = 0.46')
    IntegerTypePreserved=(ConvertTo-VbaCodeCase 'x = 1#') -cne (ConvertTo-VbaCodeCase 'x = 1')
    VbeCommentPadding=(ConvertTo-VbaCodeCase "x = 1 ' same") -ceq (ConvertTo-VbaCodeCase "x = 1    ' same")
    WhitespacePreserved=(ConvertTo-VbaCodeCase 'x = 1') -cne (ConvertTo-VbaCodeCase 'x=1')
    CommentAnsiFormatting=(ConvertTo-VbaCodeCase "' 6Ø32") -ceq (ConvertTo-VbaCodeCase "' 6O32")
}
if(@($selfChecks.Values | Where-Object {-not $_}).Count -gt 0) {throw 'VBA canonical comparison self-test failed.'}
$overrides=@{}
if($OverridesManifest){
    $manifest=Get-Content -LiteralPath $OverridesManifest -Raw -Encoding UTF8 | ConvertFrom-Json
    if(@($manifest.Modules).Count -ne $ExpectedOverrideCount -or $ExpectedOverrideCount -eq 0){throw 'Explicit override count does not match the declared bounded plan.'}
    foreach($item in $manifest.Modules){
        if($overrides.ContainsKey([string]$item.Module) -or [string]$item.Module -cne 'modAutoCADStressExport' -or [int]$item.Type -ne 1 -or [IO.Path]::GetExtension([string]$item.Source) -ine '.bas'){throw 'Only the bounded Type1 exporter override is allowed.'}
        foreach($pair in @(@([string]$item.Source,[string]$item.OriginalSHA256),@([string]$item.Candidate,[string]$item.CandidateSHA256))){if((Get-FileHash -LiteralPath $pair[0]).Hash -cne $pair[1]){throw 'Original/candidate SHA guard failed.'}}
        if([IO.Path]::GetFileNameWithoutExtension([string]$item.Source) -cne [string]$item.Module){throw 'Source/module identity mismatch.'}
        $overrides[[string]$item.Module]=$item
    }
}
if($overrides.Count -ne $ExpectedOverrideCount){throw 'Required explicit overrides were not supplied.'}
$records=@(); $missing=@(); $differences=@(); $rawDifferences=@(); $sourceNames=@{}
foreach($file in Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.cls','.bas') {
    $name=$file.BaseName; $sourceNames[$name]=$true
    if(-not $components.ContainsKey($name)) {$missing+=$name; continue}
    $expectedFile=$file.FullName
    if($overrides.ContainsKey($name)){$expectedFile=[string]$overrides[$name].Candidate;if($types[$name] -ne [int]$overrides[$name].Type){throw 'Native override component type mismatch.'}}
    $source=[regex]::Match([IO.File]::ReadAllText($expectedFile,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
    $actual=[regex]::Match($components[$name],'(?ms)^Option Explicit.*').Value
    if($name -eq 'modSectionRecipeApiGuide') {
        $source=([IO.File]::ReadAllLines($file.FullName,[Text.Encoding]::UTF8) | Where-Object {$_ -notmatch '^Attribute VB_Name = '}) -join "`n"
        $actual=$components[$name]
        if(@($source -split "`n" | Where-Object {$_.Trim() -and $_ -notmatch "^\s*'"}).Count){throw 'API guide must contain comments only.'}
    }
    if(-not $source) {throw "Source has no Option Explicit: $name"}
    $source=[regex]::Replace($source,'\r?\n',"`n").TrimEnd()
    $actual=[regex]::Replace($actual,'\r?\n',"`n").TrimEnd()
    $rawEqual=$source -ceq $actual
    if(-not $rawEqual) {$rawDifferences+=$name}
    $expectedType=1; if($file.Extension -eq '.cls') {$expectedType=2}
    $equal=$types[$name] -eq $expectedType -and (ConvertTo-VbaCodeCase $source) -ceq (ConvertTo-VbaCodeCase $actual)
    if(-not $equal) {$differences+=$name}
    $records+=@{Name=$name; Equal=$equal; RawEqual=$rawEqual; ComponentType=$types[$name]; ExpectedType=$expectedType; SourceSHA256=(Get-FileHash -LiteralPath $expectedFile -Algorithm SHA256).Hash; OriginalSourceSHA256=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash; ExplicitOverride=$overrides.ContainsKey($name)}
}
$extra=@($components.Keys | Where-Object {-not $sourceNames.ContainsKey($_) -and $types[$_] -ne 100})
$documents=@($components.Keys | Where-Object {$types[$_] -eq 100})
$documentMatches=@();$workbookDocuments=0
$eventTemplate=[IO.File]::ReadAllText((Join-Path $root 'tools/build_workbook/WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)
foreach($name in $documents){
    $actual=$components[$name].TrimEnd();$isWorkbook=$actual -match '(?i)Private Sub Workbook_Open\('
    if($isWorkbook){$expected=$eventTemplate.TrimEnd();$workbookDocuments++}else{$expected="Option Explicit`n`n' Лист книги: отображение пользовательских данных и результатов."}
    $equal=(ConvertTo-VbaCodeCase ([regex]::Replace($expected,'\r?\n',"`n"))) -ceq (ConvertTo-VbaCodeCase ([regex]::Replace($actual,'\r?\n',"`n")))
    if(-not $equal){$differences+=$name}
    $documentMatches+=,[ordered]@{Name=$name;WorkbookEvents=[bool]$isWorkbook;Equal=[bool]$equal}
}
if($workbookDocuments -ne 1 -or $documents.Count -ne 6){$differences+='DocumentInventory'}
$passed=$missing.Count -eq 0 -and $differences.Count -eq 0 -and $extra.Count -eq 0
[ordered]@{Passed=$passed; SourceComponents=$sourceNames.Count; ExportComponents=$components.Count;
    Missing=$missing; Differences=$differences; VbeCodeCaseDifferences=$rawDifferences; UnexpectedCodeComponents=$extra;
    GeneratedDocumentModules=$documents; DocumentSourceMatches=$documentMatches; ExportSHA256=(Get-FileHash -LiteralPath $export -Algorithm SHA256).Hash;
    Normalization='VBE code case; Double G15 formatting; CP1251 comments; line endings/module tail; VBE padding before inline comments; exact strings/other whitespace/integer types';
    VerifierSelfChecks=$selfChecks; Components=$records} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $report -Encoding UTF8
Write-Output "VBA_EQUALITY: passed=$passed; source=$($sourceNames.Count); export=$($components.Count); missing=$($missing.Count); differences=$($differences.Count); extra=$($extra.Count)"
if(-not $passed) {exit 1}
