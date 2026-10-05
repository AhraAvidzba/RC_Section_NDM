# Проверяет архитектурные границы СП 35 и соответствие исходников реальному
# экспорту VBA опубликованной книги. Не заменяет native-расчетные тесты.
param(
    [string]$ExportPath = 'docs/regression/SP35/Publication/VBA_All_Code.txt',
    [string]$ReportPath = 'docs/regression/SP35/Publication/SourceContracts.txt'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$lines = New-Object 'System.Collections.Generic.List[string]'
$failed = 0

# Записывает отдельный контракт, сохраняя все ошибки до итоговой сводки.
function Assert-Contract([string]$Name, [bool]$Passed, [string]$Detail) {
    if (-not $Passed) { $script:failed++ }
    $lines.Add("CONTRACT|$Name|passed=$Passed|$Detail")
}

# VBE меняет регистр идентификаторов и представление Double. Нормализуем только
# экспортную оболочку и это представление; текст строк/комментариев сохраняется.
function Get-CanonicalBody([string]$Text) {
    $body = [regex]::Match($Text, '(?ms)^Option Explicit.*').Value.Trim()
    $encoding = [Text.Encoding]::GetEncoding(1251)
    $body = $encoding.GetString($encoding.GetBytes($body))
    $canonical = foreach ($line in $body -split "`r?`n") {
        [regex]::Replace($line.TrimEnd(), '"(?:[^"]|"")*"|''.*$|(?<![\w&])(?:\d+\.\d*|\d+)(?:[Ee][+-]?\d+)?#?(?!\w)', {
            param($token)
            if ($token.Value.StartsWith('"') -or $token.Value.StartsWith("'")) { return $token.Value }
            $number = [double]::Parse($token.Value.TrimEnd('#'), [Globalization.CultureInfo]::InvariantCulture)
            return $number.ToString('G15', [Globalization.CultureInfo]::InvariantCulture)
        })
    }
    return $canonical -join "`n"
}

# Читает исполняемый код для проверки зависимостей, не учитывая комментарии.
function Get-Code([string]$RelativePath) {
    return (Get-Content -LiteralPath (Join-Path $root $RelativePath) -Encoding UTF8 |
        Where-Object { $_ -notmatch "^\s*'" }) -join "`n"
}

$sources = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
    Where-Object Extension -in '.bas', '.cls')
$exportFile = (Resolve-Path -LiteralPath (Join-Path $root $ExportPath)).Path
$export = Get-Content -LiteralPath $exportFile -Raw -Encoding UTF8
$components = @{}
foreach ($match in [regex]::Matches($export, '(?ms)^COMPONENT: ([^\r\n]+)\r?\nTYPE: \d+\r?\nLINES: \d+\r?\n=+\r?\n(.*?)(?=^=+\r?\nCOMPONENT:|\z)')) {
    $components[$match.Groups[1].Value] = Get-CanonicalBody $match.Groups[2].Value
}
$matched = 0
foreach ($file in $sources) {
    $source = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    if ($source -notmatch 'Attribute\s+VB_Name\s*=\s*"([^"]+)"') { throw "Missing VB_Name: $file" }
    $name = $Matches[1]
    $same = $components.ContainsKey($name) -and [string]::Equals(
        (Get-CanonicalBody $source), $components[$name], [StringComparison]::OrdinalIgnoreCase)
    Assert-Contract "export.$name" $same 'Source and saved VBE body agree'
    if ($same) { $matched++ }
}

foreach ($file in 'src/Section/CSectionGeometryQuery.cls', 'src/Section/CConcreteRegion.cls', 'src/Section/CSectionContours.cls') {
    $code = Get-Code $file
    Assert-Contract "geometry.$file" ($code -notmatch 'CSectionSolver|CLimitSearch|WorksheetFunction|CResultStatusPolicy|SP35|SP63') 'No solver, standard or display-status dependency'
}
$preparation = Get-Code 'src/Crack/CSP35CrackData.cls'
Assert-Contract 'SP35.readyStateOnly' ($preparation -notmatch 'CSectionSolver|CLimitSearch|CStateProvider|\.Solve\b|WorksheetFunction') 'Prepared strain plane and stresses only'
foreach ($file in 'src/Excel/CCrackSummaryWriter.cls', 'src/Excel/CNDMResultsWriter.cls') {
    $code = Get-Code $file
    Assert-Contract "writer.$file" ($code -notmatch 'New CSectionSolver|New CLimitSearch|\.SetResult\b|\.ColumnWidth\s*=|\.StandardWidth\s*=') 'No solve, status mutation or runtime width changes'
}
$width = Get-Code 'src/Crack/CCrackWidthCalculator.cls'
foreach ($name in 'SP35ReinforcementRadiusFromData', 'SP35PsiFromData', 'SP35CrackWidthFromData', 'CrackWidthFromData') {
    $method = [regex]::Match($width, "(?ms)^Public Function $name\b.*?^End Function")
    Assert-Contract "formula.$name" ($method.Success -and $method.Value -notmatch 'CSectionState|CSectionSolver|\brs\w+\b|\.Solve\b|SetWidthResult') 'Pure numerical formula'
}
Assert-Contract 'noRemovedFormulaClass' (-not $components.ContainsKey('CCrackWidthFormulaCalculator')) 'No obsolete formula facade'
$exporter = Get-Code 'src/Excel/modAutoCADStressExport.bas'
$importer = Get-Code 'src/Excel/CAutoCADSectionModelImporter.cls'
$catalog = Get-Code 'tools/build_workbook/SettingsCatalog.ps1'
Assert-Contract 'onlyCommonContourLayers' (($exporter + $importer + $catalog) -notmatch 'AutoCAD\.Layer\.Contour') 'No retired layer setting in runtime or catalog'
$materialExport = [regex]::Match($exporter, '(?ms)^Private Function DrawSavedMaterialContours\b.*?^End Function')
Assert-Contract 'universalMaterialContourExport' ($materialExport.Success -and $materialExport.Value -notmatch 'SourceType|AutoCADImport|Generated') 'One saved geometry path for generated and imported contours'
$stability = Get-Code 'src/Stability/CStabilityCalculator.cls'
Assert-Contract 'stabilityNotMigratedToQuery' ($stability -notmatch 'CSectionGeometryQuery|CConcreteRegion') 'Original discrete stability pipeline remains independent'
$lines.Add("TOTAL_SP35_SOURCE_CONTRACTS: matchingModules=$matched; sourceModules=$($sources.Count); failed=$failed")
$lines | Set-Content -LiteralPath (Join-Path $root $ReportPath) -Encoding UTF8
$lines
if ($failed) { throw "SP35 source contracts failed: $failed" }
