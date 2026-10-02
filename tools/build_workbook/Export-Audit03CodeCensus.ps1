# Read-only VBA ownership/comment/guard census. This is an inspection index,
# not a replacement for semantic review or runtime acceptance.
param([string]$OutputPrefix = 'docs/regression/Audit03/code_census_2026-10-02')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$modules = New-Object System.Collections.Generic.List[object]
$methods = New-Object System.Collections.Generic.List[object]
$guards = New-Object System.Collections.Generic.List[object]
$templatePatterns = @(
    'Проверяет отдельный расчетный или интерфейсный сценарий и фиксирует ожидаемое поведение регрессией',
    'Запускает связанный набор операций и возвращает пользователю итоговый статус выполнения',
    'Возвращает сохраненное значение',
    'Обновляет сохраненное значение',
    'Создает расчетный или интерфейсный объект из нормализованных исходных данных и локальных настроек'
)
$files = @(Get-ChildItem -LiteralPath (Join-Path $root 'src'), (Join-Path $root 'tests') -Recurse -File |
    Where-Object { $_.Extension -in @('.cls', '.bas', '.frm') } | Sort-Object FullName)
foreach ($file in $files) {
    $relative = $file.FullName.Substring($root.Length + 1).Replace('\', '/')
    $source = [IO.File]::ReadAllLines($file.FullName, [Text.Encoding]::UTF8)
    $name = $file.BaseName
    $optionLine = -1
    for ($i = 0; $i -lt $source.Length; $i++) {
        if ($source[$i] -match '^Option Explicit\s*$') { $optionLine = $i; break }
    }
    $header = New-Object System.Collections.Generic.List[string]
    if ($optionLine -ge 0) {
        for ($i = $optionLine + 1; $i -lt $source.Length; $i++) {
            if ([string]::IsNullOrWhiteSpace($source[$i])) { continue }
            if ($source[$i] -notmatch "^\s*'") { break }
            $header.Add(($source[$i] -replace "^\s*'\s?", ''))
        }
    }
    $methodCount = 0
    $templateCount = 0
    for ($i = 0; $i -lt $source.Length; $i++) {
        $line = $source[$i]
        if ($line -match '^\s*(Public|Private|Friend)\s+(?:(Static)\s+)?(Sub|Function|Property Get|Property Let|Property Set)\s+([A-Za-z_][A-Za-z0-9_]*)') {
            $visibility = $Matches[1]
            $kind = $Matches[3]
            $methodName = $Matches[4]
            $comment = New-Object System.Collections.Generic.List[string]
            for ($previous = $i - 1; $previous -ge 0; $previous--) {
                if ($source[$previous] -match "^\s*'\s?(.*)$") { $comment.Insert(0, $Matches[1]); continue }
                if ([string]::IsNullOrWhiteSpace($source[$previous]) -and $comment.Count -eq 0) { continue }
                break
            }
            $text = $comment -join ' '
            $template = $false
            foreach ($pattern in $templatePatterns) { if ($text.Contains($pattern)) { $template = $true; break } }
            if ($template) { $templateCount++ }
            $methodCount++
            $methods.Add([pscustomobject]@{
                File = $relative; Module = $name; SourceLine = $i + 1; Visibility = $visibility;
                Kind = $kind; Method = $methodName; CommentLines = $comment.Count;
                Comment = $text; TemplateSuspected = $template; SemanticReview = 'Pending'
            })
        }
        if ($line -match '^\s*[^'']*(?:\bAnd\b|\bOr\b|\bIIf\s*\(|\bLBound\s*\(|\bUBound\s*\()') {
            $guards.Add([pscustomobject]@{
                File = $relative; SourceLine = $i + 1; Expression = $line.Trim();
                CallPathReview = 'Pending'; TestId = ''; Evidence = ''
            })
        }
    }
    $modules.Add([pscustomobject]@{
        File = $relative; Module = $name; Kind = $file.Extension;
        Lines = $source.Length; Header = $header -join ' '; HeaderLines = $header.Count;
        Methods = $methodCount; SuspectedTemplates = $templateCount; OwnerReview = 'Pending'
    })
}
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
$modules | Export-Csv -LiteralPath ($prefix + '_modules.csv') -NoTypeInformation -Encoding UTF8
$methods | Export-Csv -LiteralPath ($prefix + '_methods.csv') -NoTypeInformation -Encoding UTF8
$guards | Export-Csv -LiteralPath ($prefix + '_guards.csv') -NoTypeInformation -Encoding UTF8
[pscustomobject]@{
    Date = [DateTime]::Now.ToString('s'); Modules = $modules.Count; Methods = $methods.Count;
    Guards = $guards.Count; MissingMethodComment = @($methods | Where-Object CommentLines -eq 0).Count;
    SuspectedTemplateMethods = @($methods | Where-Object TemplateSuspected).Count;
    SemanticReview = 'Pending'; ProductionClasses = @($modules | Where-Object { $_.Kind -eq '.cls' -and $_.File.StartsWith('src/') }).Count
} | ConvertTo-Json | Set-Content -LiteralPath ($prefix + '_summary.json') -Encoding UTF8
Write-Output "CODE_CENSUS: modules=$($modules.Count); methods=$($methods.Count); guards=$($guards.Count); acceptance=Pending"
