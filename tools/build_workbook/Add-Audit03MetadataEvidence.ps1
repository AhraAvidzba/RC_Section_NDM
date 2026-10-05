# Связывает actual metadata assertions с точными адресами свежего census.
# Производные ID и объединенные продолжения принимаются только по своему
# контракту. Подпись N/A не доказывает неактивность потребителя или всей Config.
param(
    [Parameter(Mandatory=$true)][string]$RegistryPath,
    [Parameter(Mandatory=$true)][string]$MetadataReport,
    [Parameter(Mandatory=$true)][string]$OutputPrefix
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$prefix=[IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
if (-not $prefix.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence must stay in docs/regression/Audit03.' }
foreach ($suffix in @('.json','.csv')) { if (Test-Path -LiteralPath ($prefix+$suffix)) { throw 'Existing evidence must not be overwritten.' } }
$registry=Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$report=Get-Content -LiteralPath (Join-Path $root $MetadataReport) -Raw -Encoding UTF8
if (-not $report.Contains('TOTAL_AUDIT03_CONFIG_METADATA: passed=209; failed=0') -or
    -not $report.Contains('SOURCE_UNCHANGED: True') -or
    -not $report.Contains('SOURCE_SHA256: '+$registry.SHA256)) { throw 'Actual metadata gate is incomplete or belongs to another source.' }
$duration=@($registry.Fields | Where-Object { $_.Block -eq 'rngStabilityDurationLoads' -and $_.Role -eq 'DerivedFormula' })
$followers=@($registry.Fields | Where-Object Role -eq 'NotEditable:MergedFollower')
$labels=@($registry.Fields | Where-Object Role -eq 'NotEditable:NotApplicableCell')
if ($duration.Count -ne 30 -or $followers.Count -ne 16 -or $labels.Count -ne 10) { throw 'Metadata inventory differs from the tested scope.' }

# Требует конкретную положительную строку, а не общий счет успешных assertions.
# Неудачная или пропущенная ячейка остается непринятой и прекращает публикацию.
function Require-MetadataAssertion([string]$Id) {
    if ($report -notmatch ('(?m)^OK: '+[regex]::Escape($Id)+';')) { throw "Missing metadata assertion: $Id" }
}

for ($slot=1; $slot -le 30; $slot++) {
    $field=$duration[$slot-1]
    $cell=([string]$field.Address -split '!')[-1]
    foreach ($suffix in @('formula','id','blank','relocated')) { Require-MetadataAssertion "metadata.duration.Slot$slot.$suffix" }
    if ($report -notmatch ('(?m)^OK: metadata\.duration\.Slot'+$slot+'\.formula; '+[regex]::Escape($cell)+'\r?$')) { throw "Derived ID address mismatch: $cell" }
    $field.Type='ReadOnly derived CombinationID'
    $field.RuntimeDefault='Formula follows rngLoadCombinations slot; no independent runtime default'
    $field.ExpectedEffect='Changed/blank LC ID and relocated named range are reflected without solve'
    $field.TestId="metadata.duration.Slot$slot"
    $field.Evidence=@($MetadataReport)
    $field.CoverageStatus='MetadataAccepted:DerivedCombinationID'
}
foreach ($field in $followers) {
    $cell=([string]$field.Address -split '!')[-1]
    foreach ($suffix in @('merged','notAnchor','noValue','singleDropdown')) { Require-MetadataAssertion "metadata.follower.$cell.$suffix" }
    $field.Type='Merged follower: not independent input'
    $field.RuntimeDefault='Value and selection belong only to the merged anchor'
    $field.ExpectedEffect='Actual MergeArea, empty follower, one anchor dropdown; no additional solve'
    $field.TestId="metadata.follower.$cell"
    $field.Evidence=@($MetadataReport)
    $field.CoverageStatus='MetadataAccepted:MergedFollower'
}
foreach ($field in $labels) {
    $cell=([string]$field.Address -split '!')[-1]
    foreach ($suffix in @('marker','notFormula')) { Require-MetadataAssertion "metadata.notApplicable.$cell.$suffix" }
    $field.Type='Not applicable label: not independent input'
    $field.TestId="metadata.notApplicable.$cell"
    $field.Evidence=@($MetadataReport)
    $field.CoverageStatus='MetadataLabelAccepted:InactiveConsumerStillPending'
}
$registry | Add-Member -NotePropertyName MetadataReviewedFields -NotePropertyValue 46 -Force
$registry | Add-Member -NotePropertyName NotApplicableLabelsReviewedFields -NotePropertyValue 10 -Force
$registry | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath ($prefix+'.json') -Encoding UTF8
$registry.Fields | Select-Object Id,Block,Address,Role,Type,RuntimeDefault,ExpectedEffect,TestId,
    @{N='Evidence';E={ $_.Evidence -join '; ' }},CoverageStatus | Export-Csv -LiteralPath ($prefix+'.csv') -NoTypeInformation -Encoding UTF8
Write-Output "METADATA_EVIDENCE: fields=$($registry.Fields.Count); derivedOrFollower=46; labels=10; inactiveConsumerAcceptance=False; fullAcceptance=False"
