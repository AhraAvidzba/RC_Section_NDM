# Связывает directed runtime с каждым фактическим полем таблицы СП35 и N/A.
# Справочные колонки остаются в знаменателе, но не становятся расчетными
# настройками. Структурная проверка не является нормативной верификацией.
param(
    [Parameter(Mandatory=$true)][string]$RegistryPath,
    [Parameter(Mandatory=$true)][string]$BehaviorReport,
    [Parameter(Mandatory=$true)][string]$InvalidReport,
    [Parameter(Mandatory=$true)][string]$DiagnosticReport,
    [Parameter(Mandatory=$true)][string]$InactiveReport,
    [Parameter(Mandatory=$true)][string]$OutputPrefix
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$allowed = [IO.Path]::GetFullPath((Join-Path $root 'docs/regression/Audit03')) + [IO.Path]::DirectorySeparatorChar
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
if (-not $prefix.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) { throw 'Evidence output is outside Audit03.' }
foreach ($suffix in @('.json','.csv')) {
    if (Test-Path -LiteralPath ($prefix+$suffix)) { throw 'Evidence already exists; choose a new prefix.' }
}
$registry = Get-Content -LiteralPath (Join-Path $root $RegistryPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$reports = @{}
$paths = @($BehaviorReport,$InvalidReport,$DiagnosticReport,$InactiveReport)
$gates = @('TOTAL_AUDIT03_SP35_BEHAVIOR','TOTAL_AUDIT03_SP35_INVALID','TOTAL_AUDIT03_SP35_DIAGNOSTIC','TOTAL_AUDIT03_INACTIVE_CONFIG')
$sourceHash = ''
for ($i=0; $i -lt $paths.Count; $i++) {
    $text = Get-Content -LiteralPath (Join-Path $root $paths[$i]) -Raw -Encoding UTF8
    if ($text -notmatch ('(?m)^'+$gates[$i]+': (?:fields=[0-9]+; )?passed=[1-9][0-9]*; failed=0\r?$') -or
        -not $text.Contains('WATCHDOG_COMPLETED: exit=0; source unchanged=True') -or
        $text -match '(?m)^FAIL:') { throw "Incomplete directed gate: $($paths[$i])" }
    $hash = [regex]::Match($text,'(?m)^SOURCE_SHA256: ([A-F0-9]{64})\r?$').Groups[1].Value
    if (-not $hash -or ($sourceHash -and $sourceHash -ne $hash)) { throw 'Directed reports do not share the same source workbook.' }
    $sourceHash = $hash
    $reports[$gates[$i]] = $text
}
$behavior = $reports[$gates[0]]
$invalid = $reports[$gates[1]]
$diagnostic = $reports[$gates[2]]
$inactive = $reports[$gates[3]]
if (-not $diagnostic.Contains('OK: audit03.table.missingName') -or
    -not $diagnostic.Contains('OK: audit03.table.recovery') -or
    $diagnostic -notmatch 'RESULTS_SAVE_REOPEN:[^\r\n]*equal=True;' -or
    $diagnostic -notmatch 'STATUS_STYLE_SAVE_REOPEN:[^\r\n]*equal=True') { throw 'Address/comment/save-reopen gate is incomplete.' }
$cells = @{}
foreach ($match in [regex]::Matches($behavior,'(?m)^SP35_BEHAVIOR: (audit03\.table\.behavior\.([0-9]+)\.([0-9]+))\|cell=([^|]+)\|reference=(True|False)\|')) {
    $cell = $match.Groups[4].Value
    if ($cells.ContainsKey($cell)) { throw "Duplicate table behavior: $cell" }
    $cells[$cell] = @{ Prefix=$match.Groups[1].Value; Row=[int]$match.Groups[2].Value; Column=[int]$match.Groups[3].Value; Reference=$match.Groups[5].Value -eq 'True' }
}
$fields = @($registry.Fields | Where-Object Block -eq 'rngSP35Table721')
if ($fields.Count -ne 144 -or $cells.Count -ne 144) { throw 'All 144 actual table cells must remain in scope.' }
foreach ($field in $fields) {
    $cell = ([string]$field.Address -split '!')[-1]
    if (-not $cells.ContainsKey($cell)) { throw "No table behavior at the actual address: $cell" }
    $entry = $cells[$cell]
    $assertions = if ($entry.Reference) { @('referenceM','referenceL','referenceNult') }
        elseif ($entry.Column -eq 3) { @('inputValid','changedWeight','active') }
        elseif ($entry.Column -eq 8) { @('activeL','unaffectedM','forceChanged') }
        else { @('activeM','unaffectedL','forceChanged') }
    foreach ($assertion in $assertions) {
        if ($behavior -notmatch ('(?m)^OK: '+[regex]::Escape($entry.Prefix+'.'+$assertion)+'(?:;|\r?$)')) { throw "Missing independent behavior assertion: $cell/$assertion" }
    }
    if (-not $entry.Reference) {
        for ($variant=0; $variant -lt 5; $variant++) {
            $test = "audit03.table.invalid.$($entry.Row).$($entry.Column).$variant"
            foreach ($suffix in @('status','notNumFail','comment')) {
                if ($invalid -notmatch ('(?m)^OK: '+[regex]::Escape($test+'.'+$suffix)+'\r?$')) { throw "Missing invalid table test: $test/$suffix" }
            }
        }
    }
    $field.Type = if ($entry.Reference) { 'Reference table value; not a solver input' } else { 'Positive dimensionless table value' }
    $field.InternalUnits = '-'
    $field.RuntimeDefault = 'No fallback for an active invalid table; inactive branch does not consume it'
    $field.ExpectedEffect = if ($entry.Reference) { 'Changing reference columns leaves interpolation and Nult unchanged' } else { 'Independent interpolation weights or phi affect the active SP35 branch; other branches remain independent' }
    $field.BlankAndErrorContract = if ($entry.Reference) { 'Not consumed; diagnostic reference values do not block the calculation' } else { 'Empty/text/CVErr/zero/negative -> InputErr; l0/i nodes strictly increase; error address follows named range' }
    $field.TestId = $entry.Prefix
    $field.Evidence = @($BehaviorReport,$InvalidReport,$DiagnosticReport)
    $field.CoverageStatus = if ($entry.Reference) { 'ReferenceBehaviorAccepted:NotIndependentCalculationInput' } else { 'DirectedNormativeTableBehaviorAccepted:NormativeTraceSeparate' }
}
$inactiveFields = @($registry.Fields | Where-Object Role -eq 'NotEditable:NotApplicableCell')
if ($inactiveFields.Count -ne 10) { throw 'All ten not-applicable fields must remain in scope.' }
foreach ($field in $inactiveFields) {
    $cell = ([string]$field.Address -split '!')[-1]
    for ($variant=0; $variant -lt 5; $variant++) {
        foreach ($suffix in @('accepted','samePayload')) {
            $test = "audit03.inactive.$cell.$variant.$suffix"
            if ($inactive -notmatch ('(?m)^OK: '+[regex]::Escape($test)+'\r?$')) { throw "Missing inactive consumer test: $test" }
        }
    }
    $field.Type = 'Not applicable display; no active numeric consumer'
    $field.RuntimeDefault = 'Not a calculation setting'
    $field.BlankAndErrorContract = 'Number/text/CVErr/Empty/NA formula do not enter the normalized input payload'
    $field.ExpectedEffect = 'No normalized input or engineering result change'
    $field.TestId = 'modTestBatchCalculation.RunAudit03InactiveConfigCellTests'
    $field.Evidence = @($InactiveReport)
    $field.CoverageStatus = 'InactiveConsumerAccepted:NotEditable'
}
$registry | Add-Member -NotePropertyName DirectedTableEvidenceSHA256 -NotePropertyValue $sourceHash -Force
$registry | ConvertTo-Json -Depth 50 | Set-Content -LiteralPath ($prefix+'.json') -Encoding UTF8
$registry.Fields | Select-Object Id,Block,Address,Role,Type,UserUnits,InternalUnits,RuntimeDefault,ExpectedEffect,TestId,
    @{N='Evidence';E={$_.Evidence -join '; '}},CoverageStatus | Export-Csv -LiteralPath ($prefix+'.csv') -NoTypeInformation -Encoding UTF8
Write-Output "NORMATIVE_TABLE_EVIDENCE: fields=$($registry.Fields.Count); tableCells=144; inactiveFields=10; fullAcceptance=False"
