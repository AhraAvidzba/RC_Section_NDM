# Builds a field-by-field inventory from the saved Config census and catalog.
# Unverified runtime semantics remain explicit; structural presence is not PASS.
param(
    [string]$CensusPath = 'docs/regression/Audit03/config_census_2026-10-02.json',
    [string]$OutputPath = 'docs/regression/Audit03/config_field_registry_2026-10-02.json'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$census = Get-Content -LiteralPath (Join-Path $root $CensusPath) -Raw -Encoding UTF8 | ConvertFrom-Json
$byAddress = @{}
foreach ($cell in $census.Cells) { $byAddress[$cell.Address] = $cell }
$ranges = @{}
foreach ($range in $census.NamedRanges) { $ranges[$range.Name] = $range.Rectangle }
$defaults = @{}
foreach ($group in (Get-SystemSettingsCatalog)) {
    foreach ($row in $group.Rows) { if ($row[0] -match '^[A-Za-z]') { $defaults[$row[0]] = $row[1] } }
}
foreach ($row in (Get-CalculationProfilesCatalog)) {
    if ($row.Key) {
        foreach ($profile in @('PR1','PR2','PR3','PR4')) { $defaults["$profile.$($row.Key)"] = $row[$profile] }
    }
}
foreach ($row in (Get-UnitSettingsCatalog)) {
    $defaults["Units.$($row[0]).Input"]=$row[1]
    $defaults["Units.$($row[0]).Output"]=$row[3]
}
foreach ($row in (Get-SignConventionSettingsCatalog)) { $defaults["Signs.$($row[0])"]=$row[1] }
foreach ($row in @((Get-SteelMaterialParametersCatalog)) + @((Get-ConcreteMaterialParametersCatalog))) {
    $defaults["$($row[0]).Compression"]=$row[1]
    $defaults["$($row[0]).Tension"]=$row[2]
}
foreach ($shape in (Get-GeometrySettingsCatalog)) {
    foreach ($row in $shape.Rows) {
        if ($row[0] -match '^[A-Za-z].*\.') { $defaults[$row[0]]=$row[1] }
    }
}
$fields = [System.Collections.Generic.List[object]]::new()
$registered = @{}

# Converts a numeric column without relying on Excel or the current locale.
function Get-CellAddress([int]$row, [int]$column) {
    $letters = ''
    while ($column -gt 0) {
        $column--
        $letters = [char](65 + $column % 26) + $letters
        $column = [int][Math]::Floor($column / 26)
    }
    return "$letters$row"
}

# Missing blank cells are represented explicitly rather than omitted from coverage.
function Get-Cell([int]$row, [int]$column) {
    $address = Get-CellAddress $row $column
    if ($byAddress.ContainsKey($address)) { return $byAddress[$address] }
    return [pscustomobject]@{ Address=$address; Value=''; Formula=''; RowHidden=$false; ColumnHidden=$false; Validations=@() }
}

# Records the physical field, scope and exact saved validation. Tests and unresolved
# contracts are deliberately not inferred from a nonempty value or an old suite.
function Add-Field([string]$id, [string]$block, [int]$row, [int]$column,
        [string]$units, [string]$consumer, [string]$activity, [string]$role = 'UserInput') {
    $cell = Get-Cell $row $column
    if ($registered.ContainsKey($cell.Address)) { throw "Duplicate field: $($cell.Address)." }
    $registered[$cell.Address] = $true
    $validation = @($census.Validations | Where-Object { $cell.Validations -contains $_.Range })
    $type = 'TextOrNumber:PendingContractReview'
    if ($validation.Count -gt 0) { $type = ($validation.Type | Select-Object -Unique) -join ',' }
    if ($cell.Formula) { $role = 'DerivedFormula'; $type = 'Formula' }
    if ($cell.Value -eq '-') { $role = 'NotEditable:NotApplicableCell' }
    $mergeRange = ''
    $mergeAnchor = ''
    foreach ($merged in $census.MergedRanges) {
        $rectangle = $merged.Rectangle
        if ($row -ge $rectangle.Top -and $row -le $rectangle.Bottom -and
            $column -ge $rectangle.Left -and $column -le $rectangle.Right) {
            $mergeRange = $merged.Range
            $mergeAnchor = $merged.Anchor
            if ($cell.Address -ne $mergeAnchor) { $role = 'NotEditable:MergedFollower'; $type = 'MergedFollower' }
            break
        }
    }
    $default = 'PendingCatalogMapping'
    if ($defaults.ContainsKey($id)) { $default = $defaults[$id] }
    $helpAddress = ''
    if ($block -eq 'rngSystemSettings') { $helpAddress="E$row" }
    if ($block -in @('rngPlotAnnotationSettings','rngSteelMaterialParameters','rngConcreteMaterialParameters')) { $helpAddress="M$row" }
    $help = @($census.Hyperlinks | Where-Object { $helpAddress -and $_.Range -eq $helpAddress })
    $fields.Add([pscustomobject]@{
        Id=$id; Block=$block; Address=('Config!' + $cell.Address); Role=$role
        MergeRange=$mergeRange; MergeAnchor=$mergeAnchor
        SavedValue=$cell.Value; Formula=$cell.Formula; Type=$type; UserUnits=$units
        InternalUnits='PendingPerFieldReview'; Validation=$validation
        NewWorkbookDefault=$default; RuntimeDefault='PendingCallSiteReview'
        BlankAndErrorContract='PendingBoundaryTests'; Activity=$activity; Consumers=$consumer
        ExpectedEffect='PendingDirectedBehaviorTest'; TestId='NotRun'; Evidence='NotRun'
        HelpLinks=$help; Hidden=($cell.RowHidden -or $cell.ColumnHidden)
        CoverageStatus='Inventoried:BehaviorNotAccepted'
    })
}

# Reads key/value tables using their actual named rectangles, skipping group captions.
foreach ($name in @('rngSystemSettings','rngCircleGeometry')) {
    $rect = $ranges[$name]
    for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
        $key = [string](Get-Cell $row $rect.Left).Value
        if ($key -notmatch '^[A-Za-z][A-Za-z0-9.]+\.') { continue }
        $consumer = 'CSystemSettingsReader -> settings consumers'
        $activity = 'PendingPerKeyActiveInactiveReview'
        if ($name -eq 'rngCircleGeometry') {
            $consumer = 'CSystemSettingsReader -> CSectionTypeRegistry -> CGeometryCircle / CCircleRebarLayoutBuilder'
            $activity = 'Geometry.Source=Generated; Geometry.Type=Circle; extra rows require an active first row'
        }
        Add-Field $key $name $row ($rect.Left+1) (Get-Cell $row ($rect.Left+2)).Value $consumer $activity
    }
}
foreach ($name in @('rngSteelMaterialParameters','rngConcreteMaterialParameters')) {
    $rect = $ranges[$name]
    for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
        $key = [string](Get-Cell $row $rect.Left).Value
        foreach ($side in @('Compression','Tension')) {
            $column = $rect.Left+1
            if ($side -eq 'Tension') { $column++ }
            Add-Field "$key.$side" $name $row $column (Get-Cell $row ($rect.Left+3)).Value `
                'CSystemSettingsReader -> CMaterialModelProvider -> material spec / solver / engineering calculators' `
                'Selected material, value set, sign branch and TwoLine/ThreeLine model'
        }
    }
}
$rect = $ranges['rngUnitSettings']
for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
    $key = (Get-Cell $row $rect.Left).Value
    Add-Field "Units.$key.Input" 'rngUnitSettings' $row ($rect.Left+1) '-' 'CUnitSystem -> input boundaries' 'All input for this physical quantity'
    Add-Field "Units.$key.Output" 'rngUnitSettings' $row ($rect.Left+3) '-' 'CUnitSystem -> writers / export / plot' 'All output for this physical quantity'
}
$rect = $ranges['rngSignConventionSettings']
for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
    Add-Field ('Signs.' + (Get-Cell $row $rect.Left).Value) 'rngSignConventionSettings' $row ($rect.Left+1) '-' `
        'CUnitSystem -> loads and output; internal signs unchanged' 'All input/output for selected sign convention'
}
$rect = $ranges['rngCalculationProfiles']
for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
    $key = [string](Get-Cell $row ($rect.Left+1)).Value
    if ($key -notmatch '^[A-Za-z].*\.') { continue }
    for ($column=$rect.Left+2; $column -le $rect.Left+5; $column++) {
        $profile = (Get-Cell ($rect.Top+1) $column).Value
        Add-Field "$profile.$key" 'rngCalculationProfiles' $row $column '-' `
            'CCalculationProfileCatalog -> batch / material spec / snapshot presentation' 'LC selects this ProfileId; own calculation or visualization role'
    }
}
$rect = $ranges['rngPlotAnnotationSettings']
for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
    foreach ($kind in @('Rebar','Dimension')) {
        $column = $rect.Left+1
        if ($kind -eq 'Dimension') { $column++ }
        Add-Field ('Plot.' + $kind + '.' + (Get-Cell $row $rect.Left).Value) 'rngPlotAnnotationSettings' $row $column `
            (Get-Cell $row ($rect.Left+3)).Value 'CSystemSettingsReader -> CPlotAnnotationLayout / CSectionPlotter' `
            "Enabled annotation type $kind; presentation only"
    }
}

# Composite geometry extends beyond its legacy named rectangle. Explicit blocks
# come from the actual saved headers and the current composite-table generators.
$geometry = @(
    @{Name='rngRoundedRectangleGeometry'; Blocks=@(@(3,3,2,3,'Dimensions'),@(7,10,2,3,'Side'),@(13,16,2,4,'MainRebar'),@(19,22,2,7,'ExtraRebar'))},
    @{Name='rngHollowRectangleGeometry'; Blocks=@(@(3,4,2,2,'General'),@(7,7,1,6,'Dimensions'),@(11,18,2,4,'MainRebar'),@(21,28,2,7,'ExtraRebar'))},
    @{Name='rngRectSetGeometry'; Blocks=@(@(3,4,2,2,'General'),@(8,8,1,4,'Dimensions'),@(11,18,2,6,'MainRebar'),@(21,28,2,7,'ExtraRebar'))}
)
foreach ($shape in $geometry) {
    $rect = $ranges[$shape.Name]
    foreach ($block in $shape.Blocks) {
        for ($r=$block[0]; $r -le $block[1]; $r++) {
            for ($c=$block[2]; $c -le $block[3]; $c++) {
                $row=$rect.Top+$r-1; $column=$rect.Left+$c-1
                $rowLabel = (Get-Cell $row $rect.Left).Value
                $headerRow = $rect.Top+$block[0]-2
                $header = (Get-Cell $headerRow $column).Value
                $id = "$($shape.Name).$($block[4]).$rowLabel.$header.Config!$(Get-CellAddress $row $column)"
                Add-Field $id $shape.Name $row $column 'Length input / count / selector: see physical field' `
                    'CSystemSettingsReader -> geometry / rebar builder -> CSectionModel' `
                    'Generated selected shape; face/side/row contract; inactive geometry does not affect active shape'
            }
        }
    }
}
foreach ($name in @('rngLoadCombinations','rngStabilityDurationLoads')) {
    $rect=$ranges[$name]
    for ($row=$rect.Top+1; $row -le $rect.Bottom; $row++) {
        for ($column=$rect.Left; $column -le $rect.Right; $column++) {
            $slot=$row-$rect.Top
            $header=(Get-Cell $rect.Top $column).Value
            $consumer='CLoadCombinationReader -> batch -> typed results'
            if ($name -eq 'rngStabilityDurationLoads') { $consumer='modWorkbookCalculation.LoadStabilityDurationLoadsFromWorkbook -> CBatchSectionCalculator -> CStabilityCalculator' }
            Add-Field "$name.Slot$slot.$header" $name $row $column 'User load units / ID / path / profile / comment' `
                $consumer 'Requested source LC slot; gaps and IDs preserved'
        }
    }
}
$rect=$ranges['rngSP35Table721']
for ($row=$rect.Top; $row -le $rect.Bottom; $row++) {
    for ($column=$rect.Left; $column -le $rect.Right; $column++) {
        $cell=Get-Cell $row $column
        $number=0.0
        if ([double]::TryParse([string]$cell.Value,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$number)) {
            Add-Field ('SP35.Table721.' + $cell.Address) 'rngSP35Table721' $row $column '-' `
                'modWorkbookCalculation.ReadSP35Table721FromWorkbook -> CBatchSectionCalculator -> CStabilityCalculator table interpolation' 'SP35 table branch; normative values, no arbitrary calibration' 'NormativeTableValue'
        }
    }
}
for ($row=3; $row -lt $ranges['rngSP35Table721'].Top; $row++) {
    if ((Get-Cell $row 35).Value -notin @('Concrete','Steel')) { continue }
    foreach ($column in @(39,40)) {
        $cell=Get-Cell $row $column
        Add-Field ('MaterialControl.' + $cell.Address) 'UnnamedMaterialDiagramControl' $row $column `
            'Strain / MPa' 'No solver consumer; visible independent formula control' 'Corresponding material parameters; checking only' 'ReadOnlyControlValue'
    }
}
$result=[pscustomobject]@{
    Workbook=$census.Workbook; SHA256=$census.SHA256; CensusPath=$CensusPath
    FieldCount=$fields.Count; UniqueAddressCount=$registered.Count
    MergedFollowerCount=@($fields | Where-Object Role -eq 'NotEditable:MergedFollower').Count
    EditableAddressCount=@($fields | Where-Object Role -eq 'UserInput').Count
    BehavioralAcceptance='NotComplete'; Fields=$fields.ToArray()
}
$output=Join-Path $root $OutputPath
$result | ConvertTo-Json -Depth 13 | Set-Content -LiteralPath $output -Encoding UTF8
$fields | Select-Object Id,Block,Address,Role,MergeRange,MergeAnchor,SavedValue,Formula,Type,UserUnits,NewWorkbookDefault,RuntimeDefault,Activity,Consumers,TestId,CoverageStatus |
    Export-Csv -LiteralPath ([IO.Path]::ChangeExtension($output,'.csv')) -NoTypeInformation -Encoding UTF8
$fields | Group-Object Block | Select-Object Name,Count | Format-Table -AutoSize
Write-Output "CONFIG_FIELD_REGISTRY: fields=$($fields.Count); unique addresses=$($registered.Count); behavioral acceptance=NotComplete"
