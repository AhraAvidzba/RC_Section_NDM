# Общая поддержка импортируемых примеров; только явно указанная isolated fixture.
function Get-GeometryExampleFiles([string]$Root, [string[]]$Paths) {
    $exampleRoot=[IO.Path]::GetFullPath((Join-Path $Root 'docs/geometry_refactor/examples'))+[IO.Path]::DirectorySeparatorChar
    foreach ($relative in $Paths) {
        $file=Get-Item -LiteralPath (Join-Path $Root $relative)
        if (-not $file.FullName.StartsWith($exampleRoot,[StringComparison]::OrdinalIgnoreCase) -or $file.Extension -notin '.cls','.bas') {throw 'Example source must be under docs/geometry_refactor/examples.'}
        $file
    }
}

function Add-GeometryExampleRegistration([string]$Body, [string[]]$Registration) {
    if (-not $Registration.Count) {return $Body}
    if ($Registration.Count -ne 3) {throw 'Registration requires TypeID, class name and one root range.'}
    foreach ($value in $Registration) {if ($value -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {throw 'Invalid example registration token.'}}
    $pattern='(?ms)(^Private Sub Class_Initialize\(\).*?)(\r?\nEnd Sub)'
    if ([regex]::Matches($Body,$pattern).Count -ne 1) {throw 'Registry initializer is not unique.'}
    $line='    RegisterRecipe "'+$Registration[0]+'", New '+$Registration[1]+', Array("'+$Registration[2]+'")'
    return [regex]::Replace($Body,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $match.Groups[1].Value+"`r`n"+$line+$match.Groups[2].Value})
}
