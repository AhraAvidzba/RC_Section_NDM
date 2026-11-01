# Вставляет заново набранные нормативные формулы с проверенными математической записью и ссылками.
# Источники PDF и MathText хранятся в каталоге; исходные растры PDF не увеличиваются.
$script:NormativeFormulaRoot = Join-Path $PSScriptRoot 'NormativeFormulaImages'
$script:NormativeFormulaCatalog = @{}
foreach ($item in (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'NormativeFormulaImages.json') -Raw -Encoding UTF8 | ConvertFrom-Json)) {
    if ($script:NormativeFormulaCatalog.ContainsKey([string]$item.Key)) { throw 'Повторный ключ нормативной формулы.' }
    $script:NormativeFormulaCatalog[[string]$item.Key] = $item
}

# Вставляет PNG с четырьмя пикселями на пункт; несовпадение SHA останавливает сборку.
function Add-GuideNormativeFormula {
    param([object]$Sheet, [int]$Row, [string]$Key)
    if (-not $script:NormativeFormulaCatalog.ContainsKey($Key)) { throw "Неизвестная нормативная формула: $Key" }
    $item = $script:NormativeFormulaCatalog[$Key]
    $sha=[Security.Cryptography.SHA256]::Create()
    try {$expressionHash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes([string]$item.MathText+"`n"+[string]$item.Reference))).Replace('-','')} finally {$sha.Dispose()}
    if($expressionHash -cne [string]$item.ExpressionSHA256){throw "Математическая запись изменена: повторите Render-NormativeFormulas.py для $Key"}
    $sourcePath = Join-Path $script:NormativeFormulaRoot ([string]$item.File)
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -cne [string]$item.SHA256) { throw "Изменён нормативный рисунок: $Key" }
    if ([double]$item.PixelsPerPoint -lt 4 -or [double]$item.WidthPoints -le 0 -or [double]$item.HeightPoints -le 0) { throw "Некорректные размеры формулы: $Key" }
    [void][IO.Directory]::CreateDirectory($script:FormulaImageDir)
    $path = New-FormulaImagePath ('normative_'+$Key)
    Copy-Item -LiteralPath $sourcePath -Destination $path # Уникальный путь исключает кэш Excel от предыдущей вставки.
    $image = @{Path=$path;WidthPoints=[double]$item.WidthPoints;HeightPoints=[double]$item.HeightPoints}
    return (Add-GuideFormulaPicture $Sheet $Row $image ([string]$item.Formula+'      '+[string]$item.Reference))
}
