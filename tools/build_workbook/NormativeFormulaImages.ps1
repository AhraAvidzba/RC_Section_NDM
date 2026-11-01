# Формирует нормативные рисунки справки из точных фрагментов PDF проекта.
# Математическая запись не перепечатывается; номер и пункт добавляются
# единым малым шрифтом нормативной ссылки внутри итогового изображения.
$script:NormativeFormulaRoot = Join-Path $PSScriptRoot 'NormativeFormulaImages'
$script:NormativeFormulaCatalog = @{}
foreach ($item in (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'NormativeFormulaImages.json') -Raw -Encoding UTF8 | ConvertFrom-Json)) {
    if ($script:NormativeFormulaCatalog.ContainsKey([string]$item.Key)) { throw 'Повторный ключ нормативной формулы.' }
    $script:NormativeFormulaCatalog[[string]$item.Key] = $item
}

# Добавляет формулу с проверенным происхождением. Отсутствующий или изменённый
# нормативный растр останавливает сборку, а не подменяется похожей формулой.
function Add-GuideNormativeFormula {
    param([object]$Sheet, [int]$Row, [string]$Key)
    if (-not $script:NormativeFormulaCatalog.ContainsKey($Key)) { throw "Неизвестная нормативная формула: $Key" }
    $item = $script:NormativeFormulaCatalog[$Key]
    $sourcePath = Join-Path $script:NormativeFormulaRoot ([string]$item.File)
    if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -cne [string]$item.SHA256) { throw "Изменён нормативный рисунок: $Key" }
    if (-not (Initialize-FormulaImageRenderer)) { throw 'Недоступен рендер нормативной формулы.' }
    $source = [System.Drawing.Image]::FromFile($sourcePath)
    $referenceFont = New-FormulaFont 16
    $referenceScriptFont = New-FormulaFont 11
    $measureBitmap = New-Object System.Drawing.Bitmap(10, 10)
    $measure = [System.Drawing.Graphics]::FromImage($measureBitmap)
    $bitmap = $null; $graphics = $null; $brush = $null
    try {
        $reference = Measure-MathText $measure ([string]$item.Reference) $referenceFont $referenceScriptFont
        $sourceWidth = [int]$source.Width * 2 # В PDF основные символы высотой около 16 px; масштаб сохраняет привычные 16 pt на листе.
        $sourceHeight = [int]$source.Height * 2
        $width = [int][Math]::Ceiling(48 + $sourceWidth + 18 + $reference.Width)
        $height = [int][Math]::Ceiling(24 + [Math]::Max($sourceHeight, $reference.Height))
        $bitmap = New-Object System.Drawing.Bitmap($width, $height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $graphics.DrawImage($source, 24, [int](($height-$sourceHeight)/2), $sourceWidth, $sourceHeight)
        $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(30,30,30))
        Draw-MathText $graphics $reference.Runs $referenceFont $referenceScriptFont $brush (24+$sourceWidth+18) ($height/2) | Out-Null
        $path = New-FormulaImagePath ('normative_'+$Key)
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $image = @{Path=$path;WidthPoints=$width/2.0;HeightPoints=$height/2.0}
        return (Add-GuideFormulaPicture $Sheet $Row $image ([string]$item.Formula+'      '+[string]$item.Reference))
    } finally {
        if ($brush) {$brush.Dispose()}; if ($graphics) {$graphics.Dispose()}; if ($bitmap) {$bitmap.Dispose()}
        $measure.Dispose(); $measureBitmap.Dispose(); $source.Dispose(); $referenceFont.Dispose(); $referenceScriptFont.Dispose()
    }
}
