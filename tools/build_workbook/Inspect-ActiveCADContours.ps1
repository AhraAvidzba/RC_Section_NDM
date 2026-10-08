# Читает фактические объекты контурных слоев активного AutoCAD без изменений.
# Не копирует/взрывает объекты, не переключает окна и не закрывает документы.
param([string]$ReportPath = 'docs/regression/Performance/PolylineImport_2026-10-08/ActiveCAD.json')
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$report = [IO.Path]::GetFullPath((Join-Path $root $ReportPath))
if (-not $report.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Report must remain inside the project.' }
$acad = $null
foreach ($id in @('AutoCAD.Application.24.2', 'AutoCAD.Application')) {
    try { $acad = [Runtime.InteropServices.Marshal]::GetActiveObject($id); break } catch {}
}
if (-not $acad) { throw 'No running AutoCAD instance found.' }
$doc = $acad.ActiveDocument
$rows = @()
foreach ($entity in $doc.ModelSpace) {
    $layer = [string]$entity.Layer
    if ($layer -notin 'RC_NDM_Contour', 'RC_NDM_Opening') { continue }
    $kind = [string]$entity.ObjectName
    $entry = [ordered]@{Handle = [string]$entity.Handle; Layer = $layer; Kind = $kind; LayerLocked = [bool]$doc.Layers.Item($layer).Lock}
    if ($kind -in 'AcDbPolyline', 'AcDb2dPolyline') {
        $entry.Closed = [bool]$entity.Closed; $entry.Elevation = [double]$entity.Elevation
        $entry.Coordinates = @($entity.Coordinates); $entry.Normal = @($entity.Normal)
        $stride = 2; if ($kind -eq 'AcDb2dPolyline') { $stride = 3 }
        $bulges = @(); $count = $entry.Coordinates.Count / $stride
        for ($i = 0; $i -lt $count; $i++) { $bulges += [double]$entity.GetBulge($i) }
        $entry.Bulges = $bulges
    } elseif ($kind -eq 'AcDbLine') {
        $entry.StartPoint = @($entity.StartPoint); $entry.EndPoint = @($entity.EndPoint)
    } elseif ($kind -eq 'AcDbCircle') {
        $entry.Center = @($entity.Center); $entry.Radius = [double]$entity.Radius
    }
    $rows += $entry
}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($report))
[ordered]@{ReadOnly = $true; Application = [string]$acad.FullName; Document = [string]$doc.FullName;
    ModelSpaceCount = [int]$doc.ModelSpace.Count; Entities = $rows} |
    ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $report -Encoding UTF8
Write-Output "READ_ONLY_CAD: contours=$($rows.Count); report=$report"
