param([string]$OriginalPath, [string]$CompiledPath, [string]$OutputPath)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
[IO.File]::Copy([IO.Path]::GetFullPath($OriginalPath), [IO.Path]::GetFullPath($OutputPath), $false)
$source = $null
$target = $null
try {
    $source = [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($CompiledPath))
    $target = [IO.Compression.ZipFile]::Open([IO.Path]::GetFullPath($OutputPath), [IO.Compression.ZipArchiveMode]::Update)
    $entry = $source.GetEntry('xl/vbaProject.bin')
    if ($null -eq $entry) { throw 'Compiled VBA project is missing' }
    $target.GetEntry('xl/vbaProject.bin').Delete()
    $replacement = $target.CreateEntry('xl/vbaProject.bin', [IO.Compression.CompressionLevel]::Optimal)
    $inputStream = $entry.Open()
    $outputStream = $replacement.Open()
    try { $inputStream.CopyTo($outputStream) }
    finally { $inputStream.Dispose(); $outputStream.Dispose() }
}
finally {
    if ($source) { $source.Dispose() }
    if ($target) { $target.Dispose() }
}
'Only xl/vbaProject.bin updated; worksheets and native objects preserved'
