param([string]$OriginalPath, [string]$CandidatePath, [string]$ReportPath)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archives = @()
$lines = New-Object System.Collections.Generic.List[string]
try {
    $archives += [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($OriginalPath))
    $archives += [IO.Compression.ZipFile]::OpenRead([IO.Path]::GetFullPath($CandidatePath))
    $patterns = '^xl/(worksheets/|drawings/|charts/|media/|ctrlProps/|activeX/|embeddings/|printerSettings/|sharedStrings\.xml$|styles\.xml$|theme/)'
    $originalEntries = @($archives[0].Entries | Where-Object { $_.FullName -match $patterns })
    $candidateEntries = @($archives[1].Entries | Where-Object { $_.FullName -match $patterns })
    if ($originalEntries.Count -ne $candidateEntries.Count) { throw 'Worksheet/object entry count changed' }
    foreach ($entry in $originalEntries) {
        $other = $archives[1].GetEntry($entry.FullName)
        if ($null -eq $other) { throw "Missing entry: $($entry.FullName)" }
        $hashes = @()
        foreach ($item in @($entry, $other)) {
            $stream = $item.Open()
            $sha = [Security.Cryptography.SHA256]::Create()
            try { $hashes += [BitConverter]::ToString($sha.ComputeHash($stream)) }
            finally { $stream.Dispose(); $sha.Dispose() }
        }
        if ($hashes[0] -ne $hashes[1]) { throw "Changed worksheet/object entry: $($entry.FullName)" }
        $lines.Add('PASS: unchanged ' + $entry.FullName)
    }
    $books = @()
    foreach ($archive in $archives) {
        $reader = New-Object IO.StreamReader($archive.GetEntry('xl/workbook.xml').Open())
        try { $books += [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    foreach ($node in @('sheets', 'definedNames')) {
        if ($books[0].workbook.$node.OuterXml -cne $books[1].workbook.$node.OuterXml) { throw "Changed workbook $node" }
        $lines.Add('PASS: unchanged ' + $node)
    }
    $lines.Add('PASS: no calculation or user-sheet changes')
}
finally {
    foreach ($archive in $archives) { $archive.Dispose() }
    [IO.File]::WriteAllLines([IO.Path]::GetFullPath($ReportPath), $lines, (New-Object Text.UTF8Encoding($true)))
}
'PRESERVATION: ' + $lines.Count + ' checks passed'
