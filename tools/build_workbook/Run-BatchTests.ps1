param(
    [string]$WorkbookPath = "workbook/output/RC_Section_NDM.xlsm"
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullWorkbookPath = Join-Path $root $WorkbookPath

if (-not (Test-Path -LiteralPath $fullWorkbookPath)) {
    throw "Workbook not found: $fullWorkbookPath"
}

function Remove-DuplicatePrintAreaName {
    param([string]$Path)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.GetEntry("xl/workbook.xml")
        if ($entry -eq $null) { return }
        $reader = New-Object IO.StreamReader($entry.Open())
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $newXml = [regex]::Replace($xml, '<definedName name="Print_Area"[^>]*>.*?</definedName>\s*', '')
        if ($newXml -ne $xml) {
            $stream = $entry.Open()
            $stream.SetLength(0)
            $writer = New-Object IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
            $writer.Write($newXml)
            $writer.Close()
        }
    }
    finally {
        if ($zip -ne $null) { $zip.Dispose() }
    }
}

Remove-DuplicatePrintAreaName $fullWorkbookPath

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1

    $workbook = $excel.Workbooks.Open($fullWorkbookPath)
    $result = $excel.Run("'RC_Section_NDM.xlsm'!modTestBatchCalculation.RunBatchCalculationTests")
    Write-Output $result
}
finally {
    if ($workbook -ne $null) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel -ne $null) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
