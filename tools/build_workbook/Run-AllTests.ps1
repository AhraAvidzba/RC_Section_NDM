# служебный PowerShell-скрипт поддерживает сборку, проверку или обновление Excel-книги проекта.

param(
    [string]$ReportPath = "docs/regression/Stage01_AllTests_Report.txt"
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "../..")
$fullReportPath = Join-Path $root $ReportPath
$reportDirectory = Split-Path -Parent $fullReportPath
if (-not (Test-Path -LiteralPath $reportDirectory)) {
    New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}

$scripts = @(
    "Validate-Workbook.ps1",
    "Run-GeometryTests.ps1",
    "Run-MaterialTests.ps1",
    "Run-SectionSolverTests.ps1",
    "Run-CapacityTests.ps1",
    "Run-CrackTests.ps1",
    "Run-BatchTests.ps1",
    "Run-WorkbookInterfaceTests.ps1",
    "Run-RegressionBaselineTests.ps1"
)

$lines = New-Object System.Collections.Generic.List[string]
$failed = $false

foreach ($script in $scripts) {
    $path = Join-Path $PSScriptRoot $script
    $lines.Add("===== $script =====")
    try {
        $global:LASTEXITCODE = 0
        $output = & $path 2>&1 | Out-String -Stream
        $scriptExitCode = $global:LASTEXITCODE
        foreach ($line in $output) {
            $text = [string]$line
            $lines.Add($text)
            if ($text -match "failed=([1-9][0-9]*)" -or $text -match "^FAIL:" -or $text -match "RUNTIME ERROR") {
                $failed = $true
            }
        }
        if ($scriptExitCode -ne 0) {
            $failed = $true
            $lines.Add("SCRIPT EXIT CODE: $scriptExitCode")
        }
    }
    catch {
        $failed = $true
        $lines.Add("SCRIPT ERROR: $($_.Exception.Message)")
    }
    $lines.Add("")
}

$lines | Set-Content -LiteralPath $fullReportPath -Encoding UTF8
$lines
Write-Output "All tests report: $fullReportPath"

if ($failed) {
    exit 1
}
