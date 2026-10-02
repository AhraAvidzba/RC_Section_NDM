# Собирает результаты выполненных нагрузочных тестов в проверяемый CSV-реестр.
# Читает только реальные логи: статусы/коды не назначает, численные эталоны
# не пересчитывает. Неуспешные и незавершенные прогоны явно остаются в manifest.
param(
    [string]$LogDirectory = 'docs/regression/Audit03',
    [string]$OutputPrefix = 'docs/regression/Audit03/load_matrix_summary_2026-10-02'
)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$directory = (Resolve-Path -LiteralPath (Join-Path $root $LogDirectory)).Path
$rows = [Collections.Generic.List[object]]::new()
$runs = [Collections.Generic.List[object]]::new()

# Разбирает явно именованные поля строки протокола. Значение после первого
# равенства сохраняется целиком; разделитель полей определен самим VBA-тестом.
function Read-ProtocolFields([string]$Line) {
    $fields = @{}
    foreach ($part in $Line.Split('|')) {
        $index = $part.IndexOf('=')
        if ($index -gt 0) { $fields[$part.Substring(0, $index)] = $part.Substring($index + 1) }
    }
    return $fields
}

foreach ($file in Get-ChildItem -LiteralPath $directory -File -Filter 'broad_matrix_*_v*.txt' |
        Where-Object { $_.Name -notmatch '\.progress\.txt$' } | Sort-Object Name) {
    $lines = [IO.File]::ReadAllLines($file.FullName, [Text.Encoding]::UTF8)
    $sourceHash = ''; $mode = ''; $current = $null; $count = 0
    $total = ''; $reopen = $false; $finished = $false; $failed = -1
    foreach ($line in $lines) {
        if ($line.StartsWith('SOURCE_SHA256: ')) { $sourceHash = $line.Substring(15) }
        if ($line -match '^GLOBAL_MODE: General\.DiagramExtension=(Yes|No);') { $mode = $Matches[1] }
        if ($line.StartsWith('MATRIX_CASE: ')) {
            $f = Read-ProtocolFields $line.Substring(13)
            $current = [ordered]@{
                Log = $file.Name; SourceSHA256 = $sourceHash; ExtensionEnabled = $mode
                Shape = $f['shape']; Family = $f['family']; CombinationID = $f['id']; Load = $f['load']
                N_Internal = $f['N']; Mx_Internal = $f['Mx']; My_Internal = $f['My']
                Profile = $f['profile']; RequestedPath = $f['path']
                CapacityScaledComponent = ''; LambdaCapacity = ''; CapacityMethod = ''; CapacityLimit = ''; CapacityExternalStatus = ''
                FormationPath = ''; LambdaCrc = ''; Ncrc_Internal = ''; Mcrc_Internal = ''; HasFormationPoint = ''; CrackFormed = ''
                CrackDataLogged = $false; WidthCalculated = ''; SigmaS_MPa = ''; SigmaSCrc_MPa = ''; PsiS = ''
                Abt_mm2 = ''; As_mm2 = ''; Ds_mm = ''; Ls_mm = ''; Acrc_mm = ''; SelectedRebars = ''
            }
            foreach ($kind in @('direct', 'capacity', 'formation', 'current', 'width', 'longitudinal', 'stability')) {
                foreach ($field in @('InternalStatus', 'ResultCode', 'Applies', 'Calculated', 'ResultComment')) { $current[$kind + '_' + $field] = '' }
            }
            foreach ($kind in @('Strength', 'Crack', 'Stability', 'Batch')) { $current[$kind + '_OutputComment'] = '' }
            $rows.Add($current); $count++
            continue
        }
        if ($null -ne $current -and $line.StartsWith('CASE: audit03.matrix.')) {
            $f = Read-ProtocolFields $line
            if ($f.ContainsKey('requested')) {
                $current.FormationPath = $f['resolved']; $current.LambdaCrc = $f['lambda']
                $current.Ncrc_Internal = $f['Ncrc']; $current.Mcrc_Internal = $f['Mcrc']
                $current.HasFormationPoint = $f['point']; $current.CrackFormed = $f['formed']
            } else {
                $current.CapacityScaledComponent = $f['path']; $current.LambdaCapacity = $f['lambda']
                $current.CapacityMethod = $f['method']; $current.CapacityLimit = $f['limit']
                $current.CapacityExternalStatus = $f['status']
            }
            continue
        }
        if ($null -ne $current -and $line -match '^META: audit03\.comments\.[^|]+\.(direct|capacity|formation|current|width|longitudinal|stability)\|') {
            $kind = $Matches[1]; $f = Read-ProtocolFields $line
            foreach ($pair in @(@('InternalStatus','status'), @('ResultCode','code'), @('Applies','applies'), @('Calculated','calculated'), @('ResultComment','comment'))) {
                $current[$kind + '_' + $pair[0]] = $f[$pair[1]]
            }
            continue
        }
        if ($null -ne $current -and $line.StartsWith('OUTPUT: audit03.comments.')) {
            $f = Read-ProtocolFields $line
            foreach ($kind in @('Strength', 'Crack', 'Stability', 'Batch')) { $current[$kind + '_OutputComment'] = $f[$kind.ToLowerInvariant()] }
        }
        if ($null -ne $current -and $line.StartsWith('MATRIX_CRACK_DATA: ')) {
            $f = Read-ProtocolFields $line
            $current.CrackDataLogged = $true
            foreach ($pair in @(@('WidthCalculated','calculated'), @('SigmaS_MPa','sigmaS'),
                    @('SigmaSCrc_MPa','sigmaSCrc'), @('PsiS','psi'), @('Abt_mm2','Abt'),
                    @('As_mm2','As'), @('Ds_mm','ds'), @('Ls_mm','ls'), @('Acrc_mm','acrc'),
                    @('SelectedRebars','bars'))) {
                $current[$pair[0]] = $f[$pair[1]]
            }
        }
        if ($line.StartsWith('TOTAL_AUDIT03_LOAD_MATRIX: ')) {
            $total = $line
            if ($line -match '; failed=(\d+)$') { $failed = [int]$Matches[1] }
        }
        if ($line -match '^RESULTS_SAVE_REOPEN: .*; equal=True;') { $reopen = $true }
        if ($line -eq 'WATCHDOG_COMPLETED: exit=0; source unchanged=True') { $finished = $true }
    }
    $runs.Add([pscustomobject]@{
        Log = $file.Name; SourceSHA256 = $sourceHash; ExtensionEnabled = $mode
        ParsedCases = $count; FailedAssertions = $failed; ResultsReopenEqual = $reopen
        CompletedSuccessfully = $finished; Acceptance = $(if ($finished -and $reopen -and $failed -eq 0) { 'PASS' } else { 'NotPassed' })
        Total = $total
    })
}
$prefix = [IO.Path]::GetFullPath((Join-Path $root $OutputPrefix))
$rows | ForEach-Object { [pscustomobject]$_ } | Export-Csv -LiteralPath ($prefix + '_cases.csv') -NoTypeInformation -Encoding UTF8
$runs | Export-Csv -LiteralPath ($prefix + '_runs.csv') -NoTypeInformation -Encoding UTF8
Write-Output "LOAD_MATRIX_SUMMARY: runs=$($runs.Count); cases=$($rows.Count); passedRuns=$(@($runs | Where-Object Acceptance -eq 'PASS').Count)"
