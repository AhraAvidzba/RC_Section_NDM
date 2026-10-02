# Reads the current VBA selection for a compile-error diagnostic; changes nothing.
$ErrorActionPreference = 'Stop'
$excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$pane = $excel.VBE.ActiveCodePane
$startRow = 0
$startColumn = 0
$endRow = 0
$endColumn = 0
$pane.GetSelection([ref]$startRow, [ref]$startColumn, [ref]$endRow, [ref]$endColumn)
Write-Output "VBA_SELECTION: $($pane.CodeModule.Name):$startRow; columns=$startColumn-$endColumn"
Write-Output $pane.CodeModule.Lines($startRow, 1)
