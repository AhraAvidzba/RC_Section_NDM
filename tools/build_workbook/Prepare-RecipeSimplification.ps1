# Готовит две проверочные книги: исходные рецепты с новым снимком правил и обновлённый кандидат.
param([Parameter(Mandatory=$true)][ValidateSet('BaselineProbe','Candidate')][string]$Stage)
$ErrorActionPreference='Stop'
$directory='docs/regression/ExcelGuardGuide_2026-10-13/RecipeSimplification_2026-10-16'
$target=Join-Path $directory $Stage
New-Item -ItemType Directory -Path $target -Force | Out-Null
$book=Join-Path $target 'RecipeComparison.xlsm'
if(Test-Path -LiteralPath $book){throw 'Private comparison workbook already exists.'}
Copy-Item -LiteralPath (Join-Path $directory 'Baseline.xlsm') -Destination $book
$modules=@('modTestAudit02Snapshots')
if($Stage -eq 'Candidate'){
    $modules+=@('CSectionRecipeCircle','CSectionRecipeRectSet','CSectionRecipeRoundedRectangle',
        'CSectionRecipeHollowRectangle','modSectionRecipeRows','modSectionRecipeApiGuide','modTestAudit02Api')
}
& (Join-Path $PSScriptRoot 'Update-PrivateWorkbookCode.ps1') -WorkbookPath $book -ModuleNames $modules
if($Stage -eq 'Candidate'){
    & (Join-Path $PSScriptRoot 'Update-RecipeAuthorGuide.ps1') -WorkbookPath $book -PreviousLayoutPath (Join-Path $directory 'BaselineGuideLayout.json')
}
