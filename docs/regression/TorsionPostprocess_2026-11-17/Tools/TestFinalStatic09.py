"""ДЛЯ ТЕСТОВ: точная сверка исходного NM-ядра, имён и изоляции раскладки."""
from pathlib import Path
import subprocess,json,hashlib
from datetime import datetime,timezone
R=Path(r'C:/Users/avidzba/Desktop/программы/В_РАБОТЕ!/RC_Section_NDM')
P=Path(r'C:/Users/avidzba/.codex/visualizations/2026/10/12/01a12ece-8192-7141-89b3-f83f1176184c/TorsionPostprocess_2026-11-17')
BASE='1315d2fd28df06465c318d70af822c3a60ac4bda'
def git(*args):
    return subprocess.check_output(['git','-c','safe.directory='+R.as_posix(),*args],cwd=R)
def norm(raw):return raw.decode('utf-8-sig').replace('\r\n','\n')
files=git('ls-tree','-r','--name-only',BASE,'src/Batch','src/Solver','src/Materials','src/Geometry','src/Section','src/Crack','src/Stability','src/Interfaces').decode().splitlines()
changed=[f for f in files if norm(git('show',BASE+':'+f))!=norm((R/f).read_bytes())]
old=json.loads(git('show',BASE+':tools/build_workbook/WorkbookLayout.json'))
new=json.loads((R/'tools/build_workbook/WorkbookLayout.json').read_text(encoding='utf-8-sig'))
names={x['Name']:x for x in new['Names']}
name_diff=[x['Name'] for x in old['Names'] if names.get(x['Name'])!=x]
newnames=sorted(set(names)-{x['Name'] for x in old['Names']})
old_sheets={x['Name']:x for x in old['Sheets']};new_sheets={x['Name']:x for x in new['Sheets']}
sheet_diff=[s for s in old_sheets if s!='Config' and old_sheets[s]!=new_sheets[s]]
result={'ReviewedUTC':datetime.now(timezone.utc).isoformat(),'Base':BASE,'CoreFilesCompared':len(files),'CoreContentChanged':changed,'OriginalNamesChanged':name_diff,'AddedNames':newnames,'OtherSheetLayoutsChanged':sheet_diff,'MainWorkbookGitBlob':git('hash-object','workbook/output/RC_Section_NDM.xlsm').decode().strip(),'ApprovedWorkbookGitBlob':git('rev-parse',BASE+':workbook/output/RC_Section_NDM.xlsm').decode().strip()}
result['Passed']=len(files)==79 and not changed and not name_diff and not sheet_diff and len(newnames)==5 and result['MainWorkbookGitBlob']==result['ApprovedWorkbookGitBlob']
(P/'FinalStatic09.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(result,ensure_ascii=False))
assert result['Passed']
