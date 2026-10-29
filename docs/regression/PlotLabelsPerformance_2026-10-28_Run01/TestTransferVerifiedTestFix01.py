"""ДЛЯ ТЕСТОВ: переносит три ограниченные тестовые правки после выхода всех процессов."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib, json, subprocess
e=Path(__file__).resolve().parent;root=e.parents[2]
l=root/'docs/regression/FullSuiteLaunch_2026-10-28_PlotLabels_Run01'
def read(p):return json.loads(p.read_text('utf-8-sig'))
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest().upper()
req=read(l/'LaunchRequest.json');closed=read(l/'AllOwnClosuresReviewed.json')
assert closed['AllOwnIdentitiesTerminal'] and closed['PrimaryLogsReviewed']
assert (closed['CoreMacrosReviewed'],closed['FocusedSingleCasesReviewed'],closed['FocusedRecordCount'])==(182,122,26)
assert closed['SourceSHA256']==req['SourceSHA256'] and closed['SourceCommit']==req['SourceCommit']
# Повторный аудит проверяет текущие PID, время запуска и неизменность всех 504 входов.
closure_script=str(l/'Test-VerifyAllOwnClosures.ps1').replace("'","''")
closure_command="$ErrorActionPreference='Stop'; try { Import-Module 'C:/Windows/System32/WindowsPowerShell/v1.0/Modules/Microsoft.PowerShell.Utility/Microsoft.PowerShell.Utility.psd1' -Global -ErrorAction Stop; & '"+closure_script+"'; exit 0 } catch { Write-Output $_; exit 1 }"
subprocess.run(['C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe','-NoProfile','-ExecutionPolicy','Bypass','-Command',closure_command],check=True,cwd=root)
manifest=read(e/'ParallelFix04/DeferredSourceManifest.json')
source=root/manifest['Source'];candidate=root/manifest['Candidate']
assert sha(source)==manifest['OriginalSHA256'] and sha(candidate)==manifest['CandidateSHA256']
accept=read(e/'ParallelFix01/NativeDropdownAcceptance01.json')
assert len(accept)==2 and {a['Mode'] for a in accept}=={'No','Yes'}
assert all(a['Passed']==1326 and a['Failed']==0 and a['SaveCloseOpen'] and a['EveryOptionAccepted'] and a['JoinedCommaRejected'] and a['JoinedSemicolonRejected'] and a['UnknownRejected'] for a in accept)
dynamic=read(e/'ParallelFix04/NativeAcceptance01.json')
assert dynamic['Status']=='PASS_ISOLATED_DYNAMIC_ID_AND_HORIZONTAL_PROFILE_TESTS'
assert dynamic['ProductionChanged'] is False and dynamic['RootSourceUnchanged']
review=read(e/'ParallelFix04/NativePrimaryReview01.json')
assert review['AcceptedFailed']==0 and review['PositiveRecords']==5
assert review['WrongIdSourceRejected'] and review['OtherEmptySourceRejected']
assert read(e/'ParallelFix04/OwnClosure01.json')['AllOwnClosed']
assert read(e/'ParallelFix01/OwnClosure01.json')['OwnIdentityTerminal']
timer=read(e/'ParallelFix02/DeferredSourceManifest.json')
timer_source=root/timer['Source'];timer_candidate=root/timer['Candidate']
assert sha(timer_source)==timer['OriginalSHA256'] and sha(timer_candidate)==timer['CandidateSHA256']
assert timer['ProductionChanged'] is False
sections=read(e/'ParallelFix03/DeferredSourceManifest.json')
sections_source=root/sections['Source'];sections_candidate=root/sections['Candidate']
assert sha(sections_source)==sections['OriginalSHA256'] and sha(sections_candidate)==sections['CandidateSHA256']
assert read(e/'ParallelFix03/NativePrimaryReview01.json')['AcceptedPassed']==1240
assert read(e/'ParallelFix03/NativePrimaryReview01.json')['AcceptedFailed']==0
assert read(e/'ParallelFix03/OwnClosure01.json')['AllOwnClosed']
out=e/'ParallelFix01/SourceTransfer01.json'
assert not out.exists(), 'Transfer already recorded; inspect rather than repeat'
source.write_bytes(candidate.read_bytes())
timer_source.write_bytes(timer_candidate.read_bytes())
sections_source.write_bytes(sections_candidate.read_bytes())
assert sha(source)==manifest['CandidateSHA256']
assert sha(timer_source)==timer['CandidateSHA256']
assert sha(sections_source)==sections['CandidateSHA256']
sections_record={'Status':'TEST_CAPTION_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED','AtUTC':datetime.now(timezone.utc).isoformat(),'Source':sections['Source'],'OriginalSHA256':sections['OriginalSHA256'],'CandidateSHA256':sections['CandidateSHA256'],'ProductionChanged':False,'OnlySixExpectedStringsChanged':True,'NativePassed':1240,'NativeFailed':0,'CompileAndFullRepeatStillRequired':True}
(e/'ParallelFix03/SourceTransfer01.json').write_text(json.dumps(sections_record,ensure_ascii=False,indent=2),encoding='utf-8')
timer_record={'Status':'TEST_TIMER_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED','AtUTC':datetime.now(timezone.utc).isoformat(),'Source':timer['Source'],'OriginalSHA256':timer['OriginalSHA256'],'CandidateSHA256':timer['CandidateSHA256'],'ProductionChanged':False,'OnlyElapsedReportChanged':True,'CompileAndFullRepeatStillRequired':True}
(e/'ParallelFix02/SourceTransfer01.json').write_text(json.dumps(timer_record,ensure_ascii=False,indent=2),encoding='utf-8')
out.write_text(json.dumps({'Status':'VERIFIED_TEST_ONLY_PATCH_TRANSFERRED_AFTER_ALL_FULL_OWN_EXITED','AtUTC':datetime.now(timezone.utc).isoformat(),'OriginalSHA256':manifest['OriginalSHA256'],'CandidateSHA256':manifest['CandidateSHA256'],'Source':manifest['Source'],'FullSourceSHA256':req['SourceSHA256'],'ProductionChanged':False,'FirstFullPass':False,'AllOwnClosuresProofSHA256':sha(l/'AllOwnClosuresReviewed.json'),'NextRequired':'FreshBuild03, native source equality, new committed epoch and authorized full repeat'},ensure_ascii=False,indent=2),encoding='utf-8')
combined=read(out)
combined['IncludesPF1AndPF4']=True
combined['SupersededPF1CandidateSHA256']=manifest['PF1CandidateSHA256']
combined['DynamicNativeReviewSHA256']=sha(e/'ParallelFix04/NativePrimaryReview01.json')
out.write_text(json.dumps(combined,ensure_ascii=False,indent=2),encoding='utf-8')
(e/'ParallelFix04/SourceTransfer01.json').write_text(json.dumps(combined,ensure_ascii=False,indent=2),encoding='utf-8')
print('TRANSFERRED_THREE_TEST_SOURCES_PF1_PF4_COMBINED_PF2_PF3; full repeat still required')
