"""ДЛЯ ТЕСТОВ: проверяет явно разрешённую приёмку полных групп по отдельным запускам на одной книге."""
from pathlib import Path
import argparse
import hashlib
import importlib.util
import json
import re
from datetime import datetime

ROOT = Path.cwd().resolve()


def test_read(path):
    """Читает сохранённое доказательство, не исправляя его статус или содержимое."""
    return json.loads(path.read_text(encoding="utf-8-sig"))


def test_sha(path):
    """Вычисляет SHA256 фактических байтов книги либо первичного отчёта."""
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def test_module(name, filename):
    """Подключает действующий строгий проверяющий инструмент без изменения его правил."""
    spec = importlib.util.spec_from_file_location(name, ROOT / "tools/build_workbook" / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_closure(report):
    """Требует доказанный выход собственного Excel после выхода дочернего процесса."""
    proof = test_read(Path(str(report) + ".ProcessClosure.json"))
    assert proof["Passed"] and proof["ChildHasExited"] and proof["ChildExitCode"] == 0, report
    assert proof["RunnerCleanupValid"] and proof["Closure"]["Passed"], report
    assert proof["Closure"]["FinalState"]["State"] == "Absent", report
    return {"Report": str(report.relative_to(ROOT)), "ClosureSHA256": test_sha(Path(str(report) + ".ProcessClosure.json"))}


def test_finalize(core, parent):
    """Проверяет все группы и сохраняет исходный FAIL отдельно от согласованного итога по частям."""
    for path in (core, parent):
        assert path.resolve().is_relative_to(ROOT / "docs/regression"), path
    admission = test_read(parent / "Admission.json")
    assert admission["SplitAcceptanceExplicitlyAuthorized"] is True
    assert admission["UserInstruction"] == "Повтори только часовой план"
    assert admission["RetainedCoreExecutions"] == 202 and admission["RetainedSupplementExecutions"] == 8
    source = admission["SourceSHA256"]
    primary = test_read(core / "Run.json")
    original_sha = test_sha(core / "Run.json")
    summary = test_module("split_core", "Summarize-FullOnOffAcceptance.py").summarize(core)
    # Сохраняется исходный неуспешный Focused. Никакие иные причины отказа основных групп не разрешаются.
    assert summary["status"] == "INCOMPLETE_OR_FAILED"
    assert set(summary["errors"]) == {"Driver is not successfully complete", "Mandatory focused acceptance did not finish successfully"}, summary["errors"]
    assert summary["observedSuites"] == summary["expectedSuites"] == 202
    assert len(summary["jobs"]) == 12 and all(job["status"] == "PASS" for job in summary["jobs"])
    assert summary["sourceSHA256"].upper() == source and primary["SourceSHA256"] == source
    assert not primary["SourceChanged"] and not primary["RunnerChanged"]
    old_focused = ROOT / primary["FocusedAcceptance"]["Directory"]
    assert test_sha(old_focused / "Acceptance.json") == primary["FocusedAcceptance"]["AcceptanceSHA256"]
    assert test_read(old_focused / "Acceptance.json")["Status"] == "INCOMPLETE_OR_FAILED"
    closures = [test_closure(ROOT / job["Report"]) for job in primary["CompletedJobs"]]

    focused = parent / "Focused"
    run = test_read(focused / "Run.json")
    target = test_module("split_focused", "Verify-FocusedAcceptance.py").verify(focused, run)
    assert target["Status"] == "PASS", target["Errors"]
    assert target["IndividualCases"] == 122 and run["SourceSHA256"] == source
    assert len(run["Records"]) == 26 and len({r["ID"] for r in run["Records"]}) == 26
    assert run["ProtectedInputsUnchanged"] is True
    assert datetime.fromisoformat(run["FinishedUTC"]) < datetime.fromisoformat("2026-11-24T10:00:00+03:00")
    closures += [test_closure(ROOT / r["Report"]) for r in run["Records"] if r["Kind"] == "VBA"]

    supplement = test_read(core / "Supplement/Acceptance.json")
    macros = ["modTestBatchCalculation.RunZeroLoadCombinationTests", "modTestTorsionPost.RunTorsionOnlyTests", "modTestWorkbookInterface.RunAudit03ExcelGuardTests", "modTestWorkbookInterface.RunTestCalculationFreshnessWorkflow"]
    assert supplement["Status"] == "PASS" and supplement["SourceSHA256"] == source
    assert supplement["Macros"] == macros and [r["Mode"] for r in supplement["Results"]] == ["No", "Yes"]
    patterns = [r"^TOTAL_ZERO_LOAD_COMBINATIONS: passed=([1-9][0-9]*); failed=0$", r"^TorsionOnly: Passed=([1-9][0-9]*); Failed=0$", r"^TOTAL_AUDIT03_EXCEL_GUARD: passed=([1-9][0-9]*); failed=0$", r"^TOTAL_FRESHNESS_WORKFLOW: passed=([1-9][0-9]*); failed=0$"]
    for item in supplement["Results"]:
        assert item["Passed"] and item["ExitCode"] == 0 and item["SourceUnchanged"]
        assert item["SuitesStarted"] == item["SuitesFinished"] == macros
        report = ROOT / item["Report"]
        assert test_sha(report) == item["ReportSHA256"]
        body = report.read_text(encoding="utf-8-sig")
        assert not re.search(r"(?m)^(FAIL(?::| )|RUNTIME ERROR|SCRIPT ERROR|WATCHDOG_FAILURE|RESTORE ERROR|CLEANUP ERROR)", body)
        assert "SOURCE_SHA256: " + source in body and "WATCHDOG_COMPLETED: exit=0; source unchanged=True" in body
        assert "GLOBAL_MODE: General.DiagramExtension=" + item["Mode"] + ";" in body
        assert "MODE_AFTER_SUITE: " + item["Mode"] in body
        assert all(len(re.findall(pattern, body, re.M)) == 1 for pattern in patterns)
        closures.append(test_closure(report))

    assert test_sha(Path(run["SourceWorkbook"])) == source
    assert test_sha(core / "Source/RC_Section_NDM.xlsm") == source
    assert test_sha(ROOT / "workbook/output/RC_Section_NDM.xlsm") == source
    review = test_read(ROOT / "docs/regression/ResultTailLatency_2026-11-24_Repair03/CurrentWorkbookAndGoalAudit.json")
    assert review["WorkbookSHA256"] == source and review["RuntimeCodeEqualsNativeTimingAndDisplayCandidate"]
    assert test_sha(ROOT / "workbook/output/VBA_All_Code.txt") == review["ExportSHA256"]
    independent = test_read(focused / "Independent.json")
    native = test_read(focused / "NativeCAD/Stage.json")
    assert independent["Passed"] and independent["Unsupported"] == []
    assert independent["SourceSHA256"] == source
    assert native["Passed"] and native["UserDocumentsUsed"] is False
    assert test_sha(core / "Run.json") == original_sha
    return {"Status": "PASS", "AcceptanceMode": "UserAuthorizedSeparateCompleteGroups", "SingleInvocation": False,
            "UserInstruction": admission["UserInstruction"], "AdmissionSHA256": test_sha(parent / "Admission.json"),
            "SourceSHA256": source, "SameWorkbookBytes": True, "CoreSuites": 202, "FocusedSingleCases": 122, "SupplementSuites": 8,
            "CoreDirectory": str(core), "FocusedDirectory": str(focused), "CoreSourceCommit": primary["SourceCommit"], "FocusedSourceCommit": run["SourceCommit"],
            "OriginalRun03Status": summary["status"], "OriginalFocusedStatus": "INCOMPLETE_OR_FAILED", "OriginalRun03SHA256": original_sha, "OriginalFailureEvidencePreserved": True,
            "FocusedAcceptanceSHA256": test_sha(focused / "Acceptance.json"), "IndependentIntegrals": independent["IndependentIntegrals"],
            "ReviewedNumFailStates": independent["ReviewedNumFailStates"], "NativeCADPassed": True, "OwnExcelClosures": closures,
            "WorkingWorkbookUnchanged": True, "FinishedUTC": datetime.now().astimezone().isoformat(), "Meaning": "Все согласованные полные группы проверены на одной книге в отдельных запусках; исходный неуспешный Run03 сохранён."}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--core-directory", type=Path, required=True)
    parser.add_argument("--focused-parent", type=Path, required=True)
    args = parser.parse_args()
    result = test_finalize(args.core_directory.resolve(), args.focused_parent.resolve())
    (args.focused_parent / "AcceptanceByParts.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False))
