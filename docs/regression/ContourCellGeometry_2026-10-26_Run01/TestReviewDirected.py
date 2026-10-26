"""Read complete primary reports; terminal markers alone do not imply PASS."""
import argparse
import hashlib
import json
import re
from pathlib import Path

parser = argparse.ArgumentParser()
parser.add_argument("reports", nargs="+")
parser.add_argument("--output", required=True)
args = parser.parse_args()
reviews = []
for filename in args.reports:
    path = Path(filename)
    raw = path.read_bytes()
    body = raw.decode("utf-8-sig")
    starts = re.findall(r"^SUITE_STARTED: (.*?);", body, re.M)
    finishes = re.findall(r"^SUITE_FINISHED: (.*?);", body, re.M)
    errors = re.findall(r"^(?:FAIL\b|RUNTIME ERROR\b|WATCHDOG_FAILURE\b).*", body, re.M)
    totals = re.findall(r"^(TOTAL[^\r\n]*?): passed=(\d+); failed=(\d+)", body, re.M)
    watchdog = re.findall(r"^WATCHDOG_COMPLETED: exit=(\d+); source unchanged=(\w+)", body, re.M)
    source = re.search(r"^SOURCE_SHA256: ([A-Fa-f0-9]{64})", body, re.M)
    passed = (bool(starts) and starts == finishes and len(totals) == len(starts)
              and not errors and all(int(f) == 0 for _, _, f in totals)
              and watchdog == [("0", "True")] and "SOURCE_UNCHANGED: True" in body)
    reviews.append(dict(Path=str(path), SHA256=hashlib.sha256(raw).hexdigest().upper(),
                        Bytes=len(raw), FullPrimaryRead=True, Started=starts, Finished=finishes,
                        SourceSHA256=source.group(1) if source else None,
                        Totals=[dict(Name=n, Passed=int(p), Failed=int(f)) for n,p,f in totals],
                        Errors=errors, TerminalWatchdog=watchdog, Passed=passed))
result = dict(Passed=all(r["Passed"] for r in reviews), FullRegressionPass=False, Reports=reviews)
Path(args.output).write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(dict(Passed=result["Passed"], Reports=[dict(Path=r["Path"], Passed=r["Passed"],
     Totals=r["Totals"], Errors=r["Errors"]) for r in reviews]), ensure_ascii=False))
raise SystemExit(0 if result["Passed"] else 1)
