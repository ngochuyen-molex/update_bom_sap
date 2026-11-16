"""Runs the whole chain on CHAIN_FILE: create BOM -> assign components -> production version.

Usage:
    python run_all.py [check] [slow] [from=2|3]

    (no option)  phase 1, 2, 3 in SAP, one after the other
    check        all phases without SAP: build + validate the work lists, write the sheets
    slow         wait after every SAP action (passed on to every phase)
    from=2       start at phase 2 (or from=3), e.g. after fixing a problem in phase 2

Each phase is the normal script, started exactly as you would start it by hand:
    phase 1  python bom_create_main.py              (always continues, SUCCESS BOMs never created again)
    phase 2  python routing_main.py fresh           (rebuilds from Sheet1, keeps SUCCESS rows)
    phase 3  python production_version.py fresh     (rebuilds from Routing Result, keeps SUCCESS rows)

After each phase the result sheet is counted. The chain stops when the next phase has
nothing to work with (no SUCCESS rows) or the result sheet is missing.
FAILED rows do not stop the chain: the next phase skips them, fix them and run again.
"""
import subprocess
import sys
from collections import Counter
from config import CHAIN_FILE
from workbook import (ROUTING_SHEET, PROD_VERSION_SHEET, STATUS_COL,
                      has_sheet, read_sheet)

PHASES = [
    # (number, name, script, options, result sheet: 0 = first sheet)
    (1, "Create BOM", "bom_create_main.py", [], 0),
    (2, "Assign components", "routing_main.py", ["fresh"], ROUTING_SHEET),
    (3, "Production version", "production_version.py", ["fresh"], PROD_VERSION_SHEET),
]


def status_counts(sheet):
    # STATUS counts of a result sheet, None if the sheet is not there
    if sheet != 0 and not has_sheet(CHAIN_FILE, sheet):
        return None
    df = read_sheet(CHAIN_FILE, sheet)
    if STATUS_COL not in df.columns:
        return None
    return Counter(s for s in df[STATUS_COL] if s)


def run_phase(script, options):
    # Same Python, same folder; the phase can still ask [s]/[r]/[q] when the file is open
    command = [sys.executable, script, *options]
    print(f"\n> {' '.join(command[1:])}")
    return subprocess.run(command).returncode


def main():
    args = sys.argv[1:]
    extra = [a for a in ("check", "slow") if a in args]
    start = next((int(a.split("=")[1]) for a in args if a.startswith("from=")), 1)
    print(f"Run all, file: {CHAIN_FILE}, options: {extra or 'none'}, from phase {start}")

    summary = []
    for number, name, script, options, sheet in PHASES:
        if number < start:
            continue
        print("\n" + "#" * 80 + f"\n# Phase {number} - {name}\n" + "#" * 80)
        code = run_phase(script, options + extra)
        counts = status_counts(sheet)
        summary.append((number, name, counts))
        print(f"\nPhase {number} result: {dict(counts) if counts else 'no result sheet'}")

        if code != 0:
            print(f"STOP: {script} ended with an error (exit code {code}).")
            break
        if counts is None:
            print(f"STOP: no result sheet after phase {number}. See the messages above.")
            break
        if number < len(PHASES) and not counts.get("SUCCESS") and "check" not in extra:
            print(f"STOP: phase {number} has no SUCCESS rows, phase {number + 1} would have nothing to do.")
            break

    print("\n" + "=" * 80 + "\nSummary (rows per STATUS; phase 1 counts component rows)")
    for number, name, counts in summary:
        print(f"  Phase {number} {name:<20} {dict(counts) if counts else 'no result sheet'}")
    print(f"Details: sheets of {CHAIN_FILE}")


if __name__ == "__main__":
    main()
