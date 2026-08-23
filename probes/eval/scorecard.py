#!/usr/bin/env python3
"""Scorecard for xmux harness evaluation runs.

Reads one pilot run file (probes/eval/runs/pilot-<timestamp>.jsonl, written by the
eval suite) and prints the quantitative report the improvement loop acts on:

- goal attainment per case (share of repetitions whose checks all passed)
- reproducibility per case: the recorded numeric_reproducibility check when the
  runner logged one, else identical final numeric content across repetitions
- false-refusal rate
- efficiency: median wall-clock per case

Usage:
    python3 probes/eval/scorecard.py                # newest pilot-*.jsonl
    python3 probes/eval/scorecard.py <file.jsonl>   # one specific run file

Historical runs stay in their own timestamped files precisely so a scorecard never
mixes a crashed cycle into a clean one.
"""

from __future__ import annotations

import json
import statistics
import sys
from collections import defaultdict
from pathlib import Path

REFUSAL_MARK = "주장과 일치시키지 못했습니다"
REPRO_CHECK = "numeric_reproducibility"


def numbers(text: str) -> set[float]:
    out: set[float] = set()
    token = ""
    for ch in text:
        if ch.isdigit() or (ch in ",." and token):
            token += ch
        elif token:
            try:
                out.add(float(token.replace(",", "")))
            except ValueError:
                pass
            token = ""
    if token:
        try:
            out.add(float(token.replace(",", "")))
        except ValueError:
            pass
    return out


def newest_run(runs_dir: Path) -> Path | None:
    files = sorted(runs_dir.glob("pilot-*.jsonl"), key=lambda p: p.stat().st_mtime)
    return files[-1] if files else None


def main(run_file: Path) -> None:
    cases: dict[str, list[dict]] = defaultdict(list)
    for line in run_file.read_text().splitlines():
        if not line.strip():
            continue
        record = json.loads(line)
        cases[record["case"]].append(record)

    if not cases:
        print(f"no evaluation records found in {run_file.name}")
        return

    print(f"run: {run_file.name}")
    print(f"{'case':6} {'reps':>4} {'goal':>7} {'reprod':>7} {'refusal':>8} {'med ms':>9}")
    totals: list[float] = []
    for case in sorted(cases):
        reps = cases[case]
        goal = statistics.mean(
            1.0 if all(check["pass"] for check in rep.get("checks", [])) else 0.0
            for rep in reps
        )
        recorded_repro = [
            check["pass"]
            for rep in reps
            for check in rep.get("checks", [])
            if check["name"] == REPRO_CHECK
        ]
        if recorded_repro:
            reproducible = statistics.mean(1.0 if ok else 0.0 for ok in recorded_repro)
        elif len(reps) >= 2:
            answers = [frozenset(numbers(rep.get("answer", ""))) for rep in reps]
            reproducible = 1.0 if all(a == answers[0] for a in answers[1:]) else 0.0
        else:
            reproducible = float("nan")
        refusals = statistics.mean(
            1.0 if REFUSAL_MARK in rep.get("answer", "") else 0.0 for rep in reps
        )
        median_ms = statistics.median(rep.get("ms", 0) for rep in reps)
        print(
            f"{case:6} {len(reps):>4} {goal:>6.0%} {reproducible:>7.0%} "
            f"{refusals:>8.0%} {median_ms:>9.0f}"
        )
        totals.append(goal)

    print(f"\noverall goal attainment: {statistics.mean(totals):.0%}")


if __name__ == "__main__":
    runs_dir = Path(__file__).parent / "runs"
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else newest_run(runs_dir)
    if target is None or not target.exists():
        print("no run file given and none found")
        raise SystemExit(1)
    main(target)
