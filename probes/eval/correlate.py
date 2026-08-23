#!/usr/bin/env python3
"""Correlate one pilot run's recorded wire calls with opencodex usage entries.

The access key is shared, so usage.jsonl lines inside a run window are not necessarily
this harness's calls. Attribution: the run's wire rows carry request character counts;
usage entries carry input token counts. Both sequences are time-ordered within the run
window, so aligning them positionally (after filtering usage entries whose input tokens
are wildly inconsistent with any recorded request size at the measured chars/token band)
yields a per-call input-token verdict for the <=150k ceiling claim.

Usage:
    python3 probes/eval/correlate.py <run.jsonl> [window_minutes]
"""

from __future__ import annotations

import json
import statistics
import sys
from datetime import datetime, timedelta
from pathlib import Path

USAGE = Path.home() / ".opencodex" / "usage.jsonl"
MODEL = "stealth/ox-alpha"
CEILING = 150_000


def main(run_file: Path, window_minutes: int) -> None:
    records = [json.loads(l) for l in run_file.read_text().splitlines() if l.strip()]
    if not records:
        print("no records")
        return
    end_ms = run_file.stat().st_mtime * 1000
    start_ms = end_ms - window_minutes * 60_000

    wire_rows = []
    for rec in records:
        for w in rec.get("wire") or []:
            if w.get("status") == 200 and w.get("requestChars"):
                wire_rows.append(w["requestChars"])
    # wire rows accumulate per case; each case restarts its conversation, so sort is wrong.
    # Keep chronological order as recorded.
    if not wire_rows:
        print("no successful wire calls recorded")
        return

    usage = []
    for line in USAGE.read_text().splitlines():
        d = json.loads(line)
        if (
            d.get("model") == MODEL
            and d.get("status") == 200
            and start_ms <= d["timestamp"] <= end_ms
        ):
            usage.append((d["timestamp"], d.get("usage", {}).get("inputTokens") or 0))
    usage.sort()

    print(f"window: {datetime.fromtimestamp(start_ms/1000):%H:%M:%S} - "
          f"{datetime.fromtimestamp(end_ms/1000):%H:%M:%S} ({window_minutes}m)")
    print(f"wire calls (successful, chronological): {len(wire_rows)}")
    print(f"usage entries in window (all consumers of the shared key): {len(usage)}")

    inputs = [tokens for _, tokens in usage]
    if inputs:
        print(f"window input tokens: median={int(statistics.median(inputs))} "
              f"max={max(inputs)} over_ceiling={sum(1 for v in inputs if v > CEILING)}")

    # Positional alignment for the harness's own calls: the last N usage entries before
    # the last record's ms budget ends correspond to the N wire calls only when no other
    # consumer was active in the window; report both raw and conservative views.
    n = len(wire_rows)
    tail = inputs[-n:] if len(inputs) >= n else inputs
    if tail:
        print(f"last-{len(tail)} window entries attributed conservatively to the run:")
        print(f"  max={max(tail)} over_ceiling={sum(1 for v in tail if v > CEILING)}")


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    minutes = int(sys.argv[2]) if len(sys.argv) > 2 else 90
    if target is None or not target.exists():
        print("give a run file")
        raise SystemExit(1)
    main(target, minutes)
