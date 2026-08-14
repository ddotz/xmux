#!/bin/bash
# Drives Excel through the exact interaction xmux must support, while ax_probe
# records the AX state. Answers:
#   P1  Is the in-progress formula text readable from AX during edit mode?
#   P2  Does Excel expose a caret/selection range inside the editor?
#   P3  What does the status bar say in each mode (Ready / Edit / Point)?
#   P4  What does Tab actually do inside a formula edit?
#
# Requires: Accessibility permission for the invoking terminal, and
# /tmp/xmux_probe.xlsx open in Excel.
set -uo pipefail
cd "$(dirname "$0")/.."

OUT=${1:-/tmp/xmux_state.log}
SECONDS_TOTAL=${2:-26}

./probes/ax_probe state --hz 25 --seconds "$SECONDS_TOTAL" >"$OUT" 2>&1 &
PROBE=$!
sleep 1.5

osascript <<'APPLESCRIPT'
tell application "Microsoft Excel"
    activate
    select range "B2" of worksheet "Main" of active workbook
end tell
delay 1.5
tell application "System Events"
    tell process "Microsoft Excel"
        -- F2: enter cell edit mode
        key code 120
        delay 2.0
        -- Tab twice: what does Excel do inside a formula edit?
        key code 48
        delay 2.0
        key code 48
        delay 2.0
        -- Escape: abandon the edit, leave the sheet untouched
        key code 53
    end tell
end tell
APPLESCRIPT

wait $PROBE
echo "=== $OUT ==="
cat "$OUT"
