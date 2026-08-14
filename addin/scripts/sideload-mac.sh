#!/bin/bash
# Sideload the 땡땡엑셀 manifest into Excel for Mac.
#
# Excel picks up any manifest dropped into its container's wef directory at launch,
# so this copies the manifest there and tells you to restart Excel. The directory
# does not exist until the first sideload.
set -euo pipefail
cd "$(dirname "$0")/.."

WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$WEF"
rm -f "$WEF/xmux.manifest.xml"
cp manifest.xml "$WEF/ddot-excel.manifest.xml"

echo "installed -> $WEF/ddot-excel.manifest.xml"
echo
echo "next:"
echo "  1. pnpm dev                 (serves https://localhost:3927)"
echo "  2. quit and reopen Excel"
echo "  3. 홈 탭 -> 땡땡엑셀"
