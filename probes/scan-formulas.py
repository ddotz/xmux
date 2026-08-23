#!/usr/bin/env python3
"""Count formula-dense sheets in corpus workbooks (evaluation fixture scouting).

Run as a subprocess so the tolerant-CellStyle patch below applies exactly once;
pasting it twice into one interpreter makes the wrapper close over itself.
"""

from __future__ import annotations

import sys
from pathlib import Path

import openpyxl
from openpyxl.styles.cell_style import CellStyle

if not getattr(CellStyle, "_xmux_tolerant", False):
    _attrs = set(getattr(CellStyle, "__attrs__", ()))
    _original = CellStyle.__init__

    def _tolerant(self, **kwargs):
        return _original(self, **{k: v for k, v in kwargs.items() if k in _attrs})

    CellStyle.__init__ = _tolerant
    CellStyle._xmux_tolerant = True


def main(paths: list[str]) -> None:
    for path in paths:
        try:
            book = openpyxl.load_workbook(path, data_only=False)
        except Exception as error:
            print(f"ERR {Path(path).name[:50]} {error}")
            continue
        info = []
        for ws in book.worksheets[:8]:
            count = 0
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        count += 1
            info.append(f"{ws.title[:16]}:{count}")
        print(f"{Path(path).name[:48]} -> {' '.join(info)}")
        book.close()


if __name__ == "__main__":
    main(sys.argv[1:])
