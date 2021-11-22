#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
Load model data from sheet

Pass in master_file return data structure
"""
from dataclasses import dataclass
from dataclasses_json import dataclass_json
from pathlib import Path
from openpyxl import load_workbook # pylint: disable=import-error
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from .utilities import status_msg

@dataclass_json
@dataclass
class MarkUp:
    """Mark-up rates by deprtment"""
    policy: str
    markup_1: float
    markup_2: float
    discount: float

def load_mark_ups(xlsx_file: Path) -> dict[str, MarkUp]:
    """read makrkup file into object """
    status_msg('Loading Mark Ups', 1)
    status_msg(f'  {xlsx_file.name}', 2)
    xlsx: Workbook = load_workbook(xlsx_file.as_posix(), data_only=True)
    sheet: Worksheet = xlsx.active
    mark_ups: dict[str, MarkUp] = {}
    for row in sheet.iter_rows(min_row=2, max_col=4):
        if not isinstance(row[0].value, str):
            continue
        mark_up: MarkUp = MarkUp(
            row[0].value,
            float(row[1].value),
            float(row[2].value),
            float(row[3].value))
        status_msg(f"    {mark_up}", 3)
        mark_ups[row[0].value] = mark_up
    xlsx.close()
    return mark_ups

if __name__ == "__main__":
    pass
