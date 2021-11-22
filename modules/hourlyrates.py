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
class HourlyRate:
    """Hourly rate by department"""
    dept: str
    rate: float


def load_hourly_rates(xlsx_file: Path) -> dict[str, HourlyRate]:
    """Read consuables sheet"""
    status_msg('Loading Consumables', 1)
    status_msg(f'  {xlsx_file.name}', 2)
    xlsx: Workbook = load_workbook(xlsx_file.as_posix(), data_only=True)
    sheet: Worksheet = xlsx.active
    hourly_rates: dict[str, HourlyRate] = {}
    for row in sheet.iter_rows(min_row=2, max_col=2):
        if not isinstance(row[0].value, str):
            continue
        hourly_rate: HourlyRate = HourlyRate(
            row[0].value,
            float(row[1].value))
        status_msg(f"    {hourly_rate}", 3)
        hourly_rates[row[0].value] = hourly_rate
    xlsx.close()
    return hourly_rates

if __name__ == "__main__":
    pass
