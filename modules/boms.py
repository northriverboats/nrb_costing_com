#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
Load model data from sheet

Pass in master_file return data structure
"""
from dataclasses import dataclass, field
from dataclasses_json import dataclass_json
from datetime import datetime
from pathlib import Path
from typing import Optional, Union
from openpyxl import load_workbook # pylint: disable=import-error
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from .resources import Resource
from .utilities import status_msg

@dataclass_json
@dataclass(order=True)
class BomPart:
    """Part Information from Section of a BOM Parts Sheet"""
    # pylint: disable=too-many-instance-attributes
    part: str
    qty: float = field(compare=False)
    smallest:float = field(compare=False)
    biggest: float = field(compare=False)
    percent: float = field(compare=False)  # FT field
    description: str = field(compare=False)
    uom: str = field(compare=False)
    unitprice: float = field(compare=False)
    oem: str = field(compare=False)
    vendorpart: str = field(compare=False)
    vendor: str = field(compare=False)
    updated: datetime = field(compare=False)

@dataclass_json
@dataclass(order=True)
class BomSection:
    """Group of BOM Parts"""
    name: str
    parts: dict[str, BomPart] = field(compare=False)

@dataclass_json
@dataclass(order=True)
class Bom:
    """BOM sheet"""
    name: str
    beam: str = field(compare=False)
    smallest: float = field(compare=False)
    biggest: float = field(compare=False)
    sizes: list[float] = field(compare=False)
    sections: list[BomSection] = field(compare=False)


def find_excel_files_in_dir(base: Path) -> list[Path]:
    """get list of spreadsheets in folder"""
    return list(base.glob('[!~]*.xlsx'))

def make_bom_part(row, resources: dict[str, Resource]) -> BomPart:
    """Create bom part from row in spreadsheet"""
    qty: float = float(row[0].value)
    smallest: float = (
        0.0 if row[1].value is None else float(row[1].value))
    biggest: float = (
        0.0 if row[2].value is None else float(row[2].value))
    percent: float = (
        0.0 if row[3].value is None else float(row[3].value))
    part: str = str(row[5].value)

    # fww resolve at later point, we need to throw errors on fail
    try:
        resource = resources[part]
    except KeyError:
        resource = Resource(part, "Unknown", "EA", 0.0, "Unknown", part,
                            "Unknown", datetime.now())

    return BomPart(part,
                   qty,
                   smallest,
                   biggest,
                   percent,
                   resource.description,
                   resource.uom,
                   resource.unitprice,
                   resource.oem,
                   resource.vendorpart,
                   resource.vendor,
                   resource.updated)

def section_add_part(parts: dict[str, BomPart], part: BomPart) -> None:
    """Insert new part or add qty to existing part"""
    try:
        item: BomPart = parts[part.part]
        item.qty += part.qty
    except KeyError:
        parts[part.part] = part


def get_hull_sizes(sheet: Worksheet) -> list:
    """find all hull sizes listed in sheet"""
    sizes = []
    for values  in sheet.iter_rows(
            min_row=1,max_row=1,min_col=13,values_only=True):
        for value in values:
            if value:
                sizes.append(float(value))
    return sizes

def get_bom_sections(sheet: Worksheet,
                     resources: dict[str, Resource]) -> list[BomSection]:
    """read in BOM items into sections"""
    sections: list[BomSection] = []
    for row in sheet.iter_rows(min_row=18,max_col=8):
        qty: Optional[Union[str, int, float]] = row[0].value
        if isinstance(qty, str) and qty != "QTY":
            if qty == 'CANVAS':
                continue
            if 'section' in locals():
                sections.append(section)
            section: BomSection = BomSection(qty, {})
        elif isinstance(qty, (float, int)):
            bom_part: BomPart = make_bom_part(row,  resources)
            status_msg(f"    {bom_part}",3)
            section_add_part(section.parts, bom_part)
    sections.append(section)
    return sections

def load_bom(xlsx_file: Path, resources: dict[str, Resource]) -> Bom:
    """load individual BOM sheet"""
    status_msg(f'  {xlsx_file.name}', 2)
    xlsx: Workbook = load_workbook(xlsx_file.as_posix(), data_only=True)
    sheet: Worksheet = xlsx.active

    name: str = str(sheet["A1"].value)
    beam: str = "" if sheet["A2"].value is None else sheet["A2"].value
    smallest: float = float(
        0 if sheet["M1"].value == "ANY" else sheet["G13"].value)
    biggest:float  = float(
        0 if sheet["M1"].value == "ANY" else sheet["G14"].value)
    sizes = [] if smallest == 0 else get_hull_sizes(sheet)
    sections: list[BomSection] = get_bom_sections(sheet, resources)
    bom: Bom = Bom(name, beam, smallest, biggest, sizes, sections)
    xlsx.close()
    return bom

def load_boms(bom_folder: Path,
              resources: dict[str, Resource]) -> dict[str, Bom]:
    """load all BOM sheets"""
    status_msg('Loading BOMs', 1)
    bom_files: list[Path] = find_excel_files_in_dir(bom_folder)
    boms: dict[str, Bom] = {}
    for bom_file in bom_files:
        bom = load_bom(bom_file, resources)
        boms[bom.name] = bom
    return boms

if __name__ == "__main__":
    pass
