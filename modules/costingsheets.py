#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
Generate Costing Sheets
"""
from copy import copy, deepcopy
# from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
import os
# from typing import Optional, Union
from openpyxl import load_workbook # pylint: disable=import-error
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook
from .boms import Bom, BomPart
from .models import Model
from .utilities import (logger, normalize_size, status_msg, SHEETS_FOLDER,
                        TEMPLATE_FILE,)

@dataclass
class SectionInfo:
    """Information about each section on xlsx sheet"""
    name: str
    start: int
    finish: int
    total: int
    offset: int = field(init=False)
    break_point: int = field(init=False)
    lines: int = field(init=False)

    def __post_init__(self):
        self.offset = 0
        self.break_point = self.total
        self.lines = self.finish - self.start + 1

section_info_master: list[SectionInfo] = [
        SectionInfo('FABRICATION', 14, 17, 20),
        SectionInfo('PAINT', 26, 38, 40),
        SectionInfo('OUTFITTING', 48, 70, 72),
        SectionInfo('CANVAS', 77, 139, 141),
        SectionInfo('BIG TICKET ITEMS', 148, 149, 151),
        SectionInfo('OUTBOARD MOTORS', 156, 158, 160),
        SectionInfo('INBOARD MOTORS & JETS', 165, 168, 170),
        SectionInfo('TRAILER', 175, 175, 177),
]


# UTILITY FUNCTIONS ===========================================================
def build_name(size: float, model: Model) -> dict[str, str]:
    """build file name for sheet

    Arguments:
        size: float  -- length of boat ending in .0 or .5
        model: Model -- name of model of boat

    Returns:
        dict -- size of boat as text  such as 18' or 18'6"
                model of boat as text
                option of boat as text
                full text of model + option
                all text size + model + option
    """
    option: str = "" if model.sheet2 is None else ' ' + model.sheet2
    name: dict[str, str] = {
        'size': normalize_size(size),
        'model': model.sheet1,
        'option': option,
        'full': model.sheet1 + option,
        'all': normalize_size(size) + ' ' + model.sheet1 + option,
    }
    return name


# BOM MANIPULATION FUNCTIONS ==================================================
def ordered_parts(parts: dict[str, BomPart], name: str) -> list[str]:
    """set correct sort order for parts for each section. Outftting is sorted
    by vender then part number. All other sections are ordered by part number

    Arguments:
        parts: dict[str, BomPart] -- dictionary of parts with the key being a
                                     part number
        name: str -- name of section

    Returns:
        list[str] -- Correctly ordered list of keys
    """
    if name != 'OUTFITTING':
        return sorted(parts)
    return sorted(parts, key=lambda k: (parts[k].vendor, k))


def bom_merge_section(target_parts: dict[str, BomPart],
                      source_parts: dict[str, BomPart]) -> None:
    """Merege two sections

    Arguments:
        target_parts: dict[str, BomPart] -- longer dict of BomParts to append
                                            to or merge part with
        source_parts: dict[str, BomPart] -- shorter dict of BomParts to add or
                                            merge parts from

    Returns:
        None -- target_parts modified in place
    """
    for key in source_parts:
        try:
            target_part: BomPart = target_parts[key]
            target_part.qty += source_parts[key].qty
        except KeyError:
            target_parts[key] = deepcopy(source_parts[key])

def bom_merge(target_bom: Bom, source_bom: Bom) -> Bom:
    """Merge two BOMs creating a new BOM in the process. Does not modify either
    original Bom.

    Arguments:
        target_bom: Bom -- sheet1 bom we want to merge parts into
        source_bom: Bom -- sheet2 bom with parts to be added/merged

    Returns:
        Bom -- new merged list of both boms combined
    """
    bom: Bom = deepcopy(target_bom)
    for target, source in zip(bom.sections, source_bom.sections):
        bom_merge_section(target.parts, source.parts)
    return bom

def get_bom(boms: dict[str, Bom], model: Model) -> Bom:
    """Combine sheets if necessary and return BOM.
       Assumes if sheet is not None that there will be a match

    Arguments:
        bom: list[Bom] --
        model: Model -- sheet1 can not be None and must be found
                        sheet2 can be None but *must* be found if not None

    Returns:
        Bom -- Returns new Bom of combined Bom(s)
    """
    bom1: Bom = (boms[model.sheet1]
                 if model.sheet1 in boms
                 else  Bom('', "", 0.0, 0.0, [], []))
    bom2: Bom = (boms[model.sheet2]
                 if model.sheet2 in boms
                 else  Bom('', "", 0.0, 0.0, [], []))
    if bom1.name == "":
        logger.debug("bom1 not found error %s", model.sheet1)
    if bom2.name == "":
        logger.debug("bom2 not found error %s %s  %s",
                     model.sheet1, model.sheet2, model.folder)
    return bom_merge(bom1, bom2)


# WRITING SHEET FUNCTIONS =====================================================
def generate_section(xxxxx, yyyyy, zzzzz):
    """this is a doc string"""
    print(xxxxx, yyyyy, zzzzz)

def compute_section_sizes(bom_sections, section_info) -> None:
    """x"""
    offset = 0
    print("sta fin tot lin off bp     sta fin tot off bp")
    print("--- --- --- --- --- ---    --- --- --- --- ---")
    for section, info in zip(bom_sections, section_info):
        s = info.start
        f = info.finish
        t = info.total
        o = offset
        b = info.break_point

        info.offsest = offset
        items = len(section.parts) - info.lines
        info.start += offset
        offset = offset + items
        info.finish += offset
        info.total += offset
        print(f"{s:3}-{f:3} {t:3} {items:3} {o:3} {b:3}    "
              f"{info.start:3}-{info.finish} {info.total:3} {offset:3} {b:3}")
    print()


    pass

def resize_sections():
    """x"""
    pass

def generate_sections(bom: Bom, sheet: Worksheet) -> None:
    """Manage filling in sections

    Works by iterating over the sections of the sheet from the last section to
    the first section. Order matters here.
    * bom.Sections were built from first section to last section. Needs to be
      iterated over in reverse order
    * section_info was constructed in reverse order

    Arguements:
    """
    for section, info in reversed(list(zip(bom.sections, section_info))):
        # offset: int = generate_section(sheet, section, info)
        # status_msg(f"        {section.name:25} {offset}",3)
        print(section.name, info.name)

def generate_heading(bom: Bom, name: dict[str, str], sheet: Worksheet) -> None:
    """Fill out heading at top of sheet
    Arguments:
        bom: Bom -- bom with information for all sizes of current model/option
        name: dict -- parts and full name of current sheet

    Returns:
        None
    """
    sheet["C4"].value = name['full'].title()
    sheet["C5"].value = name['size']
    sheet["C6"].value = bom.beam

def generate_sheet(bom: Bom,
                   name: dict[str, str],
                   file_name: Path) -> None:
    """genereate costing sheet

    Arguments:
        bom: Bom -- bom with information for all sizes of current model/option
        name: dict -- parts and full name of current sheet
        file_name: Path -- filename with full pathing to xls sheet to be
                           created

    Returns:
        None
    """
    #  file_name.parent.mkdir(parents=True, exist_ok=True)
    xlsx: Workbook = load_workbook(
        TEMPLATE_FILE.as_posix(), data_only=False)
    sheet: Worksheet = xlsx.active
    generate_heading(bom, name, sheet)
    generate_sections(bom, sheet)
    status_msg(f"      {name['all']}",2)
    xlsx.save(os.path.abspath(str(file_name)))


# MODEL/SIZE IETERATION FUNCTIONS =============================================
def generate_sheets_for_model(model: Model,
                              bom: Bom) -> None:
    """"cycle through each size to create sheets

    Arguments:
        model: Model -- Model of boat to process
        bom: Bom -- Base Bom for model

    Returns:
        None
    """
    for size in bom.sizes:
        name: dict[str, str] = build_name(size, model)
        file_name: Path = (SHEETS_FOLDER / model.folder /
                           (name['all'] + '.xlsx'))
        generate_sheet(bom, name, file_name)


def generate_sheets_for_all_models(models: dict[str, Model],
                                   boms: dict[str, Bom]) -> None:
    """" cycle through each sheet/option combo to create sheets

    Arguments:
        target_parts: dict[str, BomPart] -- longer dict of BomParts to append
                                            to or merge part with
        source_parts: dict[str, BomPart] -- shorter dict of BomParts to add or
                                            merge parts from

    Returns:
        None
    """
    status_msg("Merging", 1)
    # for model in models:
    # for model in models:  # fww
    xlsx: Workbook = load_workbook(TEMPLATE_FILE.as_posix(), data_only=False)
    sheet: Worksheet = xlsx.active
    sheet.insert_rows(78,5)

    for col in range(1,16):
        cell = sheet.cell(row=77, column=col)
        style = copy(cell.style)
        font = copy(cell.font)
        border = copy(cell.border)
        number_format = copy(cell.number_format)
        alignment = copy(cell.alignment)

        for row in range(78,83):
            cell = sheet.cell(row=row, column=col)
            cell.style = copy(style)
            cell.font = copy(font)
            cell.border = copy(border)
            cell.number_format = copy(number_format)
            cell.alignment = copy(alignment)
            if col == 5:
                cell.value = 'ea'
            if col == 7:
                cell.value = f'=D{row}*F{row}'
            if col == 8:
                cell.value = 0
            if col == 9:
                cell.value = f'=G{row}+H{row}'

    xlsx.save(os.path.abspath(str('/home/fwarren/test.xlsx')))
    return
    for key in {"SOUNDER 8'6'' OPEN": models["SOUNDER 8'6'' OPEN"]}:  # fww
        bom = get_bom(boms, models[key])
        # status_msg(f"  {models[key].folder}", 1)
        # generate_sheets_for_model(models[key], bom)
        section_info = deepcopy(section_info_master)
        compute_section_sizes(bom.sections, section_info)

if __name__ == "__main__":
    pass
