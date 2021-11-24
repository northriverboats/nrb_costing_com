#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheets
"""
from copy import deepcopy
from dataclasses import dataclass, field
from typing import Any, Optional, TypedDict
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
from .boms import Bom, BomPart
from .models import Model
from .utilities import (logger, normalize_size, status_msg, SHEETS_FOLDER)

COLUMNS = [
    {'columns': 'A:A', 'width': 17.40, 'style': 'generic1'},
    {'columns': 'B:B', 'width': 18.10, 'style': 'generic1'},
    {'columns': 'C:C', 'width': 44.56, 'style': 'generic1'},
    {'columns': 'D:D', 'width': 15.29, 'style': 'generic1'},
    {'columns': 'E:E', 'width': 4.86, 'style': 'generic1'},
    {'columns': 'F:F', 'width': 13.14, 'style': 'generic1'},
    {'columns': 'G:G', 'width': 12.86, 'style': 'generic1'},
    {'columns': 'H:H', 'width': 9.86, 'style': 'generic1'},
    {'columns': 'I:I', 'width': 17.00, 'style': 'generic1'},
    {'columns': 'J:T', 'width': 13.00, 'style': 'generic1'},
]


# DATA CLASSES ================================================================
@dataclass
class Xlsx():
    """Bundle up xlsxwriter information

    Arguments:
        workbook: obj -- xlswriter workbook class object

    Returns:
        None
    """
    workbook: Any
    bom: Bom
    sheet: Any = field(default=None)
    styles: dict = field(init=False, default_factory=dict)
    worksheets: dict = field(init=False, default_factory=dict)

    def add_worksheet(self, name: Optional[str] = None) -> None:
        """add new sheet to workbook

        Arguments:
            name: str -- name of worksheet, if None system will name

        Raise:
            uplicateWorksheetName -– if a duplicate worksheet name is used.
            InvalidWorksheetName -– if an invalid worksheet name is used.
            ReservedWorksheetName -– if a reserved worksheet name is used.

        Returns:
            None
        """
        if name:
            worksheet = self.workbook.add_worksheet(name)
        else:
            worksheet = self.workbook.add_worksheet()
        self.worksheets[worksheet.get_name()] = worksheet

    def set_active(self, name):
        """set active worksheet"""
        self.sheet = self.worksheets[name]

    def write(self, *args):
        """write value to sheet"""
        return self.sheet.write(*args)

    def add_format(self, name, *args):
        """add new formatter"""
        style = self.workbook.add_format(*args)
        self.styles[name] = style



# UTILITY FUNCTIONS ===========================================================
class FileNameInfo(TypedDict):
    """file name info"""
    size: str
    model: str
    option: str
    with_options: str
    size_with_options: str
    file_name: Path

def build_name(size: float, model: Model, folder: str) -> FileNameInfo:
    """build file name for sheet

    Arguments:
        size: float  -- length of boat ending in .0 or .5
        model: Model -- name of model of boat

    Returns:
        dict -- size of boat as text  such as 18' or 18'6"
                model: of boat as text
                option: of boat as text
                with_options: text of model + option
                size_with_options: size + model + option
                file_name: full absolute path to file
    """
    option: str = "" if model.sheet2 is None else ' ' + model.sheet2
    size_with_options: str = normalize_size(size) + ' ' + model.sheet1 + option
    file_name: Path = SHEETS_FOLDER / folder / (size_with_options + '.xlsx')
    name: FileNameInfo = {
        'size': normalize_size(size),
        'model': model.sheet1,
        'option': option,
        'with_options': model.sheet1 + option,
        'size_with_options': size_with_options,
        'file_name': file_name,
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
    """Merges sheets if necessary and returns a BOM.
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

def filter_bom(original_bom: Bom, size: float) -> Bom:
    """fitler out parts based on size if necessary
    also correcting costs as necessary

    Arguments:
        bom: Bom -- Bom with parts sections
        size: float -- size of boat to filter for

    Returns:
        Bom -- Deepcopy Bom with correct parts
    """
    bom: Bom = deepcopy(original_bom)
    for section in bom.sections:
        section.parts = {
            k:v
            for k, v in section.parts.items()
            if (v.smallest == 0 or size >= v.smallest) and
            (v.biggest == 0 or size <= v.biggest)
        }
        for v in section.parts.values():
            if v.percent:
                v.unitprice = size / v.percent * v.unitprice
    return bom


# WRITING SHEET FUNCTIONS =====================================================
def create_formats(xlsx) -> None:
    """create necessary cell formatting"""
    xlsx.add_format('generic1', {
        'font_name': 'Arial',
        'font_size': 10})

    xlsx.add_format('headingCustomer1', {
        'font_name': 'Arial',
        'font_size': 18,
        'bold': True})

    xlsx.add_format('headingCustomer2', {
        'font_name': 'Arial',
        'font_size': 20,
        'bold': True,
        'pattern': 1,
        'bg_color': 'yellow'})

def set_column_widths(xlsx) -> None:
    """set column width and default style"""
    for col in COLUMNS:
        xlsx.sheet.set_column(col['columns'], col['width'])

def generate_sheet(filtered_bom: Bom, file_name_info: FileNameInfo) -> None:
    """genereate costing sheet

    Arguments:
        filterd_bom: Bom -- bom with only parts fr the current size
        name: dict -- parts and full name of current sheet
        file_name: Path -- filename with full pathing to xls sheet to be
                           created

    Returns:
        None
    """
    # create parent folder if necessay
    file_name_info['file_name'].parent.mkdir(parents=True, exist_ok=True)
    # caculate the size of each section

    # create new workbook / xlsx file
    with Workbook(file_name_info['file_name']) as workbook:
        xlsx: Xlsx = Xlsx(workbook, filtered_bom)
        xlsx.add_worksheet()
        xlsx.set_active('Sheet1')
        create_formats(xlsx)
        set_column_widths(xlsx)
        xlsx.sheet.set_row(1, 26.25)
        xlsx.write('B2', 'Customer:', xlsx.styles['headingCustomer1'])
        xlsx.write('C2', '', xlsx.styles['headingCustomer2'])
        # write custom header
        # create necessary formats
        # dump formats
        # setup column defaults
        # write header
        # write sections
        # write footer


# MODEL/SIZE IETERATION FUNCTIONS =============================================
def generate_sheets_for_model(model: Model, bom: Bom) -> None:
    """"cycle through each size to create sheets
    * build the filname and size as as a text name
    * filter out parts that are not needed for this size of boat and correct
      costing for size of boat if necessay
    * computing section sizes is done in genereate_sheet

    Arguments:
        model: Model -- Model of boat to process
        bom: Bom -- Base Bom for model

    Returns:
        None
    """
    status_msg(f"  {model.folder}", 1)
    for size in bom.sizes:
        file_name_info: FileNameInfo = build_name(size, model, model.folder)
        status_msg(f"    {file_name_info['file_name']}", 2)
        filtered_bom = filter_bom(bom, size)
        generate_sheet(filtered_bom, file_name_info)


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
    # for model in models:  # fww
    for key in {"SOUNDER 8'6'' OPEN": models["SOUNDER 8'6'' OPEN"]}:  # fww
        # merge "model" sheet 1 Bom with "option" sheet 2 Bom
        bom: Bom = get_bom(boms, models[key])
        generate_sheets_for_model(models[key], bom)
        # for section in bom.sections:
        #   print(f"{section.name:25} {len(section.parts)}")

if __name__ == "__main__":
    pass
