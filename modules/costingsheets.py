#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from copy import deepcopy
from datetime import date
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
# from xlsxwriter.utility import xl_rowcol_to_cell # type: ignore
from .boms import Bom, BomPart
from .costing_data import FileNameInfo, SectionInfo, Xlsx, COLUMNS, STYLES
from .costing_headers import generate_header
from .costing_sections import generate_sections
from .costing_totals import generate_totals
from .models import Model
from .utilities import (logger, normalize_size, status_msg, SHEETS_FOLDER,
                        SUBJECT)

# UTILITY FUNCTIONS ===========================================================

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
    size_with_folder: str = normalize_size(size) + ' ' + folder
    file_name: Path = SHEETS_FOLDER / folder / (size_with_folder + '.xlsx')
    name: FileNameInfo = {
        'size': normalize_size(size),
        'model': model.sheet1,
        'option': option,
        'with_options': model.sheet1 + option,
        'size_with_options': size_with_options,
        'size_with_folder': size_with_folder,
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
            target_part.qty = ((target_part.qty  or 0) +
                               (source_parts[key].qty or 0))
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
        # logger.debug("bom2 not found error %s %s  %s",
        #             model.sheet1, model.sheet2, model.folder)
        return bom1
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
            if ((v.smallest or 0) == 0 or size >= (v.smallest or 0)) and
            ((v.biggest or 0) == 0 or size <= (v.biggest or 0))
        }
        for v in section.parts.values():
            if v.percent:
                v.unitprice = size / (v.percent or 8) * (v.unitprice or 0)
    return bom


# WRITING SHEET FUNCTIONS =====================================================
def properties(xlsx):
    """set sheet properties"""
    return {
        'title': xlsx.file_name_info['size_with_options'],
        'subject': SUBJECT,
        'author': 'Sara Lynn',
        'company': 'North River Boats Inc.',
        'created': date.today(),
        'comments': 'Created with Python and XlsxWriter',
    }

def generate_sheet(filtered_bom: Bom, file_name_info: FileNameInfo) -> None:
    """genereate costing sheet

    Arguments:
        filterd_bom: Bom -- bom with only parts from the current size
        name: dict -- parts and full name of current sheet
        file_name: Path -- filename with full pathing to xls sheet to be
                           created

    Returns:
        None
    """
    # create parent folder if necessay
    file_name_info['file_name'].parent.mkdir(parents=True, exist_ok=True)

    # create new workbook / xlsx file
    section_info: dict[str, SectionInfo] =  {}
    with Workbook(file_name_info['file_name'],
                  {'remove_timezone': True}) as workbook:
        xlsx: Xlsx = Xlsx(workbook, filtered_bom)

        xlsx.file_name_info = file_name_info
        xlsx.workbook.set_properties(properties(xlsx))
        xlsx.add_worksheet()
        xlsx.set_active('Sheet1')
        xlsx.sheet.set_default_row(12.75)
        xlsx.load_formats(STYLES)
        xlsx.columns = COLUMNS
        xlsx.apply_columns()

        generate_header(xlsx)
        generate_sections(xlsx, section_info)
        generate_totals(xlsx, section_info)


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
    # for key in {"SOUNDER 8'6'' OPEN": models["SOUNDER 8'6'' OPEN"]}:  # fww
    for model in models:  # fww
        bom: Bom = get_bom(boms, models[model])
        generate_sheets_for_model(models[model], bom)

if __name__ == "__main__":
    pass
