#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from datetime import date
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
from .boms import Bom, MergedBom
from .costing_data import (FileNameInfo, SectionInfo, XlsxBom,
                           BOM_COLUMNS, BOM_STYLES)
from .costing_headers import generate_header
from .costing_merge import get_bom
from .costing_sections import generate_sections
from .costing_totals import generate_totals
from .models import Model
from .settings import Settings
from .utilities import normalize_size, status_msg, SHEETS_FOLDER, SUBJECT
from . import config

# UTILITY FUNCTIONS ===========================================================

def build_name(size: str, model: Model, folder: str) -> FileNameInfo:
    """build file name for sheet

    Arguments:
        size: str  -- length of boat ending in .0 or .5 from str(float)
        model: Model -- name of model of boat

    Returns:
        dict -- size of boat as text  such as 18' or 18'6"
                model: of boat as text
                option: of boat as text
                with_options: text of model + option
                size_with_options: size + model + option
                file_name: full absolute path to file
    """
    sized = normalize_size(float(size))
    option: str = "" if model.sheet2 is None else ' ' + model.sheet2
    size_with_options: str = sized + ' ' + model.sheet1 + option
    size_with_folder: str = sized + ' ' + folder
    file_name: Path = SHEETS_FOLDER / folder / (size_with_folder + '.xlsx')
    name: FileNameInfo = {
        'size': sized,
        'model': model.sheet1,
        'option': option,
        'folder': folder,
        'with_options': model.sheet1 + option,
        'size_with_options': size_with_options,
        'size_with_folder': size_with_folder,
        'file_name': file_name,
    }
    return name


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

def generate_sheet(merged_bom: MergedBom,
                   file_name_info: FileNameInfo,
                   settings: Settings,
                   size: str) -> None:
    """genereate costing sheet

    Arguments:
        filterd_bom -- bom with only parts from the current size
        name -- parts and full name of current sheet
        file_name -- filename with full pathing to xls sheet to be created
        settings -- consumables, labor rates, mark ups
        size -- size of boat as text from float 18.5, 21, etc

    Returns:
        None
    """
    # create parent folder if necessay
    file_name_info['file_name'].parent.mkdir(parents=True, exist_ok=True)

    # create new workbook / xlsx file
    section_info: dict[str, SectionInfo] =  {}
    with Workbook(file_name_info['file_name'],
                  {'remove_timezone': True}) as workbook:
        xlsx: XlsxBom = XlsxBom(workbook)
        xlsx.bom = merged_bom
        xlsx.size = size
        xlsx.settings = settings
        xlsx.file_name_info = file_name_info

        xlsx.workbook.set_properties(properties(xlsx))
        xlsx.add_worksheet()
        xlsx.set_active('Sheet1')
        xlsx.sheet.set_default_row(12.75)
        xlsx.load_formats(BOM_STYLES)
        xlsx.columns = BOM_COLUMNS
        xlsx.apply_columns()

        generate_header(xlsx)
        generate_sections(xlsx, section_info)
        generate_totals(xlsx, section_info)


# MODEL/SIZE IETERATION FUNCTIONS =============================================
def generate_sheets_for_model(boms: dict[str, Bom],
                              model: Model,
                              settings: Settings) -> None:
    """"cycle through each size to create sheets
    * build the filname and size as as a text name
    * filter out parts that are not needed for this size of boat and correct
      costing for size of boat if necessay
    * computing section sizes is done in genereate_sheet

    Arguments:
        boms --  all boats/cabin boms
        model -- Model of boat to process
        settings -- consumables, labor rates, mark ups

    Returns:
        None
    """
    status_msg(f"  {model.folder}", 1)
    for size in boms[model.sheet1].sizes:
        file_name_info: FileNameInfo
        file_name_info = build_name(size, model, model.folder)
        status_msg(f"    {file_name_info['file_name']}", 2)
        merged_bom: MergedBom = get_bom(boms, model, size)
        generate_sheet(merged_bom, file_name_info, settings, str(size))


def generate_sheets_for_all_models(boms: dict[str, Bom],
                                   models: dict[str, Model],
                                   settings: Settings) -> None:
    """" cycle through each sheet/option combo to create sheets

    Arguments:
        target_parts -- model info for boats/cabins
        source_parts -- all boms for boats/cabins
        settings -- consumables, labor rates, mark ups

    Returns:
        None
    """
    status_msg("Generating Sheets", 1)
    if config.hgac:
        status_msg("Generating HGAC Sheets", 0)
    for model in models:
        generate_sheets_for_model(boms, models[model], settings)

if __name__ == "__main__":
    pass
