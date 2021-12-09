#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from datetime import date
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
from .boms import Bom, Boms, MergedBom
from .costing_data import FileNameInfo, SectionInfo, Xlsx, COLUMNS, STYLES
from .costing_headers import generate_header
from .costing_merge import get_bom
from .costing_sections import generate_sections
from .costing_totals import generate_totals
from .models import Model
from .settings import Settings
from .utilities import normalize_size, status_msg, SHEETS_FOLDER, SUBJECT

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

def generate_msrp_xlsx(msrps: dict[str, float]) -> None:
    """genereate costing msrp sheet

    Arguments:
        filterd_bom -- bom with only parts from the current size
        name -- parts and full name of current sheet
        file_name -- filename with full pathing to xls sheet to be created
        settings -- consumables, labor rates, mark ups
        size -- size of boat as text from float 18.5, 21, etc

    Returns:
        None
    """
    for name, new_msrp in msrps.items():
        new_msrp = 123495.0
        new_iff: float =  (new_msrp - (new_msrp * 0.3))/0.9925
        old_iff: float = new_iff
        old_msrp: float = old_iff * 0.9925 * 1.020304051
        print(f"{name:35.35}  {new_msrp:9.2f}  {new_iff:9.2f}  "
              f"{old_msrp:9.2f}  {old_iff:9.2f}")


    # create parent folder if necessay
    #file_name_info['file_name'].parent.mkdir(parents=True, exist_ok=True)

    ## create new workbook / xlsx file
    #with Workbook(file_name_info['file_name'],
    #              {'remove_timezone': True}) as workbook:
    #    xlsx: Xlsx = Xlsx(workbook, merged_bom, 99, settings)

    #    xlsx.file_name_info = file_name_info
    #    xlsx.workbook.set_properties(properties(xlsx))
    #    xlsx.add_worksheet()
    #    xlsx.set_active('Sheet1')
    #    xlsx.sheet.set_default_row(12.75)
    #    xlsx.load_formats(STYLES)
    #    xlsx.columns = COLUMNS
    #    xlsx.apply_columns()

        # generate_header(xlsx)
        # generate_sections(xlsx, section_info)
        # generate_totals(xlsx, section_info)


# MODEL/SIZE IETERATION FUNCTIONS =============================================
def get_msrp(boms: dict[str, Bom],
             model: Model,
             settings: Settings,
             size: float) -> tuple[str, float]:
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
    file_name_info: FileNameInfo
    file_name_info = build_name(size, model, model.folder)
    status_msg(f"    {file_name_info['file_name']}", 2)
    merged_bom: MergedBom = get_bom(boms, model, size)
    # collect info here
    dept = 'Boat and options'
    markup_1 = settings.mark_ups[dept].markup_1
    markup_2 = settings.mark_ups[dept].markup_2
    rate_fabrication = settings.consumables['FABRICATION'].rate
    rate_paint = settings.consumables['PAINT'].rate

    value1 = (merged_bom.sections['FABRICATION'].total +
              merged_bom.sections['FABRICATION'].total * rate_fabrication +
              merged_bom.sections['PAINT'].total +
              merged_bom.sections['PAINT'].total * rate_paint +
              merged_bom.sections['OUTFITTING'].total)
    value2: float = value1 / markup_1 / markup_2
    msrp: float = (int(value2 / 100) * 100.0) + 95
    return file_name_info['size_with_folder'], msrp




def generate_msrp_summary(boms: dict[str, Bom],
                          models: dict[str, Model],
                          settings: Settings) -> None:
    """" cycle through each sheet/option combo to create report

    Arguments:
        target_parts -- model info for boats/cabins
        source_parts -- all boms for boats/cabins
        settings -- consumables, labor rates, mark ups

    Returns:
        None
    """
    status_msg("Generating Sheets", 1)
    msrps: dict(str, float) = {}
    models = dict(sorted(models.items()))
    for model in models.values():
        status_msg(f"  {model.folder}", 1)
        for size in boms[model.sheet1].sizes:
            name, msrp = get_msrp(boms, model, settings, size)
            msrps[name] = msrp
    generate_msrp_xlsx(msrps)

if __name__ == "__main__":
    pass
