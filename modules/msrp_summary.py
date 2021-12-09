#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
from .boms import Bom, MergedBom
from .costing_data import Columns, FileNameInfo, Format, Xlsx
from .costing_merge import get_bom
from .models import Model
from .settings import Settings
from .utilities import normalize_size, status_msg, SHEETS_FOLDER, SUMMARY

@dataclass
class Msrp():
    """track msrp price and shade of cells for row group"""
    msrp: float
    shade: str
    model: Model

# WORKBOOK LAYOUT ============================================================
MSRP_PROPERTIES = {
    'title': SUMMARY,
    'subject': 'MSRP Report',
    'author': 'Sara Lynn',
    'company': 'North River Boats Inc.',
    'created': date.today(),
    'comments': 'Created with Python and XlsxWriter',
}

MSRP_COLUMNS = [                   # PIXELS   POINTS
    Columns('A:A', 250, 'normal'), # 160.00   100.80
    Columns('B:E', 160, 'normal'), # 120.00   104.75
]

CURRENCY = (
    '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
)

MSRP_STYLES = [
    Format(
        'normal',
        {
            'font_name': 'arial',
            'font_size': 10,
        },
    ),
    Format(
        'normalBold',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'pattern': 1,
            'bg_color': '#666666',
            'color': '#FFFFFF',
        },
    ),
    Format(
        'currency',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
        },
    ),
    Format(
        'currencyBold',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'num_format': CURRENCY,
        },
    ),
]


SHADES = [
    '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF',
    '#FF99CC', '#CC99FF', '#FFCC99', '#C0C0C0',
    '#9999FF', '#FFFFCC', '#CCFFFF', '#FF8080',
    '#CCCCFF', '#33CCCC', '#99CC00', '#FFCC00',
    '#FF9900', '#FF6600', '#969696', '#FFD8A0',
    '#C5F19A', '#FEF58A', '#B9D0E8', '#F5DDB7',
    '#D6BFD4', '#F79494', '#D3D7CF', '#993366',
    '#0066CC', '#3366FF', '#666699', '#808080',
]


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
def generate_msrp_xlsx(xlsx: Xlsx,
                       msrps: dict[str, Msrp]) -> None:
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
    xlsx.write(0, 0, 'BOAT', xlsx.styles['normalBold'])
    xlsx.write(0, 1, 'GSA NEW TERMS MSRP', xlsx.styles['normalBold'])
    xlsx.write(0, 2, 'GSA NEW WITH IFF', xlsx.styles['normalBold'])
    xlsx.write(0, 3, 'GSA OLD TERMS MSRP', xlsx.styles['normalBold'])
    xlsx.write(0, 4, 'GSA OLD WITH IFF', xlsx.styles['normalBold'])
    normal: str = ''
    for row, msrp in enumerate(msrps.items(), start=1):
        name: str = msrp[0]
        shade: str = msrp[1].shade
        model: Model = msrp[1].model
        new_msrp: float = msrp[1].msrp
        new_iff: float =  round((new_msrp - (new_msrp * 0.3))/0.9925, 2)
        old_iff: float = new_iff
        old_msrp: float = round(old_iff * 0.9925 * 1.020304051, 2)

        status_msg(f"{name:35.35}  {new_msrp:9.2f}  {new_iff:9.2f}  "
                   f"{old_msrp:9.2f}  {old_iff:9.2f}", 2)

        if shade not in normal:
            normal = 'normal' + shade
            currency:str = 'currency' + shade
            xlsx.styles[normal] = xlsx.workbook.add_format({
                'pattern': 1,
                'bg_color': shade,
                'font_name': 'arial',
                'font_size': 10,
            })
            xlsx.styles[currency] = xlsx.workbook.add_format({
                'pattern': 1,
                'bg_color': shade,
                'font_name': 'arial',
                'font_size': 10,
                'num_format': CURRENCY,
            })

        link = "external:" + model.folder + '/' + name + ".xlsx"
        xlsx.write(row, 0, link, xlsx.styles[normal], name)
        xlsx.write(row, 1, new_msrp, xlsx.styles[currency])
        xlsx.write(row, 2, new_iff, xlsx.styles[currency])
        xlsx.write(row, 3, old_msrp, xlsx.styles[currency])
        xlsx.write(row, 4, old_iff, xlsx.styles[currency])
    xlsx.sheet.freeze_panes(1,0)


# MODEL/SIZE IETERATION FUNCTIONS =============================================
def get_msrp(boms: dict[str, Bom],
             model: Model,
             settings: Settings,
             size: str) -> tuple[str, float]:
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
    name: str
    msrp: float
    status_msg("Generating Sheets", 1)
    msrps: dict[str, Msrp] = {}
    models = dict(sorted(models.items()))
    model_index = set()

    for model in models.values():
        model_index.add(model.sheet1)
        index: int = len(model_index) - 1

        status_msg(f"  {model.folder}", 1)
        for size in boms[model.sheet1].sizes:
            name, msrp = get_msrp(boms, model, settings, size)
            msrps[name] = Msrp(msrp, SHADES[index], model)

    file_name = SHEETS_FOLDER / (SUMMARY + '.xlsx')
    with Workbook(file_name, {'remove_timezone': True}) as workbook:
        xlsx: Xlsx = Xlsx(workbook)
        xlsx.setup_workbook(MSRP_STYLES, MSRP_COLUMNS, MSRP_PROPERTIES)
        generate_msrp_xlsx(xlsx, msrps)

if __name__ == "__main__":
    pass
