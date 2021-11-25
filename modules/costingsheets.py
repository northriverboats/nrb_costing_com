#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheets
"""
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Optional, TypedDict
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
from .boms import Bom, BomPart
from .models import Model
from .utilities import (logger, normalize_size, status_msg, SHEETS_FOLDER,
                        SUBJECT)

# DATA CLASSES ================================================================
class FileNameInfo(TypedDict):
    """file name info"""
    size: str
    model: str
    option: str
    with_options: str
    size_with_options: str
    file_name: Path

@dataclass
class Columns():
    """column layout data and functions"""
    columns: str
    width: float
    style: Optional[str]

@dataclass
class Format():
    """format/style name and contents"""
    name: str
    style: dict[str, Any]

@dataclass
class SectionInfo():
    """Additional Section Information"""
    start: int
    finish: int
    subtotal: int
    value: float

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
    file_name_info: FileNameInfo = field(init=False)
    columns: list[Columns] = field(init=False, default_factory=list)

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

    def merge_range(self, *args):
        """write value to sheet"""
        return self.sheet.merge_range(*args)

    def add_format(self, name, *args):
        """add new formatter"""
        style = self.workbook.add_format(*args)
        self.styles[name] = style

    def apply_columns(self):
        """apply column formatting"""
        for col in self.columns:
            if col.style:
                self.sheet.set_column_pixels(col.columns,
                                             col.width,
                                             self.styles[col.style])
            else:
                self.sheet.set_column_pixels(col.columns, col.width)

    def load_formats(self, styles):
        """ load styles from a list of  Format objects"""
        # can parse styles.style and extract and do other processing
        for style in styles:
            self.add_format(style.name, style.style)


# SHEET DATA ==================================================================
COLUMNS = [                             # POINTS   PIXELS
    Columns('A:A', 126.50, 'generic1'), # 126.50   100.80
    Columns('B:B', 132, 'generic1'),    # 132.00   104.75
    Columns('C:C', 314.00, 'generic1'), # 314.00   249.20
    Columns('D:D', 106.50, 'generic1'), # 106.50    84.95
    Columns('E:E', 34, 'generic1'),     #  34.00    27.00
    Columns('F:F', 92, 'generic1'),     #  92.00    73.00
    Columns('G:G', 90, 'generic1'),     #  90.00    71.45
    Columns('H:H', 69, 'generic1'),     #  69.00    54.75
    Columns('I:I', 119, 'generic1'),    # 119.00    94.47
    Columns('J:T', 62, 'generic1'),     #  62.00    49.20
]

STYLES = [
     Format('generic1', {'font_name': 'Arial',
                        'font_size': 10,}),
     Format('generic2', {'font_name': 'Arial',
                        'font_size': 10,
                        'bold': True,}),

    Format('headingCustomer1', {'font_name': 'Arial',
                                'font_size': 18,
                                'bold': True,}),

    Format('headingCustomer2', {'font_name': 'Arial',
                                'font_size': 20,
                                'bold': True,
                                'pattern': 1,
                                'bg_color': '#FCF305',
                                'bottom': 1,}),

    Format('bgSilver', {'pattern':1,
                        'bg_color': 'silver',
                        'align': 'center',
                        'bold': True,}),

    Format('bgYellow0', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'align': 'center',}),
    Format('bgYellow1', {'pattern': 1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,}),
    Format('bgYellow2', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',}),
    Format('bgYellow3', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'border': 1,
                         'align': 'center',}),

    Format('bgGreen1', {'pattern':1,
                       'bg_color': '#1FB714',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,}),
    Format('bgGreen2', {'pattern':1,
                       'bg_color': '#1FB714',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,
                       'align': 'center',}),

    Format('bgPurple1', {'pattern':1,
                         'bg_color': '#CC99FF',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,}),
    Format('bgPurple2', {'pattern':1,
                         'bg_color': '#CC99FF',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',}),

    Format('bgCyan1', {'pattern': 1,
                       'bg_color': '#99CCFF',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,}),
    Format('bgCyan2', {'pattern':1,
                       'bg_color': '#99CCFF',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,
                       'align': 'center',}),

    Format('bgOrange1', {'pattern':1,
                         'bg_color': 'FF9900',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,}),

    Format('bgOrange2', {'pattern':1,
                         'bg_color': 'FF9900',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',}),

    Format('rightJust1', {'align': 'right',
                          'font_name': 'Arial',
                          'font_size': 10,}),
    Format('rightJust2', {'align': 'right',
                          'font_name': 'Arial',
                          'font_size': 10,
                          'bold': True}),
]

SECTION_TEST = {
    'FABRICATION': SectionInfo(20, 28, 29,10),
    'PAINT': SectionInfo(30, 38, 39,12),
    'UNUSED': SectionInfo(40, 48, 49,0),
    'OUTFITTING': SectionInfo(50, 58, 59,16),
    'BIG TICKET ITEMS': SectionInfo(60, 68, 69,18),
    'OUTBOARD MOTORS': SectionInfo(70, 78, 79,20),
    'INBOARD MOTORS & JETS': SectionInfo(80, 88, 89,22),
    'TRAILER': SectionInfo(90, 98, 99,24),
    'TOTALS': SectionInfo(100, 150, 152, 0),
}


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

def generate_header(xlsx: Xlsx) -> None:
    """generate header on costing sheet"""

    xlsx.sheet.set_row(1, 26.25)

    xlsx.write('B2', 'Customer:', xlsx.styles['headingCustomer1'])
    xlsx.write('C2', None, xlsx.styles['headingCustomer2'])
    xlsx.write('G2', 'Salesperson:')
    xlsx.merge_range('H2:I2', None, xlsx.styles['bgYellow2'])

    xlsx.write('B4', 'Boat Model:')
    xlsx.write('C4', xlsx.bom.name, xlsx.styles['bgYellow1'])
    xlsx.write('B5', 'Beam:')
    xlsx.write('C5', xlsx.bom.beam, xlsx.styles['bgYellow1'])
    xlsx.write('B6', 'Length:')
    xlsx.write('C6', xlsx.file_name_info['size'], xlsx.styles['bgYellow1'])

    xlsx.write('H4', 'Original Date Quoted:', xlsx.styles['rightJust1'])
    xlsx.write('I4', None, xlsx.styles['bgYellow1'])

    xlsx.merge_range('E5:G5',
                     'Indicate changes here',
                     xlsx.styles['bgGreen2'])
    xlsx.merge_range('E6:G6',
                     'Indicate changes here',
                     xlsx.styles['bgPurple2'])
    xlsx.merge_range('E7:G7',
                     'Indicate changes here',
                     xlsx.styles['bgCyan2'])
    xlsx.merge_range('E8:G8',
                     'Indicate changes here',
                     xlsx.styles['bgOrange2'])

    xlsx.write('H5', 'Rev1', xlsx.styles['rightJust1'])
    xlsx.write('H6', 'Rev2', xlsx.styles['rightJust1'])
    xlsx.write('H7', 'Rev3', xlsx.styles['rightJust1'])
    xlsx.write('H8', 'Rev4', xlsx.styles['rightJust1'])

    xlsx.write('I5', None, xlsx.styles['bgGreen1'])
    xlsx.write('I6', None, xlsx.styles['bgPurple1'])
    xlsx.write('I7', None, xlsx.styles['bgCyan1'])
    xlsx.write('I8', None, xlsx.styles['bgOrange1'])

def generate_totals(xlsx: Xlsx, section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet"""
    offset = section_info['TRAILER'].subtotal + 2

    # COLUMN B ================================================================


    # COLUMN C ================================================================
    xlsx.write(offset + 0, 2, 'MATERIALS', xlsx.styles['rightJust2'])
    xlsx.write(offset + 12, 2, 'Labor', xlsx.styles['rightJust2'])
    xlsx.write(offset + 20, 2,
               'Indicate boat referenced for labor hours if used',
               xlsx.styles['bgYellow0'])
    xlsx.write(offset + 23, 2, 'Other Costs', xlsx.styles['rightJust2'])
    xlsx.write(offset + 32, 2, 'NO MARGIN ITEMS', xlsx.styles['rightJust2'])
    xlsx.write(offset + 35, 2, 'Voyager/Custom - 10%, Guide/Lodge - 3%',
               xlsx.styles['bgYellow3'])

    xlsx.sheet.set_row(offset + 32, 23.85)
    xlsx.write(offset + 42, 2, 'Mark up per pricing policy: ',
               xlsx.styles['generic2'])
    xlsx.write(offset + 43, 2, 'Boat and options:',
               xlsx.styles['generic1'])
    xlsx.write(offset + 44, 2, 'Big Ticket Items',
               xlsx.styles['generic1'])
    xlsx.write(offset + 45, 2, 'OB Motors',
               xlsx.styles['generic1'])
    xlsx.write(offset + 46, 2, 'Inboard Motors & Jets',
               xlsx.styles['generic1'])
    xlsx.write(offset + 47, 2, 'Trailer',
               xlsx.styles['generic1'])
    xlsx.write(offset + 48, 2, 'No margin items: ',
               xlsx.styles['generic1'])
    xlsx.write(offset + 50, 2,
               'Total Cost (equals total cost of project box)',
               xlsx.styles['rightJust2'])

    xlsx.write(offset + 59, 2, 'Pricing Policy References: ',
               xlsx.styles['generic2'])
    xlsx.write(offset + 60, 2, 'Boat MSRP = C / .61 / 0.7',
               xlsx.styles['generic1'])
    xlsx.write(offset + 61, 2, 'Options MSRP = C / .8046 / .48',
               xlsx.styles['generic1'])
    xlsx.write(offset + 62, 2, 'Trailers MSRP = C / 0.80 / 0.7',
               xlsx.styles['generic1'])
    xlsx.write(offset + 63, 2, 'Inboard Motors MSRP = C / 0.85 / 0.7',
               xlsx.styles['generic1'])
    xlsx.write(offset + 64, 2,
               'Big Ticket Items MSRP = C / (range from 0.80 – 0.85) / 0.7',
               xlsx.styles['generic1'])

    xlsx.write(offset + 78, 2,
               'Cost estimate check list - complete prior to sending quote '
               'or submitting bid',
               xlsx.styles['generic2'])
    xlsx.write(offset + 79, 2,
               'Verify all formulas are correct and all items are included '
               'in cost total',
               xlsx.styles['generic1'])
    xlsx.write(offset + 80, 2,
               'Verify aluminum calculated with total lbs included. Include '
               'metal costing sheet separate if completed',
               xlsx.styles['generic1'])
    xlsx.write(offset + 81, 2,
               'Verify paint costing equals paint description',
               xlsx.styles['generic1'])
    xlsx.write(offset + 82, 2, 'Cost estimate includes all components on '
               'sales quote', xlsx.styles['generic1'])
    xlsx.write(offset + 83, 2,
               'Pricing policy discounts and minimum margins are met',
               xlsx.styles['generic1'])
    xlsx.write(offset + 84, 2,
               'Vendor quotes received and included in costing folder',
               xlsx.styles['generic1'])
    xlsx.write(offset + 85, 2,
               'Labor hours reviewed and correct to best knowledge of project',
               xlsx.styles['generic1'])
    xlsx.write(offset + 86, 2,
               'Name of peer who reviewed prior to submission to customer',
               xlsx.styles['generic1'])
    xlsx.write(offset + 87, 2,
               'Customer signed sales quotation',
               xlsx.styles['generic1'])
    xlsx.write(offset + 88, 2,
               'Customer provided terms and conditions, including payment '
               'schedule',
               xlsx.styles['generic1'])

    # COLUMN C ================================================================
    xlsx.write(offset + 0, 3, 'Fabrication',
               xlsx.styles['generic1'])
    xlsx.write(offset + 1, 3, 'Fab Consumables',
               xlsx.styles['generic1'])
    xlsx.write(offset + 2, 3, 'Paint',
               xlsx.styles['generic1'])
    xlsx.write(offset + 3, 3, 'Paint Consumables',
               xlsx.styles['generic1'])
    xlsx.write(offset + 4, 3, 'Outfitting',
               xlsx.styles['generic1'])
    xlsx.write(offset + 5, 3, 'Big Ticket Items',
               xlsx.styles['generic1'])
    xlsx.write(offset + 6, 3, 'OB Motors',
               xlsx.styles['generic1'])
    xlsx.write(offset + 7, 3, 'IB Motors & Jets',
               xlsx.styles['generic1'])
    xlsx.write(offset + 8, 3, 'Trailer',
               xlsx.styles['generic1'])

    xlsx.write(offset + 14, 3, 'Fabrication ',
               xlsx.styles['generic1'])
    xlsx.write(offset + 15, 3, 'Paint',
               xlsx.styles['generic1'])
    xlsx.write(offset + 16, 3, 'Outfitting',
               xlsx.styles['generic1'])
    xlsx.write(offset + 17, 3, 'Design / Drafting',
               xlsx.styles['generic1'])
    xlsx.write(offset + 20, 3, None, xlsx.styles['bgYellow3'])

    xlsx.write(offset + 25, 3, 'Test Fuel',
               xlsx.styles['generic1'])
    xlsx.write(offset + 26, 3, 'Trials',
               xlsx.styles['generic1'])
    xlsx.write(offset + 27, 3, 'Engineering',
               xlsx.styles['generic1'])

    xlsx.write(offset + 34, 3, 'Trucking',
               xlsx.styles['generic1'])
    xlsx.write(offset + 35, 3, 'Dealer commission',
               xlsx.styles['generic1'])


    xlsx.write(offset + 0, 42, 'Cost',
               xlsx.styles['generic2'])


    # COLUMN D ================================================================
    xlsx.write(offset + 19, 4, 'Total Hours',
               xlsx.styles['rightJust1'])
    # COLUMN E ================================================================
    # COLUMN F ================================================================
    # COLUMN G ================================================================
    # COLUMN H ================================================================
    # COLUMN I ================================================================
    # COLUMN J ================================================================


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
    # caculate the size of each section

    # create new workbook / xlsx file
    with Workbook(file_name_info['file_name']) as workbook:
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
        generate_totals(xlsx, SECTION_TEST)
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
