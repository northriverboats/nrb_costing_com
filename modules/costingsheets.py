#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import date
from typing import Any, Optional, TypedDict
from pathlib import Path
from xlsxwriter import Workbook # type: ignore
# from xlsxwriter.utility import xl_rowcol_to_cell # type: ignore
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
    totals: float = 0

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
                                             col.width * (100.8/127 + .00077),
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


# pylint: disable=anomalous-backslash-in-string
STYLES = [
     Format('generic1', {'font_name': 'arial',
                         'font_size': 10,
                        }),
     Format('generic2', {'font_name': 'Arial',
                         'font_size': 10,
                         'bold': True,
                        }),

     Format('currency', {'font_name': 'arial',
                                 'font_size': 10,
                                 'num_format': '_("$"* #,##0.00_);_("$"*'
                                 '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                        }),

     Format('currencyYellow', {'pattern':1,
                               'bg_color': '#FCF305',
                               'font_name': 'arial',
                               'font_size': 10,
                               'num_format': '_("$"* #,##0.00_);_("$"*'
                               '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                              }),

     Format('currencyBold', {'font_name': 'arial',
                                 'font_size': 10,
                                 'num_format': '_("$"* #,##0.00_);_("$"*'
                                 '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                                 'bold': True,
                            }),

     Format('currencyBoldYellow', {'pattern':1,
                                   'bg_color': '#FCF305',
                                   'font_name': 'arial',
                                   'font_size': 10,
                                   'num_format': '_("$"* #,##0.00_);_("$"*'
                                   '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                                   'bold': True,
                                  }),

     Format('currencyBoldYellowBorder', {'pattern':1,
                                         'bg_color': '#FCF305',
                                         'font_name': 'arial',
                                         'font_size': 10,
                                         'num_format': '_("$"* #,##0.00_);_("$"*'
                                         '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                                         'bold': True,
                                         'border': 2,
                                       }),

     Format('percentBorderYellow', {'pattern':1,
                                    'bg_color': '#FCF305',
                                    'font_name': 'arial',
                                    'font_size': 10,
                                    'num_format': '0.00%',
                                    'border': True,
                                   }),

     Format('currencyBordered', {'font_name': 'arial',
                                 'font_size': 10,
                                 'num_format': '_("$"* #,##0.00_);_("$"*'
                                 '\(#,##0.00\);_("$"* "-"??_);_(@_)',
                                 'border':1,
                                }),

     Format('percent', {'font_name': 'arial',
                        'font_size': 10,
                        'num_format': '0%',
                       }),

     Format('percentBorder', {'font_name': 'arial',
                              'font_size': 10,
                              'num_format': '0%',
                              'border': 1,
                              'align': 'center',
                             }),

    Format('headingCustomer1', {'font_name': 'Arial',
                                'font_size': 18,
                                'bold': True,
                               }),

    Format('headingCustomer2', {'font_name': 'Arial',
                                'font_size': 20,
                                'bold': True,
                                'pattern': 1,
                                'bg_color': '#FCF305',
                                'bottom': 1,
                               }),

    Format('bgSilver', {'pattern':1,
                        'bg_color': 'silver',
                        'align': 'center',
                        'bold': True,
                       }),

    Format('bgYellow0', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'align': 'center',
                        }),
    Format('bgYellow1', {'pattern': 1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                        }),
    Format('bgYellow2', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',
                        }),
    Format('bgYellow3', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'border': 1,
                         'align': 'center',
                        }),
    Format('bgYellow4', {'pattern':1,
                         'bg_color': '#FCF305',
                         'font_name': 'Arial',
                         'font_size': 10,
                        }),

    Format('bgGreen1', {'pattern':1,
                        'bg_color': '#1FB714',
                        'font_name': 'Arial',
                        'font_size': 10,
                        'bottom': 1,
                       }),
    Format('bgGreen2', {'pattern':1,
                        'bg_color': '#1FB714',
                        'font_name': 'Arial',
                        'font_size': 10,
                        'bottom': 1,
                        'align': 'center',
                       }),
    Format('bgGreen3', {'pattern':1,
                        'bg_color': '#92D050',
                        'font_name': 'Arial',
                        'font_size': 10,
                       }),

    Format('bgPurple1', {'pattern':1,
                         'bg_color': '#CC99FF',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                        }),
    Format('bgPurple2', {'pattern':1,
                         'bg_color': '#CC99FF',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',
                        }),

    Format('bgCyan1', {'pattern': 1,
                       'bg_color': '#99CCFF',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,
                      }),
    Format('bgCyan2', {'pattern':1,
                       'bg_color': '#99CCFF',
                       'font_name': 'Arial',
                       'font_size': 10,
                       'bottom': 1,
                       'align': 'center',
                      }),

    Format('bgOrange1', {'pattern':1,
                         'bg_color': 'FF9900',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                        }),

    Format('bgOrange2', {'pattern':1,
                         'bg_color': 'FF9900',
                         'font_name': 'Arial',
                         'font_size': 10,
                         'bottom': 1,
                         'align': 'center',
                        }),

    Format('rightJust1', {'align': 'right',
                          'font_name': 'Arial',
                          'font_size': 10,
                         }),
    Format('rightJust2', {'align': 'right',
                          'font_name': 'Arial',
                          'font_size': 10,
                          'bold': True
                         }),

    Format('centerJust1', {'align': 'center',
                           'font_name': 'Arial',
                           'font_size': 10,
                           'bold': True
                          }),

    Format('centerJust2', {'align': 'center',
                           'font_name': 'Arial',
                           'font_size': 10,
                          }),
    Format('centerJust3', {'align': 'center',
                           'font_name': 'Arial',
                           'font_size': 10,
                           'bold': True,
                           'text_wrap': True,
                          }),
    Format('centerJust4', {'align': 'center',
                           'font_name': 'Arial',
                           'font_size': 6,
                           'bold': True,
                           'text_wrap': True,
                          }),
    Format('centerJust5', {'align': 'center',
                           'font_name': 'Arial',
                           'font_size': 10,
                           'border': 1,
                          }),

    Format('bgSilverBorder', {'pattern':1,
                              'bg_color': '#BFBFBF',
                              'font_name': 'Arial',
                              'font_size': 10,
                              'border': 1,
                             }),
    Format('bgSilverBorderCetner', {'pattern':1,
                                    'bg_color': '#BFBFBF',
                                    'font_name': 'Arial',
                                    'font_size': 10,
                                    'border': 1,
                                    'align': 'center',
                                   }),
]

SECTION_TEST = {
    'FABRICATION': SectionInfo(20, 28, 29,10),
    'PAINT': SectionInfo(30, 38, 39, 12),
    'UNUSED': SectionInfo(40, 48, 49, 0),
    'OUTFITTING': SectionInfo(50, 58, 59,16),
    'BIG TICKET ITEMS': SectionInfo(60, 68, 69, 18),
    'OUTBOARD MOTORS': SectionInfo(70, 78, 79, 20),
    'INBOARD MOTORS & JETS': SectionInfo(80, 88, 89, 22),
    'TRAILER': SectionInfo(90, 98, 99, 24),
    'TOTALS': SectionInfo(100, 150, 152, 0),
}

DEALERS = [
    'Clemens',
    '3Rivers',
    'Y Marina',
    'Boat Country',
    'AFF',
    'Bay Co.',
    'PBH',
    'Erie Marine',
    'Valley',
]

YESNO = [
    'Yes',
    'No',
]


SALESPERSON = [
    'Mike',
    'Jordan',
    'Jesse',
    'Jay',
    'Brent'
]

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
                     xlsx.styles['bgGreen2']
                    )
    xlsx.merge_range('E6:G6',
                     'Indicate changes here',
                     xlsx.styles['bgPurple2']
                    )

    xlsx.merge_range('E7:G7',
                     'Indicate changes here',
                     xlsx.styles['bgCyan2']
                    )
    xlsx.merge_range('E8:G8',
                     'Indicate changes here',
                     xlsx.styles['bgOrange2']
                    )

    xlsx.write('H5', 'Rev1', xlsx.styles['rightJust1'])
    xlsx.write('H6', 'Rev2', xlsx.styles['rightJust1'])
    xlsx.write('H7', 'Rev3', xlsx.styles['rightJust1'])
    xlsx.write('H8', 'Rev4', xlsx.styles['rightJust1'])

    xlsx.write('I5', None, xlsx.styles['bgGreen1'])
    xlsx.write('I6', None, xlsx.styles['bgPurple1'])
    xlsx.write('I7', None, xlsx.styles['bgCyan1'])
    xlsx.write('I8', None, xlsx.styles['bgOrange1'])


def generate_totals_b(xlsx: Xlsx,
                      section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet column b"""
    offset = section_info['TRAILER'].subtotal + 2

    # COLUMN B ================================================================
    for row in range(offset + 79, offset + 86):
        xlsx.write(row, 1, None, xlsx.styles['bgYellow4'])
    for row in range(offset + 86, offset + 89):
        xlsx.write(row, 1, None, xlsx.styles['bgGreen3'])
    xlsx.sheet.data_validation(offset + 79, 1, offset + 85, 1, {
        'validate': 'list',
        'source': YESNO,
    })
    xlsx.sheet.data_validation(offset + 86, 1, offset + 86, 1, {
        'validate': 'list',
        'source': SALESPERSON,
    })
    xlsx.sheet.data_validation(offset + 87, 1, offset + 88, 1, {
        'validate': 'list',
        'source': YESNO,
    })

def totals_00(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 00 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['FABRICATION'].subtotal)
    value1 = section_info['FABRICATION'].value

    xlsx.write(row, 2, 'MATERIALS', xlsx.styles['rightJust2'])
    xlsx.write(row, 3, 'Fabrication', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_01(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 01 from the bottom of the sheet"""
    formula1 = "=I" + str(row) + '*H' + str(row + 1)
    value1 = section_info['FABRICATION'].value * 0.08

    xlsx.write(row, 3, 'Fab Consumables', xlsx.styles['generic1'])
    xlsx.write(row, 7, 0.08, xlsx.styles['percent'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_02(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 02 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['PAINT'].subtotal)
    value1 = section_info['PAINT'].value

    xlsx.write(row, 3, 'Paint', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_03(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 03 from the bottom of the sheet"""
    formula1 = "=I" + str(row) + '*H' + str(row+1)
    value1 = section_info['PAINT'].value * 0.50

    xlsx.write(row, 3, 'Paint Consumables', xlsx.styles['generic1'])
    xlsx.write(row, 7, 0.50, xlsx.styles['percent'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_04(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 04 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['OUTFITTING'].subtotal)
    value1 = section_info['OUTFITTING'].value

    xlsx.write(row, 3, 'Outfitting', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_05(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 05 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['BIG TICKET ITEMS'].subtotal)
    value1 = section_info['BIG TICKET ITEMS'].value

    xlsx.write(row, 3, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_06(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 06 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['OUTBOARD MOTORS'].subtotal)
    value1 = section_info['OUTBOARD MOTORS'].value

    xlsx.write(row, 3, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_07(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 07 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['INBOARD MOTORS & JETS'].subtotal)
    value1 = section_info['INBOARD MOTORS & JETS'].value

    xlsx.write(row, 3, 'IB Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_08(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 08 from the bottom of the sheet"""
    formula1 = "=I" + str(section_info['TRAILER'].subtotal)
    value1 = section_info['TRAILER'].value

    xlsx.write(row, 3, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_09(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 09 from the bottom of the sheet"""
    formula1 = "=SUM(I" + str(row - 8) +":I" + str(row)
    # value used in totals_40
    value1 = (sum([section_info[section].value for section in section_info]) +
                   section_info['FABRICATION'].value * 0.08 +
                   section_info['PAINT'].value * 0.50)

    xlsx.write(row, 7, 'Total All Materials', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_12(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 12 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'Labor', xlsx.styles['rightJust2'])

def totals_13(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 13 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 5, 'BOAT HOURS', xlsx.styles['centerJust1'])
    xlsx.write(row, 6, 'TOTAL HOURS', xlsx.styles['centerJust1'])
    xlsx.write(row, 7, 'RATE', xlsx.styles['centerJust1'])

def totals_14(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 14 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=F" + str(row + 1) + "+SUM(L:L)"
    value1 = 0
    formula2 = "=H" +str(row +1 ) +  "*G" + str(row + 1)
    value2 = 0

    xlsx.write(row, 3, 'Fabrication', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 59.22, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_15(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 15 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=F" + str(row + 1) + "+SUM(M:M)"
    value1 = 0
    formula2 = "=H" +str(row +1 ) +  "*G" + str(row + 1)
    value2 = 0

    xlsx.write(row, 3, 'Paint', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 59.22, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_16(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 16 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=F" + str(row + 1) + "+SUM(N:N)"
    value1 = 0
    formula2 = "=H" +str(row +1 ) +  "*G" + str(row + 1)
    value2 = 0

    xlsx.write(row, 3, 'Outfitting', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 59.22, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_17(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 17 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=F" + str(row + 1) + "+SUM(O:O)"
    value1 = 0
    formula2 = "=H" +str(row +1 ) +  "*G" + str(row + 1)
    value2 = 0

    xlsx.write(row, 3, 'Design / Drafting', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 59.22, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_19(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 19 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=SUM(F" + str(row - 4) + ":F" + str(row - 1)
    value1 = 0
    formula2 = "=SUM(I" + str(row - 4) + ":I" + str(row - 1)
    value2 = 0

    xlsx.write(row, 4, 'Total Hours', xlsx.styles['rightJust1'])
    xlsx.write(row, 5, formula1, xlsx.styles['bgYellow0'], value1)
    xlsx.write(row, 7, 'Total Labor Costs', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula2, xlsx.styles['currencyBoldYellow'], value2)

def totals_20(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 20 from the bottom of the sheet"""
    _ = section_info
    text1 = 'Indicate boat referenced for labor hours if used'

    xlsx.write(row, 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(row, 3, None, xlsx.styles['bgYellow4'])

def totals_23(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 23 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'Other Costs', xlsx.styles['rightJust2'])

def totals_25(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 25 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Test Fuel', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_26(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 26 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Trials', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_27(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 27 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Engineering', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_28(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 28 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_29(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 29 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_30(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 30 from the bottom of the sheet"""
    _ = section_info

    formula1 = "=SUM(I" + str(row - 5) + ":I" + str(row)
    value1 = 0

    xlsx.write(row, 7, 'Total Other Costs', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_32(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 32 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'NO MARGIN ITEMS', xlsx.styles['rightJust2'])

def totals_34(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 34 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Trucking', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_35(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 35 from the bottom of the sheet"""
    _ = section_info
    text1 = "Voyager/Custom - 10%, Guide/Lodge - 3%"

    xlsx.write(row, 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(row, 3, 'Dealer commission', xlsx.styles['generic1'])
    xlsx.write(row, 5, 'Dealer', xlsx.styles['centerJust2'])
    xlsx.write(row, 6, None, xlsx.styles['bgYellow4'])
    xlsx.sheet.data_validation(row, 6, row, 6, {
        'validate': 'list',
        'source': DEALERS,
    })
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_36(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 36 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_37(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 37 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_38(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 38 from the bottom of the sheet"""
    _ = section_info

    formula1 = "=SUM(I" + str(row - 3) + ":I" + str(row)
    value1 = 0

    xlsx.write(row, 7, 'Total No Margin Items', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_40(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 40 from the bottom of the sheet"""
    formula1 = ("=I" + str(row - 30) +
               "+I" + str(row - 20) +
               "+I" + str(row - 9) +
               "+I" + str(row - 1))
    # value used in totals_40
    value1 = (sum([section_info[section].value for section in section_info]) +
                section_info['FABRICATION'].value * 0.08 +
                section_info['PAINT'].value * 0.50)

    xlsx.write(row, 6, 'TOTAL COST OF PROJECT', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)


def totals_42(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 42 from the bottom of the sheet"""
    _ = section_info
    text1 = "Mark up per pricing policy: "

    xlsx.sheet.set_row(row, 23.85)
    xlsx.write(row, 2, text1, xlsx.styles['generic2'])
    xlsx.write(row, 3, 'Cost ', xlsx.styles['centerJust1'])
    xlsx.merge_range(row, 4, row, 5, 'Markup',  xlsx.styles['centerJust3'])
    xlsx.write(row, 6, 'MSRP ', xlsx.styles['centerJust3'])
    xlsx.write(row, 7, 'Discount ', xlsx.styles['centerJust3'])
    xlsx.write(row, 9, 'Contribution Margin', xlsx.styles['centerJust4'])

def totals_43(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 43 from the bottom of the sheet"""
    formula1 = ("=I" + str(row - 2) +
                "-I" + str(row - 4) +
                "-I" + str(row - 37) +
                "-I" + str(row - 36) +
                "-I" + str(row - 35) +
                "-I" + str(row - 34))
    # value used in totals_40
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * 0.08 +
              section_info['PAINT'].value +
              section_info['PAINT'].value * 0.50 +
              section_info['OUTFITTING'].value)
    formula2 = "=D" + str(row + 1)  + "/E" + str(row + 1) + "/F" + str(row + 1)
    value2 = value1 / 0.61 / 0.7
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = ("=IF(I" + str(row + 1) +
                "=0,0,(I" + str(row + 1) +
                "-D" + str(row + 1) +
                ")/I" + str(row + 1) +
                ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals = value3

    xlsx.write(row, 2, 'Boat and options:', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.61, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_44(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 44 from the bottom of the sheet"""
    formula1 = "=I" + str(row - 38)
    value1 = section_info['BIG TICKET ITEMS'].value
    formula2 = ("=D" + str(row + 1) +
                "/E" + str(row + 1) +
                "/F" + str(row + 1))
    value2 = value1 / 0.8 / 0.85
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = ("=IF(I" + str(row + 1) +
                "=0,0,(I" + str(row + 1) +
                "-D" + str(row + 1) +
                ")/I" + str(row + 1) +
                ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.8, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.85, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_45(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 45 from the bottom of the sheet"""
    formula1 = "=I" + str(row - 38)
    value1 = section_info['OUTBOARD MOTORS'].value
    formula2 = "=Q" + str(section_info['OUTBOARD MOTORS'].subtotal)
    value2 = 0.0
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = ("=IF(I" + str(row + 1) + "=0,0,(I" + str(row +1) + "-D" +
                str(row +1)  + ")/I" + str(row + 1) + ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(row, 4, row, 5, 'See PP',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_46(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 46 from the bottom of the sheet"""
    formula1 = "=I" + str(row - 38)
    value1 = section_info['INBOARD MOTORS & JETS'].value
    formula2 = ("=D" + str(row + 1) +
                "/E" + str(row + 1) +
                "/F" + str(row + 1))
    value2 = value1 / 0.85 / 0.7
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = ("=IF(I" + str(row +1) +
                "=0,0,(I" + str(row +1) +
                "-D" + str(row +1) +
                ")/I" + str(row +1) +
                ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Inboard Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.85, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_47(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 47 from the bottom of the sheet"""
    formula1 = "=I" + str(row - 38)
    value1 = section_info['TRAILER'].value
    formula2 = ("=D" + str(row + 1) +
                "/E" + str(row + 1) +
                "/F" + str(row + 1))
    value2 = value1 / 0.8 / 0.7
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = ("=IF(I" + str(row + 1) +
                "=0,0,(I" + str(row + 1) +
                "-D" + str(row + 1) +
                ")/I" + str(row + 1) +
                ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.8, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_48(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 48 from the bottom of the sheet"""
    _ = section_info
    formula1 = "=I" + str(row - 9)
    value1 = 0.0
    formula2 = "=D" + str(row + 1)
    value2 = 0.0
    formula3 = "=G" + str(row + 1)
    value3 = value2
    formula4 = ("=IF(I" + str(row + 1) +
                "=0,0,(I" + str(row + 1) +
                "-D" + str(row + 1) +
                ")/I" + str(row + 1) +
                ")")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'No margin items: ', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(row, 4, row, 5, 'none',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, 'none', xlsx.styles['centerJust5'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_50(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 50 from the bottom of the sheet"""
    text1 = "Total Cost (equals total cost of project box)"
    formula1 = "=SUM(D" + str(row - 6) + ":D" + str(row -1) + ")"
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * 0.08 +
              section_info['PAINT'].value +
              section_info['PAINT'].value * 0.50 +
              section_info['OUTFITTING'].value +
              section_info['BIG TICKET ITEMS'].value +
              section_info['OUTBOARD MOTORS'].value +
              section_info['INBOARD MOTORS & JETS'].value +
              section_info['TRAILER'].value)
    formula2 = "=SUM(I" + str(row - 6) + ":I" + str(row -1) + ")"
    value2 = section_info['TOTALS'].totals

    xlsx.write(row, 2, text1, xlsx.styles['rightJust2'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyYellow'], value1)
    xlsx.write(row, 7, 'Calculated Selling Price', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_52(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 52 from the bottom of the sheet"""

    xlsx.write(row, 6, 'SELLING PRICE', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, None, xlsx.styles['currencyBoldYellowBorder'])

def generate_totals(xlsx: Xlsx, section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet"""
    # START OF TEMP STUFF
    for section in section_info:
        if section in ['UNUSED', 'TOTALS']:
            continue
        xlsx.write(section_info[section].subtotal - 1, 8,
                   section_info[section].value,
                   xlsx.styles['currency'],
                  )
    xlsx.sheet.set_top_left_cell("A144")
    xlsx.write(section_info['OUTBOARD MOTORS'].subtotal -1, 16, 30.0)
    # END OF TEMP STUFF
    skip = [10, 11, 18, 21, 22, 24, 31, 33,
            39, 41, 49, 51, 53, 55, 57, 58, 77]
    # pylint: disable=unused-variable
    offset = section_info['TRAILER'].subtotal + 2
    for row in range(0, 53):
        if row in skip:
            continue
        # pylint: disable=eval-used
        eval(f"totals_{row:02}(xlsx, section_info, row + offset)")


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
