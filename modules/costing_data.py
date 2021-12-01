#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Generate Costing Sheets
"""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, TypedDict
from .boms import Bom
from .hourlyrates import HourlyRate

# DATA CLASSES ================================================================
class FileNameInfo(TypedDict):
    """file name info"""
    size: str
    model: str
    option: str
    folder: str
    with_options: str
    size_with_options: str
    size_with_folder: str
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
    # pylint: disable=too-many-instance-attributes
    workbook: Any
    bom: Bom
    size: str
    hourly_rates: dict[str, HourlyRate]
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
                self.sheet.set_column_pixels(
                    col.columns,
                    col.width * (100.8/100.8 + .00077),
                    self.styles[col.style])
            else:
                self.sheet.set_column_pixels(col.columns, col.width)

    def load_formats(self, styles):
        """ load styles from a list of  Format objects"""
        # can parse styles.style and extract and do other processing
        for style in styles:
            self.add_format(style.name, style.style)


# SHEET DATA ==================================================================
COLUMNS = [                             # PIXELS   POINTS
    Columns('A:A', 126.50, 'generic1'), # 126.50   100.80
    Columns('B:B', 132, 'generic1'),    # 132.00   104.75
    Columns('C:C', 314.00, 'generic1'), # 314.00   249.20
    Columns('D:D', 106.50, 'generic1'), # 106.50    84.95
    Columns('E:E', 34, 'generic1'),     #  34.00    27.00
    Columns('F:F', 92, 'generic1'),     #  92.00    73.00
    Columns('G:G', 90, 'generic1'),     #  90.00    71.45
    Columns('H:H', 69, 'generic1'),     #  69.00    54.75
    Columns('I:I', 119, 'generic1'),    # 119.00    94.47
    Columns('J:P', 62, 'generic1'),     #  62.00    49.20
    Columns('Q:Q', 90, 'generic1'),     #  90.00    71.45
    Columns('R:T', 62, 'generic1'),     #  62.00    49.20
]


CURRENCY = (
    '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'
)

# pylint: disable=anomalous-backslash-in-string
# pylint: disable=anomalous-backslash-in-string
STYLES = [
    Format(
        'generic1',
        {
            'font_name': 'arial',
            'font_size': 10,
        },
    ),
    Format(
        'generic2',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
        },
    ),
    Format(
        'thickBottom',
        {
            'font_name': 'arial',
            'font_size': 10,
            'bottom': 2,
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
        'currencyYellow',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
        },
    ),
    Format(
        'currencyBold',
        {
            'font_name': 'arial',
            'font_size': 10,
            'bold': True,
            'num_format': CURRENCY,
        },
    ),
    Format(
        'currencyBoldYellow',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'bold': True,
        },
    ),
    Format(
        'currencyYellowBorder',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'border': 1,
        },
    ),
    Format(
        'currencyBoldYellowBorder',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'bold': True,
            'border': 2,
        },
    ),
    Format(
        'percentBorderYellow',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0.00%',
            'border': 1,
        },
    ),
    Format(
        'decimal',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0.00',
        },
    ),
    Format(
        'decimalBordered',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0.00',
            'border': 1,
        },
    ),
    Format(
        'bgYellowDecimal',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0.00',
        },
    ),
    Format(
        'currencyBordered',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'border': 1,
        },
    ),
    Format(
        'currencyBorderedBold',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'border': 1,
            'bold': True,
        },
    ),
    Format(
        'normalBordered',
        {
            'font_name': 'arial',
            'font_size': 10,
            'border': 1,
        },
    ),
    Format(
        'bgGreenCurrencyBordered',
        {
            'pattern': 1,
            'bg_color': '#1FB714',
            'font_name': 'arial',
            'font_size': 10,
            'num_format': CURRENCY,
            'border': 1,
        },
    ),
    Format(
        'bgGreenNormalBordered',
        {
            'pattern': 1,
            'bg_color': '#1FB714',
            'font_name': 'arial',
            'font_size': 10,
            'border': 1,
        },
    ),
    Format(
        'percent',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0%',
        },
    ),
    Format(
        'percent1',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0.00%',
        },
    ),
    Format(
        'percentBorder',
        {
            'font_name': 'arial',
            'font_size': 10,
            'num_format': '0%',
            'border': 1,
            'align': 'center',
        },
    ),
    Format(
        'headingCustomer1',
        {
            'font_name': 'Arial',
            'font_size': 18,
            'bold': True,
        },
    ),
    Format(
        'headingCustomer2',
        {
            'font_name': 'Arial',
            'font_size': 20,
            'bold': True,
            'pattern': 1,
            'bg_color': '#FCF305',
            'bottom': 1,
        },
    ),
    Format(
        'bgSilverBold12pt',
        {
            'pattern': 1,
            'bg_color': 'silver',
            'font_name': 'Arial',
            'font_size': 12,
            'align': 'center',
            'bold': True,
        },
    ),
    Format(
        'bgSilverBold10pt',
        {
            'pattern': 1,
            'bg_color': 'silver',
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
        },
    ),
    Format(
        'bgSilver',
        {
            'pattern': 1,
            'bg_color': 'silver',
            'align': 'center',
            'bold': True,
        },
    ),
    Format(
        'bgSilverBorderedCurrency',
        {
            'pattern': 1,
            'bg_color': 'silver',
            'num_format': CURRENCY,
            'bold': True,
            'border': 1,
        },
    ),
    Format(
        'bgYellow0',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'center',
        },
    ),
    Format(
        'bgYellow1',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
        },
    ),
    Format(
        'bgYellow2',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
            'align': 'center',
        },
    ),
    Format(
        'bgYellow3',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
            'align': 'center',
        },
    ),
    Format(
        'bgYellow4',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
        },
    ),
    Format(
        'bgYellowRight',
        {
            'pattern': 1,
            'bg_color': '#FCF305',
            'font_name': 'Arial',
            'font_size': 10,
            'align': 'right',
        },
    ),
    Format(
        'bgGreen1',
        {
            'pattern': 1,
            'bg_color': '#1FB714',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
        },
    ),
    Format(
        'bgGreen2',
        {
            'pattern': 1,
            'bg_color': '#1FB714',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
            'align': 'center',
        },
    ),
    Format(
        'bgGreen3',
        {
            'pattern': 1,
            'bg_color': '#92D050',
            'font_name': 'Arial',
            'font_size': 10,
        },
    ),
    Format(
        'bgPurple1',
        {
            'pattern': 1,
            'bg_color': '#CC99FF',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
        },
    ),
    Format(
        'bgPurple2',
        {
            'pattern': 1,
            'bg_color': '#CC99FF',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
            'align': 'center',
        },
    ),
    Format(
        'bgCyan1',
        {
            'pattern': 1,
            'bg_color': '#99CCFF',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
        },
    ),
    Format(
        'bgCyan2',
        {
            'pattern': 1,
            'bg_color': '#99CCFF',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
            'align': 'center',
        },
    ),
    Format(
        'bgOrange1',
        {
            'pattern': 1,
            'bg_color': 'FF9900',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
        },
    ),
    Format(
        'bgOrange2',
        {
            'pattern': 1,
            'bg_color': 'FF9900',
            'font_name': 'Arial',
            'font_size': 10,
            'bottom': 1,
            'align': 'center',
        },
    ),
    Format(
        'rightJust1',
        {
            'align': 'right',
            'font_name': 'Arial',
            'font_size': 10,
        },
    ),
    Format(
        'rightJust2',
        {
            'align': 'right',
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
        },
    ),
    Format(
        'centerJust1',
        {
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
        },
    ),
    Format(
        'centerJust2',
        {
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 10,
        },
    ),
    Format(
        'centerJust3',
        {
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'text_wrap': True,
        },
    ),
    Format(
        'centerJust4',
        {
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 6,
            'bold': True,
            'text_wrap': True,
        },
    ),
    Format(
        'centerJust5',
        {
            'align': 'center',
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
        },
    ),
    Format(
        'bgSilverBorder',
        {
            'pattern': 1,
            'bg_color': '#BFBFBF',
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
        },
    ),
    Format(
        'bgSilverBorderCetner',
        {
            'pattern': 1,
            'bg_color': '#BFBFBF',
            'font_name': 'Arial',
            'font_size': 10,
            'border': 1,
            'align': 'center',
        },
    ),
    Format(
        'italicsNote',
        {
            'font_name': 'Arial',
            'font_size': 8,
            'italic': True,
            'text_wrap': True,
            'valign': 'top',
        },
    ),
    Format(
        'red',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'font_color': '#FF0000',
            'bold': True,
        },
    ),
    Format(
        'redMerged',
        {
            'font_name': 'Arial',
            'font_size': 12,
            'font_color': '#FF0000',
            'bold': True,
            'underline': True,
            'text_wrap': True,
        }
    ),
    Format(
        'heading1',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'align': 'center',
            'text_wrap': True,
        },
    ),
    Format(
        'heading2',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'bold': True,
            'text_wrap': True,
        },
    ),
    Format(
        'updated',
        {
            'font_name': 'Arial',
            'font_size': 10,
            'num_format': 'mm/dd/yy',
            'border': 1,
        },
    ),
    Format(
        'bgGreenUpdated',
        {
            'pattern': 1,
            'bg_color': '#1FB714',
            'font_name': 'Arial',
            'font_size': 10,
            'num_format': 'mm/dd/yy',
            'border': 1,
        },
    ),
]
SECTIONS = {
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

if __name__ == "__main__":
    pass
