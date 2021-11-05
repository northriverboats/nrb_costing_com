#!/usr/bin/env python
"""
NRB COMMERCIAL COSTING SHEET GENERATOR

DEBUG to log file
ERROR to log file and screen
CRITICAL to log file, screen and email
"""
import datetime
import logging
import logging.handlers
import os
import pprint
import sys
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Union

import click
from dotenv import load_dotenv  # pylint: disable=import-error
import openpyxl  # pylint: disable=import-error

#
# ==================== Low Level Utilities
#
def resource_path(relative_path: Union[str, Path]) -> Path:
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # pylint: disable=protected-access
        base_path = Path(sys._MEIPASS)  # type: ignore
    except AttributeError:
        base_path = Path.cwd()

    return base_path / relative_path

env_path = resource_path('.env')
load_dotenv(dotenv_path=env_path)

DATABASE: str = str(os.environ.get('DATABASE'))
COSTING_FOLDER: str = str(os.environ.get('COSTING_FOLDER'))
SHEETS_FOLDER: str = str(os.environ.get('SHEETS_FOLDER'))
TEMPLATE_FILE: str = str(os.environ.get('TEMPLATE_FILE'))
MASTER_FILE: str = str(os.environ.get('MASTER_FILE'))
BOATS_FOLDER: str = str(os.environ.get('BOATS_FOLDER'))
RESOURCES_FOLDER: str = str(os.environ.get('RESOURCES_FOLDER'))
MAIL_SERVER: str = str(os.environ.get("MAIL_SERVER"))
MAIL_FROM: str = str(os.environ.get("MAIL_FROM"))
MAIL_TO: str = str(os.environ.get("MAIL_TO"))


#
# ==================== ENALBE LOGGING
# DEBUG + = to rotating log files in current directory
# INFO + = to stdout
# CRITICAL + = to email

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

consoleHandler = logging.StreamHandler(sys.stdout)
consoleHandler.setLevel(logging.DEBUG)
consoleHandler.setFormatter(formatter)

fileHandler = logging.handlers.RotatingFileHandler(
    filename="error.log",
    maxBytes=1024000,
    backupCount=10,
    mode="a"
)
fileHandler.setLevel(logging.INFO)
fileHandler.setFormatter(formatter)

smtpHandler = logging.handlers.SMTPHandler(
    mailhost=MAIL_SERVER,
    fromaddr=MAIL_FROM,
    toaddrs=MAIL_TO,
    subject="alert!"
)
smtpHandler.setLevel(logging.CRITICAL)
smtpHandler.setFormatter(formatter)

logger.addHandler(consoleHandler)
logger.addHandler(fileHandler)
logger.addHandler(smtpHandler)


#
# ==================== Custom Errors
#
class NRBError(Exception):
    """Base class for all NRB errors"""
    def __init__(self, *args):
        if args:
            self.message = args[0]
        else:
            self.message = None
        super().__init__(self)

    def __str__(self):
        if self.message:
            return 'NRB Error, {0} '.format(self.message)
        return 'NRB Error has been raised'


#
# ==================== Dataclasses
#
@dataclass
class BoatModel:
    """Information on which sheets make a boat costing sheet"""
    sheet1: str
    sheet2: str
    folder: str

@dataclass(order=True)
class Resource:
    """BOM Part Information """
    # pylint: disable=too-many-instance-attributes
    # Eight is reasonable in this case.
    oempart: str
    description: str = field(compare=False)
    uom: str = field(compare=False)
    unitprice: float = field(compare=False)
    oem: str = field(compare=False)
    vendorpart: str = field(compare=False)
    vendor: str = field(compare=False)
    updated: datetime.datetime = field(compare=False)

@dataclass
class Consumable:
    """Consumables rate by department"""
    dept: str
    percent: float

@dataclass
class HourlyRate:
    """Hourly rate by department"""
    dept: str
    rate: float

@dataclass
class MarkUp:
    """Mark-up rates by deprtment"""
    policy: str
    markup_1: float
    markup_2: float
    discount: float

@dataclass(order=True)
class BomPart:
    """Part Information from Section of a BOM Parts Sheet"""
    part: Optional[str]
    qty: float = field(compare=False)
    smallest: Optional[float] = field(compare=False)
    biggest: Optional[float] = field(compare=False)
    percent: Optional[float] = field(compare=False)  # FT field

@dataclass(order=True)
class BomSection:
    """Group of BOM Parts"""
    name: str
    parts: List[BomPart] = field(compare=False)

@dataclass(order=True)
class BomSheet:
    """BOM sheet"""
    name: str
    smallest: float = field(compare=False)
    biggest: float = field(compare=False)
    sections: List[BomSection] = field(compare=False)

#
# ==================== Low Level Functions
#

def make_bom_part(row) -> BomPart:
    """Create bom part from row in spreadsheet"""
    qty: float = float(row[0].value)
    smallest: Optional[float] = None if row[1].value is None else float(row[1].value)
    biggest: Optional[float] = None if row[2].value is None else float(row[2].value)
    percent: Optional[float] = None if row[3].value is None else float(row[3].value)
    part: Optional[str] = None if row[5].value is None else row[5].value
    return BomPart(part, qty, smallest, biggest, percent)

def section_add_part(section: BomSection, part: BomPart) -> None:
    """Insert new part or add qty to existing part"""
    found = [item for item in section.parts if item.part == part.part]
    if found:
        found[0].qty += part.qty
    else:
        section.parts.append(part)

def load_boat_models(master_file: Path) -> List[BoatModel]:
    """Build master list of sheets to combine to create costing sheets"""
    try:
        xlsx = openpyxl.load_workbook(master_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        boats: List[BoatModel] = list()
        for row in sheet.iter_rows(min_row=2, max_col=3):
            if not isinstance(row[0].value, str):
                continue
            boat: BoatModel = BoatModel(
                row[0].value,
                row[1].value,
                row[2].value)
            boats.append(boat)
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        pass
    return boats

def load_resource_file(resource_file: Path) -> List[Resource]:
    """Read resource sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        resources: List[Resource] = list()
        for row in sheet.iter_rows(min_row=2, max_col=8):
            if not isinstance(row[0].value, str):
                continue
            resource: Resource = Resource(
                row[0].value,
                row[1].value,
                row[2].value,
                float(row[3].value),
                row[4].value,
                row[5].value,
                row[6].value,
                row[7].value)
            resources.append(resource)
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        pass
    return resources

def load_resources(resource_files: List[Path]) -> List[Resource]:
    """Load all resource files"""
    resources: List[Resource] = list()
    for resource_file in resource_files:
        resources += load_resource_file(resource_file)
    click.echo('')
    return resources

def find_excel_files_in_dir(base: Union[str, Path]) -> List[Path]:
    """get list of spreadsheets in folder"""
    if isinstance(base, str):
        base = Path(base)
    return [sheets for sheets in base.glob('[!~]*.xlsx')]  # pylint: disable=unnecessary-comprehension

def load_consumables(resource_file: Path) -> List[Consumable]:
    """Read consuables sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        consumables: List[Consumable] = list()
        for row in sheet.iter_rows(min_row=2, max_col=2):
            if not isinstance(row[0].value, str):
                continue
            consumable: Consumable = Consumable(
                row[0].value,
                float(row[1].value))
            consumables.append(consumable)
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        pass
    return consumables

def load_hourly_rates(resource_file: Path) -> List[HourlyRate]:
    """Read hourly rates sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        hourly_rates: List[HourlyRate] = list()
        for row in sheet.iter_rows(min_row=2, max_col=2):
            if not isinstance(row[0].value, str):
                continue
            hourly_rate: HourlyRate = HourlyRate(
                row[0].value,
                float(row[1].value))
            hourly_rates.append(hourly_rate)
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        pass
    return hourly_rates

def load_mark_ups(resource_file: Path) -> List[MarkUp]:
    """read makrkup file into object """
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        mark_ups: List[MarkUp] = list()
        for row in sheet.iter_rows(min_row=2, max_col=4):
            if not isinstance(row[0].value, str):
                continue
            mark_up: MarkUp = MarkUp(
                row[0].value,
                float(row[1].value),
                float(row[2].value),
                float(row[3].value))
            mark_ups.append(mark_up)
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        pass
    return mark_ups

def get_hull_sizes(sheet: openpyxl.worksheet.worksheet.Worksheet) -> List:
    """find all hull sizes listed in sheet"""
    dimensions: str = sheet.dimensions  # type: ignore
    rnge: str = 'M1:' + ''.join(
        [c for c in dimensions.split(':')[1] if c not in "0123456789"]) + "1"
    cells = sheet[rnge]
    sizes = [cell.value for cell in cells[0] if cell.value]
    if "ANY" in sizes:
        sizes = ["ANY"] + [size for size in sizes if size != "ANY"]
    return sizes


# ==================== High Level Functions


# ==================== Main Entry Point

@click.command()
def main() -> None:
    """ main program entry point """
    try:
        models: List[BoatModel] = load_boat_models(Path(MASTER_FILE))
        boat_files: List[Path] = find_excel_files_in_dir(Path(BOATS_FOLDER))
        resource_files: List[Path] = [
            sheet
            for sheet in find_excel_files_in_dir(Path(RESOURCES_FOLDER))
            if sheet.name.startswith('BOM ')]
        resources: List[Resource] = load_resources(resource_files)
        consumables: List[Consumable] = load_consumables(  # pylint: disable=unused-variable
            Path(RESOURCES_FOLDER).joinpath('Consumables.xlsx'))
        hourly_rates: List[HourlyRate] = load_hourly_rates(  # pylint: disable=unused-variable
            Path(RESOURCES_FOLDER).joinpath('HOURLY RATES.xlsx'))
        mark_ups: List[MarkUp] = load_mark_ups(Path(RESOURCES_FOLDER).joinpath('Mark up.xlsx'))  # pylint: disable=unused-variable
        click.echo(f'Models: {len(models)}   ', nl=False)
        click.echo(f'Boat Files: {len(boat_files)}   ', nl=False)
        click.echo(f'Resource Files: {len(resource_files)}   ', nl=False)
        click.echo(f'Resources: {len(resources)}   ', nl=False)
        click.echo(f'Consumables: {len(consumables)}   ', nl=False)
        click.echo(f'Hourly Rates: {len(hourly_rates)}   ', nl=False)
        click.echo(f'Mark Ups: {len(mark_ups)}   ')
        click.echo()
        # click.echo(pprint.pformat(resources, width=210))
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()
