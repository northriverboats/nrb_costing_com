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
from typing import List, Optional, Union, cast

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
class BoatModels:
    """Information on which sheets make a boat costing sheet"""
    sheet1: str
    sheet2: str
    folder: str

@dataclass(order=True)
class Resources:
    """Parts Information """
    oempart: str
    description: str = field(compare=False)
    unitprice: float = field(compare=False)
    oem: str = field(compare=False)
    vendorpart: str = field(compare=False)
    vendor: str = field(compare=False)
    updated: datetime.datetime = field(compare=False)

@dataclass
class Consumables:
    """Consumables rate by department"""
    dept: str
    percent: float

@dataclass
class HourlyRates:
    """Hourly rate by department"""
    dept: str
    rate: float

@dataclass
class MarkUps:
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
    try:
        index = section.parts.index(part)
        section.parts[index].qty += part.qty
    except ValueError:
        section.parts.append(part)

def load_boat_models(master_file: Path) -> List[BoatModels]:
    """Build master list of sheets to combine to create costing sheets"""
    try:
        xlsx = openpyxl.load_workbook(master_file.as_posix(), data_only=True)

        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        dimensions: str = sheet.dimensions  # type: ignore
        end_cell: str = dimensions.split(':')[1]
        rnge = 'A2:' + end_cell
        cells = sheet[rnge]
        boats: List[BoatModels] = [
            cast(BoatModels, [v.value for v in cell]) for cell in cells if cell[0].value]
        xlsx.close()
    except (FileNotFoundError, PermissionError):
        return list()
    return boats

def load_resource(resource_file: Path) -> List[Resources]:
    """Read resource sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        dimensions: str = sheet.dimensions  # type: ignore
        end_cell: str = dimensions.split(':')[1]
        rnge = 'A2:' + end_cell
        cells = sheet[rnge]
        resources: List[Resources] = [
            cast(Resources, [v.value for v in cell]) for cell in cells if cell[0].value]
    except (FileNotFoundError, PermissionError):
        return list()
    finally:
        if xlsx:
            xlsx.close()
    return resources

def load_resources(resource_files: List[Path]) -> List[Resources]:
    """Load all resource files"""
    resources: List[Resources] = list()
    for resource_file in resource_files:
        click.echo('.', nl=False)
        resources += load_resource(resource_file)
    click.echo('')
    return resources

def find_excel_files_in_dir(base: Union[str, Path]) -> List[Path]:
    """get list of spreadsheets in folder"""
    if isinstance(base, str):
        base = Path(base)
    return [sheets for sheets in base.glob('[!~]*.xlsx')]  # pylint: disable=unnecessary-comprehension

def load_consumables(resource_file: Path) -> List[Consumables]:
    """Read consuables sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)

        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        dimensions: str = sheet.dimensions  # type: ignore
        end_cell: str = dimensions.split(':')[1]
        rnge = 'A1:' + end_cell
        cells = sheet[rnge]
        consumables: List[Consumables] = [
            cast(Consumables, [v.value for v in cell]) for cell in cells if cell[0].value]
    except (FileNotFoundError, PermissionError):
        return list()
    else:
        if xlsx:
            xlsx.close()
    return consumables

def load_hourly_rates(resource_file: Path) -> List[HourlyRates]:
    """Read hourly rates sheet"""
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)

        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        dimensions: str = sheet.dimensions  # type: ignore
        end_cell: str = dimensions.split(':')[1]
        rnge = 'A1:' + end_cell
        cells = sheet[rnge]
        hourly_rates: List[HourlyRates] = [
            cast(HourlyRates, [v.value for v in cell]) for cell in cells if cell[0].value]
    except (FileNotFoundError, PermissionError):
        return list()
    else:
        if xlsx:
            xlsx.close()
    return hourly_rates

def load_mark_ups(resource_file: Path) -> List[MarkUps]:
    """read makrkup file into object """
    try:
        xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)

        sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
        dimensions: str = sheet.dimensions  # type: ignore
        end_cell: str = dimensions.split(':')[1]
        rnge: str = 'A2:' + end_cell
        cells = sheet[rnge]
        mark_ups: List[MarkUps] = [
            cast(MarkUps, [v.value for v in cell]) for cell in cells if cell[0].value]
    except (FileNotFoundError, PermissionError):
        return list()
    else:
        if xlsx:
            xlsx.close()
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
        models: List[BoatModels] = load_boat_models(Path(MASTER_FILE))
        boat_files: List[Path] = find_excel_files_in_dir(Path(BOATS_FOLDER))
        resource_files: List[Path] = [
            sheet
            for sheet in find_excel_files_in_dir(Path(RESOURCES_FOLDER))
            if sheet.name.startswith('BOM ')]
        resources: List[Resources] = load_resources(resource_files)
        consumables: List[Consumables] = load_consumables(  # pylint: disable=unused-variable
            Path(RESOURCES_FOLDER).joinpath('Consumables.xlsx'))
        hourly_rates = load_hourly_rates(  # pylint: disable=unused-variable
            Path(RESOURCES_FOLDER).joinpath('HOURLY RATES.xlsx'))
        mark_ups = load_mark_ups(Path(RESOURCES_FOLDER).joinpath('Mark up.xlsx'))  # pylint: disable=unused-variable
        click.echo(f'Models: {len(models)}   ', nl=False)
        click.echo(f'Boat Files: {len(boat_files)}   ', nl=False)
        click.echo(f'Resource Files: {len(resource_files)}   ', nl=False)
        click.echo(f'Resources: {len(resources)}   ')
        click.echo(pprint.pformat(resources, width=210))
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()
