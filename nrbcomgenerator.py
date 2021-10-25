#!/usr/bin/env python
"""
NRB COMMERCIAL COSTING SHEET GENERATOR

DEBUG to log file
ERROR to log file and screen
CRITICAL to log file, screen and email
"""
import click
import datetime
import logging
import logging.handlers
import openpyxl
import os
import pprint
import sys
import traceback
from dataclasses import dataclass
from dotenv import load_dotenv
from pathlib import Path
from typing import List, Optional, Union, cast


"""
==================== Low Level Utilities
"""
def resource_path(relative_path: Union[str, Path]) -> Path:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
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


"""
==================== ENALBE LOGGING
DEBUG + = to rotating log files in current directory
INFO + = to stdout
CRITICAL + = to email
"""
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

consoleHandler = logging.StreamHandler(sys.stdout)
consoleHandler.setLevel(logging.DEBUG)
consoleHandler.setFormatter(formatter)

fileHandler = logging.handlers.RotatingFileHandler(filename="error.log",maxBytes=1024000, backupCount=10, mode="a")
fileHandler.setLevel(logging.INFO)
fileHandler.setFormatter(formatter)

smtpHandler = logging.handlers.SMTPHandler(
              mailhost = MAIL_SERVER,
              fromaddr = MAIL_FROM,
              toaddrs = MAIL_TO,
              subject = "alert!"
            )
smtpHandler.setLevel(logging.CRITICAL)
smtpHandler.setFormatter(formatter)

logger.addHandler(consoleHandler)
logger.addHandler(fileHandler)
logger.addHandler(smtpHandler)


"""
==================== Dataclasses
"""
@dataclass
class BoatModels:
  sheet1: str
  sheet2: str
  folder: str

@dataclass
class Resources:
  oempart: str
  description: str
  unitprice: float
  oem: str
  vendorpart: str
  vendor: str
  updated: datetime.datetime

@dataclass
class Consumables:
  dept: str
  percent: float

@dataclass
class Hourly_Rates:
  dept: str
  rate: float

@dataclass
class Mark_Ups:
  policy: str
  markup_1: float
  markup_2: float
  discount: float


"""
==================== Low Level Functions
"""
def load_boat_models(master_file: Path) -> List[BoatModels]:
  try:
    xlsx = openpyxl.load_workbook(master_file.as_posix(), data_only=True)
  except (FileNotFoundError, PermissionError):
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A2:' + end_cell
  cells = sheet[range]
  boats: List[BoatModels] = [cast(BoatModels,[v.value for v in cell]) for cell in cells if cell[0].value]
  xlsx.close()
  return boats

def load_resource(resource_file: Path) -> List[Resources]:
  try:
    xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
  except (FileNotFoundError, PermissionError):
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A2:' + end_cell
  cells = sheet[range]
  resources: List[Resources] = [cast(Resources ,[v.value for v in cell]) for cell in cells if cell[0].value]
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
  return [sheet for sheet in base.glob('[!~]*.xlsx')]

def load_consumables(resource_file: Path) -> List[Consumables]:
  try:
    xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
  except (FileNotFoundError, PermissionError):
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A1:' + end_cell
  cells = sheet[range]
  consumables: List[Consumables] = [cast(Consumables ,[v.value for v in cell]) for cell in cells if cell[0].value]
  xlsx.close()
  return consumables

def load_hourly_rates(resource_file: Path) -> List[Hourly_Rates]:
  try:
    xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
  except (FileNotFoundError, PermissionError):
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A1:' + end_cell
  cells = sheet[range]
  hourly_rates: List[Hourly_Rates] = [cast(Hourly_Rates ,[v.value for v in cell]) for cell in cells if cell[0].value]
  xlsx.close()
  return hourly_rates

def load_mark_ups(resource_file: Path) -> List[Mark_Ups]:
  try:
    xlsx = openpyxl.load_workbook(resource_file.as_posix(), data_only=True)
  except (FileNotFoundError, PermissionError):
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A2:' + end_cell
  cells = sheet[range]
  mark_ups: List[Mark_Ups] = [cast(Mark_Ups ,[v.value for v in cell]) for cell in cells if cell[0].value]
  xlsx.close()
  return mark_ups


"""
==================== High Level Functions
"""


"""
==================== Main Entry Point
"""
@click.command()
def main() -> None:
  try:
    models: List[BoatModels] = load_boat_models(Path(MASTER_FILE))
    boat_files: List[Path] = find_excel_files_in_dir(Path(BOATS_FOLDER))
    resource_files: List[Path] = [sheet for sheet in find_excel_files_in_dir(Path(RESOURCES_FOLDER)) if sheet.name.startswith('BOM ')]
    resources: List[Resources] = load_resources(resource_files)
    consumables: List[Consumables] = load_consumables(Path(RESOURCES_FOLDER).joinpath('Consumables.xlsx'))
    hourly_rates = load_hourly_rates(Path(RESOURCES_FOLDER).joinpath('HOURLY RATES.xlsx'))
    mark_ups = load_mark_ups(Path(RESOURCES_FOLDER).joinpath('Mark up.xlsx'))
    click.echo(f'Models: {len(models)}   Boat Files: {len(boat_files)}   Resource Files: {len(resource_files)}   ', nl=False)
    click.echo(f'Resources {len(resources)}')
    click.echo(pprint.pformat(resources, width=210))
  except Exception as e:
    logger.critical(traceback.format_exc())
    raise
  finally:
    # program terminates normally
    sys.exit()

if __name__ == "__main__":
  main()