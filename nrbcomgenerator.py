#!/usr/bin/env python
"""
NRB COMMERCIAL COSTING SHEET GENERATOR

DEBUG to log file
ERROR to log file and screen
CRITICAL to log file, screen and email
"""
import click
import logging
import logging.handlers
import openpyxl
import os
import sys
import traceback
from dataclasses import dataclass
from dotenv import load_dotenv
from pathlib import Path
from typing import List, Optional, Union


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
              mailhost = str(MAIL_SERVER),
              fromaddr = str(MAIL_FROM),
              toaddrs = str(MAIL_TO),
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


"""
==================== Low Level Functions
"""
def load_boat_models(master_file: Path) -> List[BoatModels]:
  try:
    xlsx = openpyxl.load_workbook(master_file.as_posix())
  except FileNotFoundError:
    return list()

  sheet: openpyxl.worksheet.worksheet.Worksheet = xlsx.active
  dimensions: str = sheet.dimensions  # type: ignore
  end_cell: str = dimensions.split(':')[1]
  range = 'A2:' + end_cell
  cells = sheet[range]
  boats: List[BoatModels] = [BoatModels(cell[0].value, cell[1].value, cell[2].value) for cell in cells]
  xlsx.close()
  return boats

def find_excel_files_in_dir(base: Union[str, Path]) -> List[Path]:
  """get list of spreadsheets in folder"""
  if isinstance(base, str):
    base = Path(base)
  # print(base)
  return [sheet for sheet in base.glob('[!~]*.xlsx')]


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
    boats: List[Path] = find_excel_files_in_dir(Path(BOATS_FOLDER))
    resources: List[Path] = find_excel_files_in_dir(Path(RESOURCES_FOLDER))
    print(f'Models: {len(models)}   Boats: {len(boats)}   Resources: {len(resources)}')
  except Exception as e:
    logger.critical(traceback.format_exc())
    raise
  finally:
    # program terminates normally
    sys.exit()

if __name__ == "__main__":
  main()