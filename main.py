#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
NRB COMMERCIAL COSTING SHEET GENERATOR

The files in the RESOURCES folder are labeled BOM but in reality is just a list
of parts from each department/motor/trailer that can go on a BOM

The files in the folder BOATS are both boat files and cabin files but are the
real BOM files

resources = parts that can be used on a bom
boms = a list of parts used in building a cabin or boat

DEBUG to log file
ERROR to log file and screen
CRITICAL to log file, screen and email
"""
import logging
import logging.handlers
import os
# import pprint
import sys
import traceback
from pathlib import Path
import click
from dotenv import load_dotenv  # pylint: disable=import-error
from models import load_models, Model
from resources import load_resources, Resource
from utility import enable_logging, options, resource_path, status_msg


env_path = resource_path('.env')
load_dotenv(dotenv_path=env_path)
logger = logging.getLogger(__name__)

DATABASE: Path = Path(os.environ.get('DATABASE', ''))
SHEETS_FOLDER: Path = Path(os.environ.get('SHEETS_FOLDER', ''))
BOATS_FOLDER: Path = Path(os.environ.get('BOATS_FOLDER', ''))
RESOURCES_FOLDER: Path = Path(os.environ.get('RESOURCES_FOLDER', ''))
TEMPLATE_FILE: Path = Path(os.environ.get('TEMPLATE_FILE', ''))
MODELS_FILE: Path = Path(os.environ.get('MODELS_FILE', ''))
CONSUMABLES_FILE: Path = Path(os.environ.get('CONSUMABLES_FILE', ''))
HOURLY_RATES_FILE: Path = Path(os.environ.get('HOURLY_RATES_FILE', ''))
MARK_UPS_FILE: Path = Path(os.environ.get('MARK_UPS_FILE', ''))
MAIL_SERVER: str = str(os.environ.get("MAIL_SERVER", ''))
MAIL_FROM: str = str(os.environ.get("MAIL_FROM", ''))
MAIL_TO: str = str(os.environ.get("MAIL_TO", ''))


#
# ==================== Main Entry Point
#
@click.command()
@click.option('-v', '--verbose', count=True)
def main(verbose: int) -> None:
    """ main program entry point """
    options['verbose'] = verbose
    enable_logging(logger, MAIL_SERVER, MAIL_FROM, MAIL_TO)
    try:
        # load information from spreadsheets
        models: list[Model] = load_models(MODELS_FILE)
        resources: dict[str, Resource] = load_resources(RESOURCES_FOLDER)

        # display stats about spreadsheets
        status_msg(f"{len(models)} models loaded", 0)
        status_msg(f"{len(resources)} resources loaded", 0)
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()  # pylint: disable=E1120