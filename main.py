#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
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
import sys
import traceback
from pathlib import Path
import click
from modules.boms import load_boms, Boms
from modules.costingsheets import generate_sheets_for_all_models
from modules.models import load_models, Models
from modules.resources import load_resources, Resources
from modules.utilities import (enable_logging, logger, options, status_msg,
                               BOATS_FOLDER, DATABASE, MODELS_FILE,
                               MAIL_SERVER, MAIL_FROM, MAIL_TO,
                               RESOURCES_FOLDER,)

#
# ==================== Main Entry Point
#
@click.command()
@click.option('-l', '--load', is_flag=False, flag_value="DATABASE",
              default="", help="Load data from sqlite database")
@click.option('-s', '--save', is_flag=False, flag_value="DATABASE",
              default="", help="Save data to sqlite database")
@click.option('-v', '--verbose', count=True,
              help="Increase verbosity")
def main(load: str, save: str, verbose: int) -> None:
    """ main program entry point """
    options['verbose'] = verbose
    enable_logging(logger, MAIL_SERVER, MAIL_FROM, MAIL_TO)
    load_file: Path = Path()
    save_file: Path = Path()
    if load == "DATABASE":
        load = str(DATABASE.resolve())
    if save == "DATABASE":
        save = str(DATABASE.resolve())
    try:
        if load:
            load_file = Path(load)
            print(f"future home of loading: {load_file.resolve()}")
        else:
            # load information from spreadsheets
            all_models: Models = load_models(MODELS_FILE)
            status_msg(f"{len(all_models.models)} models loaded", 0)

            # resources is only needed to build BomPart
            all_resources: Resources = load_resources(RESOURCES_FOLDER)
            status_msg(f"{len(all_resources.resources)} resources loaded", 0)

            # build BOM information
            all_boms: Boms = load_boms(BOATS_FOLDER, all_resources.resources)
            status_msg(f"{len(all_boms.boms)} boms loaded", 0)
        # generate_sheets_for_all_models(all_models.models, all_boms.boms)
        if save:
            save_file = Path(save)
            print(f"future home of saving: {save_file.resolve()}")
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()  # pylint: disable=E1120
