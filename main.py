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
import sys
import traceback
import click
from modules.boms import load_boms, Bom
from modules.consumables import load_consumables, Consumable
from modules.costingsheets import generate_sheets_for_all_models
from modules.hourlyrates import load_hourly_rates, HourlyRate
from modules.markups import load_mark_ups, MarkUp
from modules.models import load_models, Model
from modules.resources import load_resources, Resource
from modules.utilities import (enable_logging, logger, options, status_msg,
                               BOATS_FOLDER, CONSUMABLES_FILE, MARK_UPS_FILE,
                               HOURLY_RATES_FILE, MODELS_FILE, MAIL_SERVER,
                               MAIL_FROM, MAIL_TO, RESOURCES_FOLDER,)

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
        if 1 > 2:
            # should not need to be used if TEMPLATE_FILE is updated
            consumables: dict[str, Consumable] = load_consumables(
                CONSUMABLES_FILE)
            status_msg(f"{len(consumables)} consumalbes loaded\n", 0)
            hourly_rates: dict[str, HourlyRate] = load_hourly_rates(
                HOURLY_RATES_FILE)
            status_msg(f"{len(hourly_rates)} hourly rates loaded\n", 0)
            mark_ups: dict[str, MarkUp] = load_mark_ups(MARK_UPS_FILE)
            status_msg(f"{len(mark_ups)} mark ups loaded\n", 0)

        # load information from spreadsheets
        models: dict[str, Model] = load_models(MODELS_FILE)
        status_msg(f"{len(models)} models loaded\n", 0)

        # resources is only needed to build BomPart
        resources: dict[str, Resource] = load_resources(RESOURCES_FOLDER)
        status_msg(f"{len(resources)} resources loaded\n", 0)

        # build BOM information
        boms: dict[str, Bom] = load_boms(BOATS_FOLDER, resources)
        status_msg(f"{len(boms)} boms loaded\n", 0)

        generate_sheets_for_all_models(models, boms)
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()  # pylint: disable=E1120
