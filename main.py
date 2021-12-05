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
from typing import Union
import click
from modules.boms import load_boms, Boms
from modules.consumables import load_consumables, Consumables
from modules.costingsheets import generate_sheets_for_all_models
from modules.databases import load_from_database, save_to_database
from modules.hourlyrates import load_hourly_rates, HourlyRates
from modules.markups import MarkUp, load_mark_ups, MarkUps
from modules.models import load_models, Models
from modules.resources import load_resources, Resources
from modules.settings import Settings
from modules.utilities import (enable_logging, logger, options, status_msg,
                               BOATS_FOLDER, CONSUMABLES_FILE, DATABASE,
                               HOURLY_RATES_FILE, MARK_UPS_FILE, MODELS_FILE,
                               MAIL_SERVER, MAIL_FROM, MAIL_TO,
                               RESOURCES_FOLDER,)

#
# ==================== Main Entry Point
#
@click.command()
@click.option('-b', '--buld', 'build_only', is_flag=True,
              help="Build database only do not create sheets")
@click.option('-l', '--load', 'load_file', is_flag=False,
              flag_value="DATABASE",
              default="", help="Load data from sqlite database")
@click.option('-s', '--save', 'save_file', is_flag=False,
              flag_value="DATABASE",
              default="", help="Save data to sqlite database")
@click.option('-v', '--verbose', count=True,
              help="Increase verbosity")
def main(build_only: bool,
         load_file: Union[Path, str],
         save_file: Union[Path, str],
         verbose: int) -> None:
    """ main program entry point """
    if build_only:
        load_file = ""
        save_file = DATABASE
    options['verbose'] = verbose
    enable_logging(logger, MAIL_SERVER, MAIL_FROM, MAIL_TO)
    boms: Boms
    consumables: Consumables
    hourly_rates: HourlyRates
    mark_ups: MarkUps
    models: Models
    resources: Resources
    settings: Settings
    if load_file == "DATABASE":
        load_file = DATABASE
    if save_file == "DATABASE":
        save_file = DATABASE
    try:
        if load_file:
            json = load_from_database(load_file, [
                'boms',
                'consumables',
                'hourly_rates',
                'mark_ups',
                'models',
                'resources',
                ])
            boms = Boms.from_json(json['boms'])
            consumables = Consumables.from_json(json['consumables'])
            hourly_rates = HourlyRates.from_json(json['hourly_rates'])
            mark_ups = MarkUps.from_json(json['mark_ups'])
            models = Models.from_json(json['models'])
            resources = Resources.from_json(json['resources'])
        else:
            # load information from spreadsheets
            models = load_models(MODELS_FILE)
            status_msg(f"{len(models.models)} models loaded", 0)

            # resources is only needed to build BomPart
            resources = load_resources(RESOURCES_FOLDER)
            status_msg(f"{len(resources.resources)} resources loaded", 0)

            # build BOM information
            boms = load_boms(BOATS_FOLDER, resources.resources)
            status_msg(f"{len(boms.boms)} boms loaded", 0)

            # build Consumables information
            consumables = load_consumables(CONSUMABLES_FILE)
            status_msg(f"{len(consumables.consumables)} consumables loaded", 0)

            # build Hourly Rates information
            hourly_rates = load_hourly_rates(HOURLY_RATES_FILE)
            status_msg(
                f"{len(hourly_rates.hourly_rates)} hourly rates loaded", 0)

            # build BOM information
            mark_ups = load_mark_ups(MARK_UPS_FILE)
            status_msg(f"{len(mark_ups.mark_ups)} mark ups loaded", 0)

        settings = Settings(consumables.consumables, hourly_rates.hourly_rates, mark_ups.mark_ups)
        if not build_only:
            generate_sheets_for_all_models(models.models, boms.boms, settings)
        if save_file:
            save_to_database(save_file, {
                'boms':  boms.to_json(),
                'consumables': consumables.to_json(),
                'hourly_rates': hourly_rates.to_json(),
                'mark_ups': mark_ups.to_json(),
                'models': models.to_json(),
                'resources': resources.to_json(),
            })
    except Exception:
        logger.critical(traceback.format_exc())
        raise
    finally:
        # program terminates normally
        sys.exit()

if __name__ == "__main__":
    main()  # pylint: disable=E1120
