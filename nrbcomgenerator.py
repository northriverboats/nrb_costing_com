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
from typing import List

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
              mailhost = os.environ.get("MAIL_SERVER"),
              fromaddr = os.environ.get("MAIL_FROM"),
              toaddrs = os.environ.get("MAIL_TO"),
              subject = "alert!"
            )
smtpHandler.setLevel(logging.CRITICAL)
smtpHandler.setFormatter(formatter)

logger.addHandler(consoleHandler)
logger.addHandler(fileHandler)
logger.addHandler(smtpHandler)


@dataclass
class BoatModels:
  sheet1: str
  sheet2: str
  folder: str


@click.command()
def main() -> None:
  try:
    pass
  except Exception as e:
    logger.critical(traceback.format_exc())
    raise
  finally:
    # program terminates normally
    sys.exit()

if __name__ == "__main__":
  main()