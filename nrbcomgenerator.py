#!/usr/bin/env python
"""
NRB COMMERCIAL COSTING SHEET GENERATOR
"""

import click
import logging
import logging.handlers
import os
import sys
from dotenv import load_dotenv
from functools import wraps

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

def exception(logger):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except:
                issue = "exception in "+func.__name__+"\n"
                issue = issue+"=============\n"
                # logger.exception(issue)
                logger.critical(issue, exc_info=True)

                raise
        return wrapper
    return decorator


@exception(logger)
def main():
  logger.critical("Situation is critical. Come to office immediately.")
  pass

if __name__ == "__main__":
  main()