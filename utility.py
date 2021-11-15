#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
Utility Functions
"""
import sys
from pathlib import Path
from typing import Union
import logging
import logging.handlers
from click import echo

options: dict = {'verbose': 0}

def resource_path(relative_path: Union[str, Path]) -> Path:
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # pylint: disable=protected-access
        base_path = Path(sys._MEIPASS)  # type: ignore
    except AttributeError:
        base_path = Path.cwd()

    return base_path / relative_path


def noop() -> None:
    """Empty funtcion placeholder"""


class NRBError(Exception):
    """Base class for all NRB errors
    Use as NRBError("My custom message")
    """
    def __init__(self, *args):
        if args:
            self.message = args[0]
        else:
            self.message = None
        super().__init__(self)

    def __str__(self):
        if self.message:
            return f"NRB Error, {self.message}"
        return 'NRB Error has been raised'

class NRBErrorNotFound(NRBError):
    """Error if item is not found during lookup"""
    def __str__(self):
        if self.message:
            return f"{self.message}"
        return 'NRB Not Found Error has been raised'


def status_msg(msg: str, level: int, nl: bool = True) -> None:
    """output message if verbosity is sufficent"""
    if options['verbose'] >= level:
        echo(msg, nl=nl)

def normalize_size(size: float) -> str:
    """convert float to proper feet inchs"""
    if size > int(size):
        return f"{int(size)}' 6\""
    return f"{int(size)}'"


#
# ==================== ENALBE LOGGING
# DEBUG + = to stdout
# INFO + = to rotating log files in current directory
# CRITICAL + = to email
def enable_logging(logger: logging.Logger,
                   mail_server: str,
                   mail_from: str,
                   mail_to: str) -> None:
    """enable logging for app"""
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)

    file_handler = logging.handlers.RotatingFileHandler(
        filename="error.log",
        maxBytes=1024000,
        backupCount=10,
        mode="a"
    )
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    smtp_handler = logging.handlers.SMTPHandler(
        mailhost=mail_server,
        fromaddr=mail_from,
        toaddrs=mail_to,
        subject="NRB Commerical Costing Sheet Generator Critial Error"
    )
    smtp_handler.setLevel(logging.CRITICAL)
    smtp_handler.setFormatter(formatter)

    logger.addHandler(console_handler)
    logger.addHandler(file_handler)
    # logger.addHandler(smtp_handler)

if __name__ == "__main__":
    pass
