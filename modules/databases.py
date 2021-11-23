#!/usr/bin/env python
# vim expandtab shiftwidth=2 softtabstop=2
"""
save/restore data from sqlite3

Pass in master_file return data structure
"""
from pathlib import Path
from .utilities import status_msg

def load_from_database(db_file: Path)-> None:
    """read data from database into objects"""
    status_msg(str(db_file.resolve()), 99)
    status_msg("Loading Data....", 0)

def save_to_database(db_file: Path)-> None:
    """save data to database from objects"""
    status_msg(str(db_file.resolve()), 99)
    status_msg("Saving Data....", 0)

if __name__ == "__main__":
    pass
