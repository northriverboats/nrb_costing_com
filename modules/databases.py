#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4 tabstop=8
"""
save/restore data from sqlite3

Pass in master_file return data structure
"""
from sqlite3 import connect, Connection, Cursor
from pathlib import Path
from typing import Optional, Union
from .utilities import status_msg

class dbopen():  # pylint: disable=invalid-name
    """
    Simple context manager for sqlite3 databases. Commits everything at exit.
    """
    def __init__(self, path: Union[Path, str]) -> None:
        self.path: Union[Path,str] = path
        self.conn: Optional[Connection] = None
        self.cursor: Optional[Cursor] = None

    def __enter__(self) -> Cursor:
        self.conn = connect(self.path)
        self.cursor = self.conn.cursor()
        return self.cursor

    def __exit__(self, exc_class, exc, traceback):
        self.conn.commit()
        self.conn.close()

def load_from_database(db_file: Path)-> None:
    """read data from database into objects"""
    status_msg(str(db_file.resolve()), 99)
    status_msg("Loading Data....", 0)

def save_to_database(db_file: Path)-> None:
    """jsonify and save save objects to database

    Arguments:
        db_file: Path -- name of file to save objects to

    Raise:
        N/A

    Return:
        None
    """
    status_msg(str(db_file.resolve()), 99)
    status_msg("Saving Data....", 0)
    with dbopen(db_file) as cursor:
        print(cursor)

if __name__ == "__main__":
    pass
