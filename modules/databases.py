#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4 tabstop=8
"""
save/restore data from sqlite3

Pass in master_file return data structure
"""
from sqlite3 import connect, Connection, Cursor
from pathlib import Path
from typing import Any, Optional, Union
from .boms import Boms
from .models import Models
from .resources import Resources
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


def create_schema(cursor: Cursor) -> None:
    """create needed tables if necessary

    Arguments:
        cursor: Cursor -- active database cursor

    Raises:
        OperationalError
    Returns:
        None
    """
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS storage (
            name  varchar(50) primary key not null,
            value jsonify)""")

def serialize(cursor: Cursor, name: str, value: Any) -> None:
    """convert object to json and save in database

    Arguements:
        cursor: Cursor -- active database cursor
        name: str -- name of object to save
        value: Any -- object that has a to_json() method

    Returns:
        None
    """
    cursor.execute("""
                   INSERT OR REPLACE INTO storage(name, value)
                   VALUES(?, ?)""", (name, value.to_json()))


def deserialized(cursor: Cursor, name: str) -> Any:
    """convert json from database into object

    Arguements:
        cursor: Cursor -- active database cursor
        name: str -- name of object to save

    Returns:
        any -- desearilized object
    """
    cursor.execute("""SELECT value from storage WHERE name = ?""", [name])
    row = cursor.fetchone()
    return row[0] if row else ""



def load_from_database(db_file: Path)-> tuple[Models, Resources, Boms]:
    """read data from database into objects"""
    status_msg("Reading Data from {str(db_file.resolve())}", 1)
    with dbopen(db_file) as cursor:
        # pylint: disable=no-member
        models = Models.from_json(deserialized(cursor, 'models'))
        resources = Resources.from_json(deserialized(cursor, 'models'))
        boms = Boms.from_json(deserialized(cursor, 'models'))
    print()
    return models, resources, boms

def save_to_database(db_file: Path,
                     models: Models,
                     resources: Resources,
                     boms: Boms)-> None:
    """jsonify and save save objects to database

    Arguments:
        db_file: Path -- name of file to save objects to

    Raise:

    Return:
        None
    """
    status_msg("Saving Data to {str(db_file.resolve())}", 1)
    with dbopen(db_file) as cursor:
        create_schema(cursor)
        serialize(cursor, 'models', models)
        serialize(cursor, 'resources', resources)
        serialize(cursor, 'boms', boms)

if __name__ == "__main__":
    pass
