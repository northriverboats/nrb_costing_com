#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4 tabstop=8
"""
Save/restore dataclasses from sqlite3

Will save/restore a dictionary of name/json_representation of objects

Can be called multiple times. There is less overhead with fewer calls and
larger dictionaries
"""
from sqlite3 import connect, Connection, Cursor
from pathlib import Path
from typing import Optional, Union
from .utilities import status_msg

# LOW LEVEL FUNCTIONS =========================================================
def file_message(message: str, file_name: Union[Path, str]) -> None:
    """output file related message

    Arguments:
        message: str -- text with {file_name} slot in it
        file_name: Path|str  --  path object or text of file name

    Returns:
        None
    """
    text = (file_name
            if isinstance(file_name, str)
            else str(file_name.resolve()))
    status_msg(message.format(file_name=text), 1)


class dbopen():  # pylint: disable=invalid-name
    """
    Simple context manager for sqlite3 databases. Connenction will
    automatically close when exiting the contxet managers scope, even if the
    exit happens due to a raised exception. Use as:

        with dbopen(fileanme) as db:
            # any commands or function calls here can ues db
       # anything here will happen after db has been closed

    Returns:
        cursor -- databse cursor
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

def serialize(cursor: Cursor, objects: dict[str, str]) -> None:
    """convert object to json and save in database

    Arguements:
        cursor: Cursor -- active database cursor
        objects: dict[str, str] -- name_of_object, json_representation

    Returns:
        None
    """
    for name in objects:
        cursor.execute("""
                       INSERT OR REPLACE INTO storage(name, value)
                       VALUES(?, ?)""", (name, objects[name]))

def deserialized(cursor: Cursor, names: list[str]) -> dict[str, str]:
    """convert json from database into object

    Arguements:
        cursor: Cursor -- active database cursor
        name: str -- name of object to fetch json_representation

    Returns:
        objects: dict[str, str] -- name_of_object, json_representation
    """
    objects = {}
    for name in names:
        cursor.execute("""SELECT value from storage WHERE name = ?""", [name])
        row = cursor.fetchone()
        objects[name] = row[0] if row else ""
    return objects


# High Level FunctionsA =======================================================
def load_from_database(db_file: Union[Path, str],
                       names: list[str])-> dict[str, str]:
    """read data from database into objects

    Arguments:
        db_file: Path -- name of file to save objects to
        names: list[str] -- names of json_representations to fetch

    Raise:

    Return:
        dict[str, str] -- name_of_object, json_representation
    """
    file_message("Reading Data from {file_name}", db_file)
    with dbopen(db_file) as cursor:
        create_schema(cursor)
        return deserialized(cursor, names)

def save_to_database(db_file: Union[Path, str],
                     objects: dict[str, str])-> None:
    """jsonify and save save objects to database

    Arguments:
        db_file: Path -- name of file to save objects to
        objects: dict[str, str] -- name_of_object, json_representation

    Return:
        None
    """
    file_message("Saving Data to {file_name}", db_file)
    with dbopen(db_file) as cursor:
        create_schema(cursor)
        serialize(cursor, objects)


if __name__ == "__main__":
    pass
