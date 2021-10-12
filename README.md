NRB COMMERCIAL COSTING SHEET GENERATOR
=====================================
What it does:
* Allows the collection of parts data from sheets in `RESOURCE_FOLDER` to be compiled into a database.
* Allows the collection of parts and other releated boat or cabin information in `BOAT_FOLDER` to be compiled into a database.
* Allows the collection of boats, cabins, and file names to be compiled into a database from the `MASTER_FILE`.
* Allows the generation of a costing sheet based boat/cabin from the `MASTER_FILE` and the releated data in the `BOATS_FOLDER` for each length of boat to be combined on the `TEMPLATE FILE` and saved into the `SHEETS_FOLDER`.

Data cached in the database can be used for regenerating each phase:
* sheets can be generated from tables for `MASTER_FILE`, `BOAT_FOLDER` and `PARTS_FOLDER` without re-reading any of the files in those folders.
* the `MASTER_FILE` related tables can be derived from reading the `MASTER_FILE` and using the `BOATS_FOLDER` and `PARTS_FOLDER` without re-reading any of the files in those folders.
* the `BOATS_FOLDER` table can be genereated from the files in the `BOATS_FOLDER` and the `PARTS_FOLDER` table without re-reading any of thoes files.
* the `PARTS_FOLDER` table cn be generated from the files in the `PARTS_FOLDER`.
