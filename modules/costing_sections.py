#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Sections in middle of sheet
"""
from .costing_data import SectionInfo, Xlsx

# WRITING SECTION FUNCTIONS ===================================================
def generate_sections(xlsx: Xlsx, section_info:
                      dict[str, SectionInfo]) -> None:
    """generate costing sheet sections

    Arguments:
        xlsx: Xlsx -- information about spreadsheet
        sections: SectionInfo -- information about how sections are laid out

    Returns:
        None
    """
    _ = (xlsx, section_info)

if __name__ == "__main__":
    pass
