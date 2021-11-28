#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Sections in middle of sheet
"""
from .costing_data import SectionInfo, Xlsx

# WRITING SECTION FUNCTIONS ===================================================
def section_hr_rule(xlsx: Xlsx, row: int)-> None:
    """draw thick underline between section"""
    for column in range(11):
        xlsx.write(row, column, None, xlsx.styles['thickBottom'])

def section_heading_large(xlsx: Xlsx, row: int, text: str)-> None:
    """Large Header"""
    xlsx.sheet.set_row(row, 15.75)
    xlsx.write(row, 2, text, xlsx.styles['bgSilverBold12pt'])

def section_heading_small(xlsx: Xlsx, row: int, text1: str,
                          text2: str = None)-> None:
    """Small Header"""
    xlsx.write(row, 2, text1, xlsx.styles['bgSilverBold10pt'])
    if text2:
        xlsx.write(row, 14, None, xlsx.styles['bgSilverBold10pt'])
        xlsx.merge_range(row, 11, row, 13, text2,
                         xlsx.styles['bgSilverBold10pt'])

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
    section_heading_large(xlsx, 15, 'Fabrication Materials')
    section_heading_small(
        xlsx, 17, 'Indicate revisions made here - denote by color above')
    section_hr_rule(xlsx, 19)

if __name__ == "__main__":
    pass
