#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Sections in middle of sheet
"""
from dataclasses import dataclass
from typing import Optional
from .costing_data import SectionInfo, Xlsx, YESNO
from .boms import BomPart

@dataclass
class Title():
    """Title text and formatting"""
    text: Optional[str]
    style: str

@dataclass
class ColumnInfo():
    """Title text and formatting"""
    name: Optional[str]
    style: str

TITLES = [
    Title('Vendor', 'heading1'),
    Title('Part #','heading1'),
    Title('Description', 'heading1'),
    Title('Price' ,'heading1'),
    Title('Per', 'heading1'),
    Title('Qty', 'heading1' ),
    Title('Sub Total', 'heading1'),
    Title('Shipping', 'heading1'),
    Title('Total', 'heading1'),
    Title('Last cost date', 'heading2'),
    Title('Vendor Quote', 'heading2'),
    Title('FAB' ,'heading1'),
    Title('PAINT' ,'heading1'),
    Title('OUTFIT', 'heading1'),
    Title('DESIGN', 'heading1'),
]

TITLES_FAB = [
    Title('Vendor', 'heading1'),
    Title('Part #', 'heading1'),
    Title('Description', 'heading1'),
    Title('Cost Per Pound' ,'heading1'),
    Title(None, 'heading1'),
    Title('Pounds', 'heading1' ),
    Title('Sub Total', 'heading1'),
    Title('Shipping', 'heading1'),
    Title('Total', 'heading1'),
    Title('Last cost date', 'heading2'),
    Title('Vendor Quote', 'heading2'),
]

TITLES_PAINT = [
    Title('Vendor', 'heading1'),
    Title('Part #', 'heading1'),
    Title('Description', 'heading1'),
    Title('Cost Per Gallon' ,'heading1'),
    Title(None, 'heading1'),
    Title('Qty', 'heading1' ),
    Title('Sub Total', 'heading1'),
    Title('Shipping', 'heading1'),
    Title('Total', 'heading1'),
    Title('Last cost date', 'heading2'),
    Title('Vendor Quote', 'heading2'),
]

ROW_PAINTFAB_PART = [
    ColumnInfo('vendor', 'normalBordered'),
    ColumnInfo('part', 'normalBordered'),
    ColumnInfo('description', 'normalBordered'),
    ColumnInfo('unitprice', 'currencyBordered'),
    ColumnInfo('uom', 'normalBordered'),
    ColumnInfo('qty', 'decimalBordered'),
    ColumnInfo('=D{}*F{}', 'currencyBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo('=H{}+G{}', 'currencyBordered'),
    ColumnInfo('updated', 'updated'),
    ColumnInfo(None, 'normalBordered'),
]


BLANK_PAINTFAB_PART = [
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None,'currencyBordered'),
    ColumnInfo(None, 'decimalBordered'),
    ColumnInfo(None, 'currencyBordered'),
    ColumnInfo('=D{}*F{}', 'currencyBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo('=H{}+G{}', 'normalBordered'),
    ColumnInfo(None, 'updated'),
    ColumnInfo(None, 'normalBordered'),
]


ROW_PART = [
    ColumnInfo('vendor', 'normalBordered'),
    ColumnInfo('part', 'normalBordered'),
    ColumnInfo('description', 'normalBordered'),
    ColumnInfo('unitprice', 'currencyBordered'),
    ColumnInfo('uom', 'normalBordered'),
    ColumnInfo('qty', 'normalBordered'),
    ColumnInfo('=D{}*F{}', 'currencyBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo('=H{}+G{}', 'currencyBordered'),
    ColumnInfo('updated', 'updated'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
]

BLANK_PART = [
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None,'currencyBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'currencyBordered'),
    ColumnInfo('=D{}*F{}', 'currencyBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo('=H{}+G{}', 'normalBordered'),
    ColumnInfo(None, 'updated'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
    ColumnInfo(None, 'normalBordered'),
]

GREEN_BLANK_PART = [
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None,'bgGreenCurrencyBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenCurrencyBordered'),
    ColumnInfo('=D{}*F{}', 'bgGreenCurrencyBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo('=H{}+G{}', 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenUpdated'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
    ColumnInfo(None, 'bgGreenNormalBordered'),
]


BLANK_BOM_PART = BomPart('', 0, 0, 0, 0, '', '', 0, '', '', '', None)
EMPTY_BOM_PART = BomPart('', None, None, None, None, None,
                         None, None, None, None, None, None)


# WRITING SECTION FUNCTIONS ===================================================
def section_hr_rule(xlsx: Xlsx, row: int) -> None:
    """draw thick underline between section"""
    for column in range(11):
        xlsx.write(row, column, None, xlsx.styles['thickBottom'])

def section_heading_large(xlsx: Xlsx, row: int, text: str) -> None:
    """Large Header"""
    xlsx.sheet.set_row(row, 15.75)
    xlsx.write(row, 2, text, xlsx.styles['bgSilverBold12pt'])

def section_heading_small(xlsx: Xlsx, row: int, text1: str,
                          text2: str = None) -> None:
    """Small Header"""
    xlsx.write(row, 2, text1, xlsx.styles['bgSilverBold10pt'])
    if text2:
        xlsx.write(row, 14, None, xlsx.styles['bgSilverBold10pt'])
        xlsx.merge_range(row, 11, row, 13, text2,
                         xlsx.styles['bgSilverBold10pt'])

def section_titles(xlsx: Xlsx, row: int, titles: list[Title]) -> None:
    """Titles for section"""
    xlsx.sheet.set_row(row, 23.85 )
    for column, title in enumerate(titles):
        xlsx.write(row, column, title.text, xlsx.styles[title.style])

def section_part(xlsx: Xlsx, row: int, columns_info: list[ColumnInfo],
                 part = BomPart) -> None :
    """write out one part to sheet"""
    for column, column_info in enumerate(columns_info):
        if  isinstance(column_info.name, str) and column_info.name[0] == "=":
            if part.qty is None:
                xlsx.write(row, column, None, xlsx.styles[column_info.style])
                continue
            formula = column_info.name.format(row + 1, row + 1)
            # both formulas have the same value
            total = (part.qty or 0) * (part.unitprice or 0)
            xlsx.write(row, column, formula, xlsx.styles[column_info.style],
                       total)
        else:
            if column_info.name:
                value = getattr(part, column_info.name)
            else:
                value = None
            xlsx.write(row, column, value, xlsx.styles[column_info.style])

def section_subtotal(xlsx: Xlsx, row: int, section_info: SectionInfo,
                     text: str) -> None:
    """write subtoal"""
    formula = (f"=SUM(I{str(section_info.start + 1)}"
               f":I{str(section_info.finish + 1)})")
    value = section_info.value
    xlsx.write(row, 7, text, xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula, xlsx.styles['currencyBorderedBold'], value)

def section_fabrication(xlsx: Xlsx, row: int,
                        section_info: dict[str, SectionInfo]) -> int:
    """write fabrication section"""
    total: float = 0
    section_heading_large(xlsx, row, 'Fabrication Materials')
    row += 2
    section_titles(xlsx, row, TITLES_FAB)
    row += 1
    start = row
    parts = xlsx.bom.sections[0].parts
    for part in parts:
        section_part(xlsx, row, ROW_PAINTFAB_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    section_part(xlsx, row, ROW_PAINTFAB_PART, BLANK_BOM_PART)
    finish = row
    row += 2
    xlsx.write(row, 2, 'Material sheet provided Y/N',
               xlsx.styles['bgYellowRight'])
    xlsx.write(row, 3,  None, xlsx.styles['bgYellow4'])
    xlsx.sheet.data_validation(row, 3, row, 3, {'validate': 'list', 'source':
                                                YESNO,})
    row += 1
    xlsx.write(row, 2, 'If no material sheet, indicate reference boat #',
               xlsx.styles['bgYellowRight'])
    xlsx.write(row, 3,  None, xlsx.styles['bgYellow4'])
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'TOTAL ALLOY COST')
    section_info['FABRICATION'] = section
    row += 1
    section_hr_rule(xlsx, row)
    row += 2
    return row

def section_paint(xlsx: Xlsx, row: int,
                        section_info: dict[str, SectionInfo]) -> int:
    """write paint section"""
    total: float = 0
    section_heading_large(xlsx, row, 'Paint Materials')
    row += 2
    section_titles(xlsx, row, TITLES_PAINT)
    row += 1
    start = row
    parts = xlsx.bom.sections[1].parts
    for part in parts:
        section_part(xlsx, row, ROW_PAINTFAB_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    section_part(xlsx, row, ROW_PAINTFAB_PART, BLANK_BOM_PART)
    xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
    finish = row
    row += 2
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'TOTAL PAINT COST')
    section_info['PAINT'] = section
    row += 1
    section_hr_rule(xlsx, row)
    row += 2
    return row

def section_green(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write green unused section"""
    _ = section_info
    total : float = 0
    section_heading_large(xlsx, row, 'Outfitting Materials')
    row += 2
    section_heading_small(
        xlsx,
        row,
        'Indicate revisions made here - denote by color above',
        'Labor Change add/delete')
    row += 1
    section_titles(xlsx, row, TITLES)
    row += 1
    section_part(xlsx, row, GREEN_BLANK_PART, EMPTY_BOM_PART)
    row += 1
    start = row
    for row1 in range(row, row + 22):
        section_part(xlsx, row1, GREEN_BLANK_PART, BLANK_BOM_PART)
    finish = row + 21
    row += 23
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'REVISIONS TOTAL')
    row += 1
    return row

def section_outfitting(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write outfitting section"""
    _ = section_info
    total : float = 0
    section_heading_small(
        xlsx,
        row,
        'Components / Materials',
        'Labor Change add/delete')
    row += 1
    section_titles(xlsx, row, TITLES)
    row += 1
    section_part(xlsx, row, BLANK_PART, EMPTY_BOM_PART)
    row += 1
    start = row
    parts = xlsx.bom.sections[2].parts
    for part in parts:
        section_part(xlsx, row, ROW_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
    xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
    finish = row
    row += 2
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'MATERIALS TOTAL')
    section.subtotal += 2
    section_info['OUTFITTING'] = section
    row += 2
    formula = f"=I{start - 3}+I{subtotal + 1}"
    xlsx.write(
        row,
        8,
        formula,
        xlsx.styles['bgSilverBorderedCurrency'],
        total)
    xlsx.write(
        row,
        7,
        "Total All Outfitting Components",
        xlsx.styles['rightJust2'])
    row += 2
    return row

def section_bigticket(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write big ticket section"""
    _ = section_info
    total: float = 0
    section_heading_small(
        xlsx,
        row,
        'BIG TICKET ITEMS (generator, Seakeeper, hyd pumps)',
        'Labor Change add/delete')
    row += 2
    section_titles(xlsx, row, TITLES)
    row += 1
    start = row
    parts = xlsx.bom.sections[3].parts
    for part in parts:
        section_part(xlsx, row, ROW_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
    xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
    finish = row
    row += 1
    if not parts:
        section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
        xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
        finish = row
        row += 1
    row += 1
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'BIG TICKET ITEMS TOTAL')
    section_info['BIG TICKET ITEMS'] = section
    row += 2
    return row

def section_outboard(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write outboard motors section"""
    _ = section_info
    total: float = 0
    section_heading_small(
        xlsx,
        row,
        'OB Motors',
        'Labor Change add/delete')
    row += 2
    section_titles(xlsx, row, TITLES)
    xlsx.write(row, 16, 'Dealer Net Price', xlsx.styles['heading1'])
    row += 1
    start = row
    parts = xlsx.bom.sections[4].parts
    for part in parts:
        section_part(xlsx, row, ROW_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    for row1 in range(max(3 - len(parts), 1)):
        _ = row1
        section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
        xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
        row += 1
    finish = row - 1
    row += 1
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'OB MOTORS TOTAL')
    section_info['OUTBOARD MOTORS'] = section
    for row1 in range(start, finish + 1):
        xlsx.write(row1, 16, None, xlsx.styles['currencyYellowBorder'])
    formula = f"=SUM(Q{start + 1}:Q{finish +1})"
    xlsx.write(row, 15, 'Dealer Net Total', xlsx.styles['rightJust2'])
    xlsx.write(row, 16, formula, xlsx.styles['currencyBorderedBold'], 0)
    row += 2
    return row

def section_inboard(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write inborad motors section"""
    _ = section_info
    total: float = 0
    section_heading_small(
        xlsx,
        row,
        'Inboard Motors & Jets',
        'Labor Change add/delete')
    row += 2
    section_titles(xlsx, row, TITLES)
    row += 1
    start = row
    parts = xlsx.bom.sections[5].parts
    for part in parts:
        section_part(xlsx, row, ROW_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    for row1 in range(max(4 - len(parts), 1)):
        _ = row1
        section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
        xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
        row += 1
    finish = row - 1
    row += 1
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'INBOARD MOTORS & JETS TOTAL')
    section_info['INBOARD MOTORS & JETS'] = section
    row += 2
    return row


def section_trailer(xlsx: Xlsx, row: int,
                   section_info: dict[str, SectionInfo]) -> int:
    """write trailer section"""
    _ = section_info
    text1 = ("NOTE: IF EZ LOADER, ADD 12% TO PRICING FROM JULY 2021 PRICE "
             "LIST. OTHERWISE RE-QUOTE WITH REVISED PRICING")
    total: float = 0
    section_heading_small(
        xlsx,
        row,
        'Trailer',
        'Labor Change add/delete')
    xlsx.merge_range(row, 16, row + 3, 20, text1, xlsx.styles['redMerged'])
    row += 2
    section_titles(xlsx, row, TITLES)
    row += 1
    start = row
    parts = xlsx.bom.sections[6].parts
    for part in parts:
        section_part(xlsx, row, ROW_PART, parts[part])
        total += (parts[part].qty or 0) * (parts[part].unitprice or 0)
        row += 1
    section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
    xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
    finish = row
    row += 1
    if not parts:
        section_part(xlsx, row, ROW_PART, BLANK_BOM_PART)
        xlsx.write(row, 5, None, xlsx.styles['normalBordered'])
        finish = row
        row += 1
    row += 1
    subtotal = row
    section = SectionInfo(start, finish, subtotal + 1, total)
    section_subtotal(xlsx, row, section, 'TRAILER TOTAL')
    section_info['TRAILER'] = section
    row += 2
    return row

def generate_sections(xlsx: Xlsx,
                      section_info: dict[str, SectionInfo]) -> None:
    """generate costing sheet sections

    Arguments:
        xlsx: Xlsx -- information about spreadsheet
        sections: SectionInfo -- information about how sections are laid out

    Returns:
        None
    """
    row: int = 10
    row = section_fabrication(xlsx, row, section_info)
    row = section_paint(xlsx, row, section_info)
    row = section_green(xlsx, row, section_info)
    row = section_outfitting(xlsx, row, section_info)
    row = section_bigticket(xlsx, row, section_info)
    row = section_outboard(xlsx, row, section_info)
    row = section_inboard(xlsx, row, section_info)
    row = section_trailer(xlsx, row, section_info)
    section = SectionInfo(row, row + 5, row + 8, 0)
    section_info['TOTALS'] = section


if __name__ == "__main__":
    pass
