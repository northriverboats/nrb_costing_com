#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Totals Section at bottom of sheet
"""
from .costing_data import SectionInfo, Xlsx, DEALERS, SALESPERSON, YESNO

# WRITING TOTALS FUNCTIONS ====================================================
def totals_column_b(xlsx: Xlsx,
                      section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet column b"""
    offset = section_info['TRAILER'].subtotal + 2

    # COLUMN B ================================================================
    for row in range(offset + 79, offset + 86):
        xlsx.write(row, 1, None, xlsx.styles['bgYellow4'])
    for row in range(offset + 86, offset + 89):
        xlsx.write(row, 1, None, xlsx.styles['bgGreen3'])
    xlsx.sheet.data_validation(offset + 79, 1, offset + 85, 1, {
        'validate': 'list',
        'source': YESNO,
    })
    xlsx.sheet.data_validation(offset + 86, 1, offset + 86, 1, {
        'validate': 'list',
        'source': SALESPERSON,
    })
    xlsx.sheet.data_validation(offset + 87, 1, offset + 88, 1, {
        'validate': 'list',
        'source': YESNO,
    })

def totals_00(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 00 from the bottom of the sheet"""
    formula1 = f"=I{section_info['FABRICATION'].subtotal}"
    value1 = section_info['FABRICATION'].value

    xlsx.write(row, 2, 'MATERIALS', xlsx.styles['rightJust2'])
    xlsx.write(row, 3, 'Fabrication', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_01(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 01 from the bottom of the sheet"""
    formula1 = f"=I{row}*H{row + 1}"
    value1 = section_info['FABRICATION'].value * 0.08

    xlsx.write(row, 3, 'Fab Consumables', xlsx.styles['generic1'])
    xlsx.write(row, 7, 0.08, xlsx.styles['percent'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_02(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 02 from the bottom of the sheet"""
    formula1 = f"=I{section_info['PAINT'].subtotal}"
    value1 = section_info['PAINT'].value

    xlsx.write(row, 3, 'Paint', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_03(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 03 from the bottom of the sheet"""
    formula1 = f"=I{row}*H{row+1}"
    value1 = section_info['PAINT'].value * 0.50

    xlsx.write(row, 3, 'Paint Consumables', xlsx.styles['generic1'])
    xlsx.write(row, 7, 0.50, xlsx.styles['percent'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_04(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 04 from the bottom of the sheet"""
    formula1 = f"=I{section_info['OUTFITTING'].subtotal}"
    value1 = section_info['OUTFITTING'].value

    xlsx.write(row, 3, 'Outfitting', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_05(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 05 from the bottom of the sheet"""
    formula1 = f"=I{section_info['BIG TICKET ITEMS'].subtotal}"
    value1 = section_info['BIG TICKET ITEMS'].value

    xlsx.write(row, 3, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_06(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 06 from the bottom of the sheet"""
    formula1 = f"=I{section_info['OUTBOARD MOTORS'].subtotal}"
    value1 = section_info['OUTBOARD MOTORS'].value

    xlsx.write(row, 3, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_07(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 07 from the bottom of the sheet"""
    formula1 = f"=I{section_info['INBOARD MOTORS & JETS'].subtotal}"
    value1 = section_info['INBOARD MOTORS & JETS'].value

    xlsx.write(row, 3, 'IB Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_08(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 08 from the bottom of the sheet"""
    formula1 = f"=I{section_info['TRAILER'].subtotal}"
    value1 = section_info['TRAILER'].value

    xlsx.write(row, 3, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(row, 8, formula1, xlsx.styles['currency'], value1)

def totals_09(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 09 from the bottom of the sheet"""
    formula1 = f"=SUM(I{row - 8}:I{row})"
    # value used in totals_40
    value1 = (sum([section_info[section].value for section in section_info]) +
                   section_info['FABRICATION'].value * 0.08 +
                   section_info['PAINT'].value * 0.50)

    xlsx.write(row, 7, 'Total All Materials', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_12(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 12 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'Labor', xlsx.styles['rightJust2'])

def totals_13(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 13 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 5, 'BOAT HOURS', xlsx.styles['centerJust1'])
    xlsx.write(row, 6, 'TOTAL HOURS', xlsx.styles['centerJust1'])
    xlsx.write(row, 7, 'RATE', xlsx.styles['centerJust1'])

def totals_14(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 14 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=F{row + 1}+SUM(L:L)"
    value1 = 0
    formula2 = f"=H{row +1}*G{row + 1}"
    value2 = 0

    xlsx.write(row, 3, 'Fabrication', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 59.22, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_15(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 15 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=F{row + 1}+SUM(M:M)"
    value1 = 0
    formula2 = f"=H{row +1}*G{row + 1}"
    value2 = 0

    xlsx.write(row, 3, 'Paint', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 32.97, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_16(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 16 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=F{row + 1}+SUM(N:N)"
    value1 = 0
    formula2 = f"=H{row +1}*G{row + 1}"
    value2 = 0

    xlsx.write(row, 3, 'Outfitting', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 36.91, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_17(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 17 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=F{row + 1}+SUM(O:O)"
    value1 = 0
    formula2 = f"=H{row +1}*G{row + 1}"
    value2 = 0

    xlsx.write(row, 3, 'Design / Drafting', xlsx.styles['generic1'])
    xlsx.write(row, 5, 0.0 , xlsx.styles['centerJust2'])
    xlsx.write(row, 6, formula1, xlsx.styles['centerJust2'], value1)
    xlsx.write(row, 7, 43.03, xlsx.styles['currency'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_19(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 19 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=SUM(F{row - 4}:F{row - 1})"
    value1 = 0
    formula2 = f"=SUM(I{row - 4}:I{row - 1})"
    value2 = 0

    xlsx.write(row, 4, 'Total Hours', xlsx.styles['rightJust1'])
    xlsx.write(row, 5, formula1, xlsx.styles['bgYellow0'], value1)
    xlsx.write(row, 7, 'Total Labor Costs', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula2, xlsx.styles['currencyBoldYellow'], value2)

def totals_20(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 20 from the bottom of the sheet"""
    _ = section_info
    text1 = 'Indicate boat referenced for labor hours if used'

    xlsx.write(row, 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(row, 3, None, xlsx.styles['bgYellow4'])

def totals_23(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 23 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'Other Costs', xlsx.styles['rightJust2'])

def totals_25(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 25 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Test Fuel', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_26(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 26 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Trials', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_27(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 27 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Engineering', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_28(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 28 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_29(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 29 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_30(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 30 from the bottom of the sheet"""
    _ = section_info

    formula1 = f"=SUM(I{row - 5}:I{row})"
    value1 = 0

    xlsx.write(row, 7, 'Total Other Costs', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_32(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 32 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 2, 'NO MARGIN ITEMS', xlsx.styles['rightJust2'])

def totals_34(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 34 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, 'Trucking', xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_35(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 35 from the bottom of the sheet"""
    _ = section_info
    text1 = "Voyager/Custom - 10%, Guide/Lodge - 3%"

    xlsx.write(row, 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(row, 3, 'Dealer commission', xlsx.styles['generic1'])
    xlsx.write(row, 5, 'Dealer', xlsx.styles['centerJust2'])
    xlsx.write(row, 6, None, xlsx.styles['bgYellow4'])
    xlsx.sheet.data_validation(row, 6, row, 6, {
        'validate': 'list',
        'source': DEALERS,
    })
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_36(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 36 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_37(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 37 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 3, None, xlsx.styles['generic1'])
    xlsx.write(row, 8, 0.0, xlsx.styles['currency'])

def totals_38(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 38 from the bottom of the sheet"""
    _ = section_info

    formula1 = f"=SUM(I{row - 3}:I{row})"
    value1 = 0

    xlsx.write(row, 7, 'Total No Margin Items', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_40(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 40 from the bottom of the sheet"""
    formula1 = f"=I{row -30}+I{row - 20}+I{row - 9}+I{row - 1}"
    # value used in totals_40 and totals_43
    value1 = (sum([section_info[section].value for section in section_info]) +
                   section_info['FABRICATION'].value * 0.08 +
                   section_info['PAINT'].value * 0.50)

    xlsx.write(row, 6, 'TOTAL COST OF PROJECT', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBoldYellow'], value1)


def totals_42(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 42 from the bottom of the sheet"""
    _ = section_info
    text1 = "Mark up per pricing policy: "

    xlsx.sheet.set_row(row, 23.85)
    xlsx.write(row, 2, text1, xlsx.styles['generic2'])
    xlsx.write(row, 3, 'Cost ', xlsx.styles['centerJust1'])
    xlsx.merge_range(row, 4, row, 5, 'Markup',  xlsx.styles['centerJust3'])
    xlsx.write(row, 6, 'MSRP ', xlsx.styles['centerJust3'])
    xlsx.write(row, 7, 'Discount ', xlsx.styles['centerJust3'])
    xlsx.write(row, 9, 'Contribution Margin', xlsx.styles['centerJust4'])

def totals_43(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 43 from the bottom of the sheet"""
    formula1 = (f"=I{row- 2}-I{row - 4}-I{row - 37}-I{row - 36}-I{row - 35}"
                f"-I{row - 34}")
    # value used in totals_40
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * 0.08 +
              section_info['PAINT'].value +
              section_info['PAINT'].value * 0.50 +
              section_info['OUTFITTING'].value)
    formula2 = f"=D{row + 1}/E{row + 1}/F{row + 1}"
    value2 = value1 / 0.61 / 0.7
    formula3 = f"=G{row + 1}*(1-H{row + 1})"
    value3 = value2
    formula4 = f"=IF(I{row + 1}=0,0,(I{row + 1}-D{row + 1})/I{row + 1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals = value3

    xlsx.write(row, 2, 'Boat and options:', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.61, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_44(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 44 from the bottom of the sheet"""
    formula1 = f"=I{row - 38}"
    value1 = section_info['BIG TICKET ITEMS'].value
    formula2 = f"=D{row + 1}/E{row + 1}/F{row + 1}"
    value2 = value1 / 0.8 / 0.85
    formula3 = "=G" + str(row + 1) + "*(1-H" + str(row + 1 ) + ')'
    value3 = value2
    formula4 = f"=IF(I{row + 1}=0,0,(I{row + 1}-D{row + 1})/I{row + 1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.8, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.85, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_45(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 45 from the bottom of the sheet"""
    formula1 = f"=I{row - 38}"
    value1 = section_info['OUTBOARD MOTORS'].value
    formula2 = f"=Q{section_info['OUTBOARD MOTORS'].subtotal}"
    value2 = 0.0
    formula3 = f"=G{row + 1}*(1-H{row + 1})"
    value3 = value2
    formula4 = f"=IF(I{row + 1}=0,0,(I{row + 1}-D{row + 1})/I{row + 1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(row, 4, row, 5, 'See PP',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_46(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 46 from the bottom of the sheet"""
    formula1 = f"=I{row - 38}"
    value1 = section_info['INBOARD MOTORS & JETS'].value
    formula2 = f"=D{row + 1}/E{row +1}/F{row +1}"
    value2 = value1 / 0.85 / 0.7
    formula3 = f"=G{row + 1}*(1-H{row + 1})"
    value3 = value2
    formula4 = f"=IF(I{row + 1}=0,0,(I{row + 1}-D{row + 1})/I{row + 1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Inboard Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.85, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_47(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 47 from the bottom of the sheet"""
    formula1 = f"=I{row - 38}"
    value1 = section_info['TRAILER'].value
    formula2 = f"=D{row + 1}/E{row + 1}/F{row + 1}"
    value2 = value1 / 0.8 / 0.7
    formula3 = f"=G{row + 1}*(1-H{row + 1})"
    value3 = value2
    formula4 = f"=IF(I{row + 1}=0,0,(I{row + 1}-D{row + 1})/I{row + 1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(row, 4, 0.8, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, None, xlsx.styles['percentBorderYellow'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_48(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 48 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=I{row - 9}"
    value1 = 0.0
    formula2 = f"=D{row + 1}"
    value2 = 0.0
    formula3 = f"=G{row + 1}"
    value3 = value2
    formula4 = f"=IF(I{row +1}=0,0,(I{row + 1}-D{row + 1})/I{row +1})"
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(row, 2, 'No margin items: ', xlsx.styles['generic1'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(row, 4, row, 5, 'none',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(row, 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(row, 7, 'none', xlsx.styles['centerJust5'])
    xlsx.write(row, 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(row, 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_50(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 50 from the bottom of the sheet"""
    text1 = "Total Cost (equals total cost of project box)"
    formula1 = f"=SUM(D{row - 6}:D{row -1})"
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * 0.08 +
              section_info['PAINT'].value +
              section_info['PAINT'].value * 0.50 +
              section_info['OUTFITTING'].value +
              section_info['BIG TICKET ITEMS'].value +
              section_info['OUTBOARD MOTORS'].value +
              section_info['INBOARD MOTORS & JETS'].value +
              section_info['TRAILER'].value)
    formula2 = f"=SUM(I{row - 6}:I{row - 1})"
    value2 = section_info['TOTALS'].totals

    xlsx.write(row, 2, text1, xlsx.styles['rightJust2'])
    xlsx.write(row, 3, formula1, xlsx.styles['currencyYellow'], value1)
    xlsx.write(row, 7, 'Calculated Selling Price', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula2, xlsx.styles['currency'], value2)

def totals_52(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 52 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(row, 6, 'SELLING PRICE', xlsx.styles['rightJust2'])
    xlsx.write(row, 8, None, xlsx.styles['currencyBoldYellowBorder'])

def totals_54(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 54 from the bottom of the sheet"""
    text1 = "CONTRIBUTION TO PROFIT AND OVERHEAD"
    formula1 = f"=I{row - 1}-I{row - 13}"
    value1 = -(sum([section_info[section].value for section in section_info]) +
                    section_info['FABRICATION'].value * 0.08 +
                    section_info['PAINT'].value * 0.50)

    xlsx.write(row, 6, text1, xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['currencyBold'], value1)

def totals_56(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 56 from the bottom of the sheet"""
    _ = section_info
    text1 = "CONTRIBUTION MARGIN"
    formula1 = f"=IF(I{row - 3}=0,0,SUM(I{row - 3}-I{row - 15})/I{row - 3})"

    value1 =  0.0

    xlsx.write(row, 6, text1, xlsx.styles['rightJust2'])
    xlsx.write(row, 8, formula1, xlsx.styles['percent1'], value1)


def totals_59(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 59 from the bottom of the sheet"""
    _ = section_info
    text1 = "Pricing Policy References: "
    text2 = "Discounts / Minimum contribution margins: "

    xlsx.write(row, 2, text1, xlsx.styles['generic2'])
    xlsx.write(row, 5, text2, xlsx.styles['generic2'])

def totals_60(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 60 from the bottom of the sheet"""
    _ = section_info
    text1 = "Boat MSRP = C / .61 / 0.7"
    text2 = ("Government/Commercial Discounts - Max discount 30% / "
             "Minimum margin 35%")

    xlsx.write(row, 2, text1, xlsx.styles['generic1'])
    xlsx.write(row, 5, text2, xlsx.styles['generic1'])

def totals_61(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 61 from the bottom of the sheet"""
    _ = section_info
    text1 = "Options MSRP = C / .8046 / .48"
    text2 = ("Guide / Lodge Program - Commercial Markup- Max discount "
             "30% / Minimum margin 35%")

    xlsx.write(row, 2, text1, xlsx.styles['generic1'])
    xlsx.write(row, 5, text2, xlsx.styles['generic1'])

def totals_62(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 62 from the bottom of the sheet"""
    _ = section_info
    text1 = "Trailers MSRP = C / 0.80 / 0.7"
    text2 = ("Guide / Lodge Program - Recreational Retail Price list- "
             "Max discount 20%")

    xlsx.write(row, 2, text1, xlsx.styles['generic1'])
    xlsx.write(row, 5, text2, xlsx.styles['generic1'])

def totals_63(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 63 from the bottom of the sheet"""
    _ = section_info
    text1 = "Inboard Motors MSRP = C / 0.85 / 0.7"
    text2 = ("Non-Commercial Direct Sales - Max discount 26% / Minimum "
             "margin 38.5%")

    xlsx.write(row, 2, text1, xlsx.styles['generic1'])
    xlsx.write(row, 5, text2, xlsx.styles['generic1'])

def totals_64(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 64 from the bottom of the sheet"""
    _ = section_info
    text1 = "Big Ticket Items MSRP = C / (range from 0.80 – 0.85) / 0.7"

    xlsx.write(row, 2, text1, xlsx.styles['generic1'])
    xlsx.sheet.write_rich_string(
        row, 5, xlsx.styles['generic1'], 'Voyager - ',
        xlsx.styles['red'],
        "SEE MIKE ON ALL VOYAGER OR CUSTOM DEALER REFERRAL PRICING")

def totals_65(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 60 from the bottom of the sheet"""
    _ = section_info
    text1 = "GSA Pricing"
    text2 = "1 - 2 boats"
    text3 = "30% discount"

    xlsx.write(row, 5, text1, xlsx.styles['generic1'])
    xlsx.write(row, 7, text2, xlsx.styles['generic1'])
    xlsx.write(row, 8, text3, xlsx.styles['generic1'])

def totals_66(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 66 from the bottom of the sheet"""
    _ = section_info
    text1 = "3 boats"
    text2 = "30.5% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])

def totals_67(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 67 from the bottom of the sheet"""
    _ = section_info
    text1 = "4 boats"
    text2 = "31% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])

def totals_68(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 68 from the bottom of the sheet"""
    _ = section_info
    text1 = "5 boats"
    text2 = "31.5% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])

def totals_69(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 69 from the bottom of the sheet"""
    _ = section_info
    text1 = "6 - 10 bts"
    text2 = "32% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])

def totals_70(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 70 from the bottom of the sheet"""
    _ = section_info
    text1 = "11 - 20 bts"
    text2 = "32.25% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])

def totals_71(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 71 from the bottom of the sheet"""
    _ = section_info
    text1 = "20+ boats"
    text2 = "32.5% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])
    xlsx.write(row, 8, text2, xlsx.styles['generic1'])


def totals_72(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 72 from the bottom of the sheet"""
    _ = section_info
    text1 = "Outboard motors"
    text2 = "Government agencies - 15% discount"

    xlsx.write(row, 5, text1, xlsx.styles['generic1'])
    xlsx.write(row, 7, text2, xlsx.styles['generic1'])

def totals_73(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 73 from the bottom of the sheet"""
    _ = section_info
    text1 = "(approval req. for greater discount, no more than addl. 3%)"
    text2 = "GSA Pricing - 18% discount"

    xlsx.merge_range(row, 5, row + 2, 6, text1, xlsx.styles['italicsNote'])
    xlsx.write(row, 7, text2, xlsx.styles['generic1'])

def totals_74(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 74 from the bottom of the sheet"""
    _ = section_info
    text1 = "Guides / Lodges - 10% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])

def totals_75(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 75 from the bottom of the sheet"""
    _ = section_info
    text1 = "Commercial sales - 5% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])

def totals_76(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 76 from the bottom of the sheet"""
    _ = section_info
    text1 = "Voyager - 5% discount"

    xlsx.write(row, 7, text1, xlsx.styles['generic1'])

def totals_78(xlsx: Xlsx, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 78 from the bottom of the sheet"""
    _ = section_info
    style1 = xlsx.styles['generic1']
    style2 = xlsx.styles['generic2']

    xlsx.write(row + 0, 2,
               'Cost estimate check list - complete prior to sending quote i'
               'or submitting bid', style2)
    xlsx.write(row + 1, 2,
               'Verify all formulas are correct and all items are included '
               'in cost total', style1)
    xlsx.write(row + 2, 2,
               'Verify aluminum calculated with total lbs included. Include '
               'metal costing sheet separate if completed', style1)
    xlsx.write(row + 3, 2,
               'Verify paint costing equals paint description', style1)
    xlsx.write(row + 4, 2,
               'Cost estimate includes all components on sales quote', style1)
    xlsx.write(row + 5, 2,
               'Pricing policy discounts and minimum margins are met', style1)
    xlsx.write(row + 6, 2,
               'Vendor quotes received and included in costing folder', style1)
    xlsx.write(row + 7, 2,
               'Labor hours reviewed and correct to best knowledge of project',
               style1)
    xlsx.write(row + 8, 2,
               'Name of peer who reviewed prior to submission to customer',
               style1)
    xlsx.write(row + 9, 2,
               'Customer signed sales quotation', style1)
    xlsx.write(row + 10, 2,
               'Customer provided terms and conditions, including payment '
               'schedule', style1)


def generate_totals(xlsx: Xlsx, section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet"""
    skip = [10, 11, 18, 21, 22, 24, 31, 33,
            39, 41, 49, 51, 53, 55, 57, 58, 77]
    # pylint: disable=unused-variable
    offset = section_info['TRAILER'].subtotal + 2
    for row in range(0, 79):
        if row in skip:
            continue
        # pylint: disable=eval-used
        eval(f"totals_{row:02}(xlsx, section_info, row + offset)")
    totals_column_b(xlsx, section_info)

if __name__ == "__main__":
    pass
