#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Totals Section at bottom of sheet
"""
from .costing_data import SectionInfo, XlsxBom, DEALERS, SALESPERSON, YESNO
from . import config

# UTILITY FUNCTIONS =========================================================+=
def get_labor(xlsx: XlsxBom)-> float:
    """compute labor totals"""
    fabrication = xlsx.bom.labor['Fabrication'] or 0.0
    paint = xlsx.bom.labor['Paint'] or 0.0
    outfitting = xlsx.bom.labor['Outfitting'] or 0.0
    design  = xlsx.bom.labor['Design / Drafting'] or 0.0
    labor = (
        fabrication *
        xlsx.settings.hourly_rates['Fabrication Hours'].rate +
        paint * xlsx.settings.hourly_rates['Paint Hours'].rate +
        outfitting * xlsx.settings.hourly_rates['Outfitting Hours'].rate +
        design * xlsx.settings.hourly_rates['Design Hours'].rate)
    return labor

def adjust(row: int, offset: int)-> int:
    """adjusting for commision/hmac lines as neccessary
    threshold is based on line 49 is where we start inserting things
    """
    adjusted = row + offset

    # if config.hgac and offset < 0 and adjusted  < (config.offset + 51):
    #   return adjusted - 2

    if config.hgac:
        return adjusted

    if adjusted > (config.offset + 50):
        adjusted = adjusted - 2

    if offset < 0 and (adjusted < (config.offset + 48)):
        adjusted = adjusted -2
        
    return adjusted

# WRITING TOTALS FUNCTIONS ====================================================
def totals_column_b(xlsx: XlsxBom,
                      section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet column b"""

    # COLUMN B ================================================================
    for row in range(adjust(config.offset, 81), adjust(config.offset, 88)):
        xlsx.write(row, 1, None, xlsx.styles['bgYellow4'])
    for row in range(adjust(config.offset, 88), adjust(config.offset, 91)):
       xlsx.write(row, 1, None, xlsx.styles['bgGreen3'])
    xlsx.sheet.data_validation(adjust(config.offset, 81), 1, 
                               adjust(config.offset, 87), 1, {
        'validate': 'list',
        'source': YESNO,
    })
    xlsx.sheet.data_validation(adjust(config.offset, 88), 1,
                               adjust(config.offset, 88), 1, {
        'validate': 'list',
        'source': SALESPERSON,
    })
    xlsx.sheet.data_validation(adjust(config.offset, 89), 1,
                               adjust(config.offset, 90), 1, {
        'validate': 'list',
        'source': YESNO,
    })

def totals_00(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 00 from the bottom of the sheet"""
    formula1 = f"=I{section_info['FABRICATION'].subtotal}"
    value1 = section_info['FABRICATION'].value

    xlsx.write(adjust(row, 0), 2, 'MATERIALS', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 3, 'Fabrication', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_01(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 01 from the bottom of the sheet"""
    rate = xlsx.settings.consumables['FABRICATION'].rate

    formula1 = f"=I{adjust(row, 0)}*H{adjust(row, 1)}"
    value1 = section_info['FABRICATION'].value * rate

    xlsx.write(adjust(row, 0), 3, 'Fab Consumables', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['percent'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_02(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 02 from the bottom of the sheet"""
    formula1 = f"=I{section_info['PAINT'].subtotal}"
    value1 = section_info['PAINT'].value

    xlsx.write(adjust(row, 0), 3, 'Paint', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_03(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 03 from the bottom of the sheet"""
    rate = xlsx.settings.consumables['PAINT'].rate

    formula1 = f"=I{adjust(row, 0)}*H{adjust(row, 1)}"
    value1 = section_info['PAINT'].value * rate

    xlsx.write(adjust(row, 0), 3, 'Paint Consumables', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['percent'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_04(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 04 from the bottom of the sheet"""
    formula1 = f"=I{section_info['OUTFITTING'].subtotal}"
    value1 = section_info['OUTFITTING'].value

    xlsx.write(adjust(row, 0), 3, 'Outfitting', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_05(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 05 from the bottom of the sheet"""
    formula1 = f"=I{section_info['BIG TICKET ITEMS'].subtotal}"
    value1 = section_info['BIG TICKET ITEMS'].value

    xlsx.write(adjust(row, 0), 3, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_06(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 06 from the bottom of the sheet"""
    formula1 = f"=I{section_info['OUTBOARD MOTORS'].subtotal}"
    value1 = section_info['OUTBOARD MOTORS'].value

    xlsx.write(adjust(row, 0), 3, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_07(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 07 from the bottom of the sheet"""
    formula1 = f"=I{section_info['INBOARD MOTORS & JETS'].subtotal}"
    value1 = section_info['INBOARD MOTORS & JETS'].value

    xlsx.write(adjust(row, 0), 3, 'IB Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_08(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 08 from the bottom of the sheet"""
    formula1 = f"=I{section_info['TRAILER'].subtotal}"
    value1 = section_info['TRAILER'].value

    xlsx.write(adjust(row, 0), 3, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currency'], value1)

def totals_09(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 09 from the bottom of the sheet"""
    formula1 = f"=SUM(I{adjust(row, -8)}:I{adjust(row, 0)})"
    # value used in totals_40
    rate_fabrication = xlsx.settings.consumables['FABRICATION'].rate
    rate_paint = xlsx.settings.consumables['PAINT'].rate
    value1 = (sum([section_info[section].value for section in section_info]) +
                   section_info['FABRICATION'].value * rate_fabrication +
                   section_info['PAINT'].value * rate_paint)

    xlsx.write(adjust(row, 0), 7, 'Total All Materials', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_12(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 12 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 2, 'Labor', xlsx.styles['rightJust2'])

def totals_13(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 13 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 5, 'BOAT HOURS', xlsx.styles['centerJust1'])
    xlsx.write(adjust(row, 0), 6, 'TOTAL HOURS', xlsx.styles['centerJust1'])
    xlsx.write(adjust(row, 0), 7, 'RATE', xlsx.styles['centerJust1'])

def totals_14(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 14 from the bottom of the sheet"""
    _ = section_info
    dept = 'Fabrication'
    rate = xlsx.settings.hourly_rates['Fabrication Hours'].rate

    formula1 = f"=F{adjust(row, 1)}+SUM(M:M)"
    value1 = xlsx.bom.labor[dept] or 0.0
    formula2 = f"=H{adjust(row, 1)}*G{adjust(row, 1)}"
    value2 = value1 * rate

    xlsx.write(adjust(row, 0), 3, dept, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, value1, xlsx.styles['decimal'])
    xlsx.write(adjust(row, 0), 6, formula1, xlsx.styles['decimal'], value1)
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['currency'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currency'], value2)

def totals_15(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 15 from the bottom of the sheet"""
    _ = section_info
    dept = 'Paint'
    rate = xlsx.settings.hourly_rates['Paint Hours'].rate

    formula1 = f"=F{adjust(row, 1)}+SUM(M:M)"
    value1 = xlsx.bom.labor[dept] or 0.0
    formula2 = f"=H{adjust(row, 1)}*G{adjust(row, 1)}"
    value2 = value1 * rate

    xlsx.write(adjust(row, 0), 3, dept, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, value1, xlsx.styles['decimal'])
    xlsx.write(adjust(row, 0), 6, formula1, xlsx.styles['decimal'], value1)
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['currency'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currency'], value2)

def totals_16(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 16 from the bottom of the sheet"""
    _ = section_info
    dept = 'Outfitting'
    rate = xlsx.settings.hourly_rates['Outfitting Hours'].rate

    formula1 = f"=F{adjust(row, 1)}+SUM(N:N)"
    value1 = xlsx.bom.labor[dept] or 0.0
    formula2 = f"=H{adjust(row, 1)}*G{adjust(row, 1)}"
    value2 = value1 * rate

    xlsx.write(adjust(row, 0), 3, dept, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, value1 , xlsx.styles['decimal'])
    xlsx.write(adjust(row, 0), 6, formula1, xlsx.styles['decimal'], value1)
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['currency'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currency'], value2)

def totals_17(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 17 from the bottom of the sheet"""
    _ = section_info
    dept = 'Design / Drafting'
    rate = xlsx.settings.hourly_rates['Design Hours'].rate

    formula1 = f"=F{adjust(row, 1)}+SUM(O:O)"
    value1 = (xlsx.bom.labor[dept] or 0.0)
    formula2 = f"=H{adjust(row, 1)}*G{adjust(row, 1)}"
    value2 = value1 * rate

    xlsx.write(adjust(row, 0), 3, dept, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, value1, xlsx.styles['decimal'])
    xlsx.write(adjust(row, 0), 6, formula1, xlsx.styles['decimal'], value1)
    xlsx.write(adjust(row, 0), 7, rate, xlsx.styles['currency'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currency'], value2)

def totals_19(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 19 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=SUM(F{adjust(row, -4)}:F{adjust(row, -1)})"
    fabrication = xlsx.bom.labor['Fabrication'] or 0.0
    paint = xlsx.bom.labor['Paint'] or 0.0
    outfitting = xlsx.bom.labor['Outfitting'] or 0.0
    design  = xlsx.bom.labor['Design / Drafting'] or 0.0
    value1 = fabrication + paint + outfitting + design
    value1 =  xlsx.bom.labor['Total'] or 0.0

    formula2 = f"=SUM(I{adjust(row, -4)}:I{adjust(row, -1)})"
    value2 = get_labor(xlsx)

    xlsx.write(adjust(row, 0), 4, 'Total Hours', xlsx.styles['rightJust1'])
    xlsx.write(adjust(row, 0), 5, formula1, xlsx.styles['bgYellowDecimal'], value1)
    xlsx.write(adjust(row, 0), 7, 'Total Labor Costs', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currencyBoldYellow'], value2)

def totals_20(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 20 from the bottom of the sheet"""
    _ = section_info
    text1 = 'Indicate boat referenced for labor hours if used'

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(adjust(row, 0), 3, None, xlsx.styles['bgYellow4'])

def totals_23(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 23 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 2, 'Other Costs', xlsx.styles['rightJust2'])

def totals_25(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 25 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, 'Test Fuel', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, 'See outfitting materials', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_26(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 26 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, 'Trials', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_27(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 27 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, 'Engineering', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_28(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 28 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, None, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_29(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 29 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, None, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_30(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 30 from the bottom of the sheet"""
    _ = section_info

    formula1 = f"=SUM(I{adjust(row, -5)}:I{adjust(row, 0)})"
    value1 = 0

    xlsx.write(adjust(row, 0), 7, 'Total Other Costs', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_32(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 32 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 2, 'NO MARGIN ITEMS', xlsx.styles['rightJust2'])

def totals_34(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 34 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, 'Trucking', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_35(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 35 from the bottom of the sheet"""
    _ = section_info
    text1 = "Voyager/Custom - 10%, Guide/Lodge - 3%"

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['bgYellow4'])
    xlsx.write(adjust(row, 0), 3, 'Dealer commission', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, 'Dealer', xlsx.styles['centerJust2'])
    xlsx.write(adjust(row, 0), 6, None, xlsx.styles['bgYellow4'])
    xlsx.sheet.data_validation(adjust(row, 0), 6, adjust(row, 0), 6, {
        'validate': 'list',
        'source': DEALERS,
    })
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_36(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 36 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, None, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_37(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 37 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 3, None, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, 0.0, xlsx.styles['currency'])

def totals_38(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 38 from the bottom of the sheet"""
    _ = section_info

    formula1 = f"=SUM(I{adjust(row, -3)}:I{adjust(row, 0)})"
    value1 = 0

    xlsx.write(adjust(row, 0), 7, 'Total No Margin Items', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currencyBoldYellow'], value1)

def totals_40(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 40 from the bottom of the sheet"""
    formula1 = f"=I{row -30}+I{row - 20}+I{row - 9}+I{row - 1}"

    # value used in totals_40 and totals_43
    rate_fabrication = xlsx.settings.consumables['FABRICATION'].rate
    rate_paint = xlsx.settings.consumables['PAINT'].rate
    labor = get_labor(xlsx)
    value1 = (sum([section_info[section].value for section in section_info]) +
                  section_info['FABRICATION'].value * rate_fabrication +
                  section_info['PAINT'].value * rate_paint +
                  labor)

    xlsx.write(adjust(row, 0), 6, 'TOTAL COST OF PROJECT', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currencyBoldYellow'], value1)


def totals_42(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 42 from the bottom of the sheet"""
    _ = section_info
    text1 = "Mark up per pricing policy: "

    xlsx.sheet.set_row(adjust(row, 0), 23.85)
    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic2'])
    xlsx.write(adjust(row, 0), 3, 'Cost ', xlsx.styles['centerJust1'])
    xlsx.merge_range(adjust(row, 0), 4, adjust(row, 0), 5, 'Markup',  xlsx.styles['centerJust3'])
    xlsx.write(adjust(row, 0), 6, 'MSRP ', xlsx.styles['centerJust3'])
    xlsx.write(adjust(row, 0), 7, 'Discount ', xlsx.styles['centerJust3'])
    xlsx.write(adjust(row, 0), 9, 'Contribution Margin', xlsx.styles['centerJust4'])

def totals_43(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 43 from the bottom of the sheet"""
    # pylint: disable=too-many-locals
    dept = 'Boat and options'
    markup_1 = xlsx.settings.mark_ups[dept].markup_1
    markup_2 = xlsx.settings.mark_ups[dept].markup_2
    rate_fabrication = xlsx.settings.consumables['FABRICATION'].rate
    rate_paint = xlsx.settings.consumables['PAINT'].rate
    discount = xlsx.settings.mark_ups[dept].discount

    formula1 = (f"=I{adjust(row, -2)}-I{adjust(row, -4)}-I{adjust(row, -37)}"
                f"-I{adjust(row, -36)}-I{adjust(row, -35)}-I{adjust(row, -34)}"
                )
    # value used in totals_40
    labor = get_labor(xlsx)
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * rate_fabrication +
              section_info['PAINT'].value +
              section_info['PAINT'].value * rate_paint +
              section_info['OUTFITTING'].value +
              labor)
    formula2 = f"=D{adjust(row, 1)}/E{adjust(row, 1)}/F{adjust(row, 1)}"
    value2 = value1 / markup_1 / markup_2
    formula3 = f"=G{adjust(row, 1)}*(1-H{adjust(row, 1)})"
    value3 = value2  * (1 - discount)
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals = value3

    xlsx.write(adjust(row, 0), 2, 'Boat and options:', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(adjust(row, 0), 4, markup_1, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 5, markup_2, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, discount, xlsx.styles['percentBorderYellow'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_44(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 44 from the bottom of the sheet"""
    dept = 'Big Ticket Items'
    markup_1 = xlsx.settings.mark_ups[dept].markup_1
    markup_2 = xlsx.settings.mark_ups[dept].markup_2
    discount = xlsx.settings.mark_ups[dept].discount

    formula1 = f"=I{adjust(row, 38)}"
    value1 = section_info['BIG TICKET ITEMS'].value
    formula2 = f"=D{adjust(row, 1)}/E{adjust(row, 1)}/F{adjust(row, 1)}"
    value2 = value1 / markup_1 / markup_2
    formula3 = "=G" + str(adjust(row, 1)) + "*(1-H" + str(adjust(row, 1)) + ')'
    value3 = value2  * (1 - discount)
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}" 
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'Big Ticket Items', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(adjust(row, 0), 4, markup_1, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 5, markup_2, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, discount, xlsx.styles['percentBorderYellow'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_45(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 45 from the bottom of the sheet"""
    dept = 'OB Motors'
    discount = xlsx.settings.mark_ups[dept].discount

    formula1 = f"=I{adjust(row, -38)}"
    value1 = section_info['OUTBOARD MOTORS'].value
    formula2 = f"=Q{section_info['OUTBOARD MOTORS'].subtotal}"
    value2 = 0.0
    formula3 = f"=G{adjust(row, 1)}*(1-H{adjust(row, 1)})"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'OB Motors', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(adjust(row, 0), 4, adjust(row, 0), 5, 'See PP',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, discount, xlsx.styles['percentBorderYellow'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_46(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 46 from the bottom of the sheet"""
    dept = 'Inboard Motors & Jets'
    markup_1 = xlsx.settings.mark_ups[dept].markup_1
    markup_2 = xlsx.settings.mark_ups[dept].markup_2
    discount = xlsx.settings.mark_ups[dept].discount

    formula1 = f"=I{adjust(row, -38)}"
    value1 = section_info['INBOARD MOTORS & JETS'].value
    formula2 = f"=D{adjust(row, 1)}/E{adjust(row, 1)}/F{adjust(row, 1)}"
    value2 = value1 / markup_1 / markup_2
    formula3 = f"=G{adjust(row, 1)}*(1-H{adjust(row, 1)})"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'Inboard Motors & Jets', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(adjust(row, 0), 4, 0.85, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, discount, xlsx.styles['percentBorderYellow'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_47(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 47 from the bottom of the sheet"""
    dept = 'Trailer'
    markup_1 = xlsx.settings.mark_ups[dept].markup_1
    markup_2 = xlsx.settings.mark_ups[dept].markup_2
    discount = xlsx.settings.mark_ups[dept].discount

    formula1 = f"=I{adjust(row, -38)}"
    value1 = section_info['TRAILER'].value
    formula2 = f"=D{adjust(row, 1)}/E{adjust(row, 1)}/F{adjust(row, 1)}"
    value2 = value1 / markup_1 / markup_2
    formula3 = f"=G{adjust(row, 1)}*(1-H{adjust(row, 1)})"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{row + 1})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'Trailer', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.write(adjust(row, 0), 4, 0.8, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 5, 0.7, xlsx.styles['bgSilverBorder'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, discount, xlsx.styles['percentBorderYellow'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_48(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 48 from the bottom of the sheet"""
    _ = section_info
    formula1 = f"=I{adjust(row, 9)}"
    value1 = 0.0
    formula2 = f"=D{adjust(row, 1)}"
    value2 = 0.0
    formula3 = f"=G{adjust(row, 1)}"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'No margin items: ', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(adjust(row, 0), 4, adjust(row, 0), 5, 'none',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, 'none', xlsx.styles['centerJust5'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_49(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 49 from the bottom of the sheet"""
    if not config.hgac:
        return
    
    formula1 = f"=SUM(D{adjust(row, -5)}:D{adjust(row, -1)})*0.05"
    value1 = 0.0
    formula2 = f"=D{adjust(row, 1)}"
    value2 = 0.0
    formula3 = f"=G{adjust(row, 1)}"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'Comission: ', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(adjust(row, 0), 4, adjust(row, 0), 5, 'none',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, 'none', xlsx.styles['centerJust5'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_50(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 50 from the bottom of the sheet"""
    if not config.hgac:
        return

    formula1 = (f"=(SUM(D{adjust(row, -6)}:D{adjust(row, -2)})"
                f"+D{adjust(row, 0)})*0.02")
    value1 = 0.0
    formula2 = f"=D{adjust(row, 1)}"
    value2 = 0.0
    formula3 = f"=G{adjust(row, 1)}"
    value3 = value2
    formula4 = (f"=IF(I{adjust(row, 1)}=0,0,(I{adjust(row, 1)}"
                f"-D{adjust(row, 1)})/I{adjust(row, 1)})")
    value4 = (value3 - value1) / value3 if value3 else 0
    section_info['TOTALS'].totals += value3

    xlsx.write(adjust(row, 0), 2, 'HGAC Fee: ', xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyBordered'], value1)
    xlsx.merge_range(adjust(row, 0), 4, adjust(row, 0), 5, 'none',
                     xlsx.styles['bgSilverBorderCetner'])
    xlsx.write(adjust(row, 0), 6, formula2, xlsx.styles['currencyBordered'], value2)
    xlsx.write(adjust(row, 0), 7, 'none', xlsx.styles['centerJust5'])
    xlsx.write(adjust(row, 0), 8, formula3, xlsx.styles['currencyBordered'], value3)
    xlsx.write(adjust(row, 0), 9, formula4, xlsx.styles['percentBorder'], value4)

def totals_52(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 52 from the bottom of the sheet"""
    text1 = "Total Cost (equals total cost of project box)"
    formula1 = f"=SUM(D{adjust(row, -8)}:D{adjust(row, -1)})"
    rate_fabrication = xlsx.settings.consumables['FABRICATION'].rate
    rate_paint = xlsx.settings.consumables['PAINT'].rate
    labor = get_labor(xlsx)
    value1 = (section_info['FABRICATION'].value +
              section_info['FABRICATION'].value * rate_fabrication +
              section_info['PAINT'].value +
              section_info['PAINT'].value * rate_paint +
              section_info['OUTFITTING'].value +
              section_info['BIG TICKET ITEMS'].value +
              section_info['OUTBOARD MOTORS'].value +
              section_info['INBOARD MOTORS & JETS'].value +
              section_info['TRAILER'].value +
              labor)
    formula2 = f"=SUM(I{adjust(row, -8)}:I{adjust(row, -1)})"
    value2 = section_info['TOTALS'].totals

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 3, formula1, xlsx.styles['currencyYellow'], value1)
    xlsx.write(adjust(row, 0), 7, 'Calculated Selling Price', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula2, xlsx.styles['currency'], value2)

def totals_54(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 54 from the bottom of the sheet"""
    _ = section_info

    xlsx.write(adjust(row, 0), 6, 'SELLING PRICE', xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, None, xlsx.styles['currencyBoldYellowBorder'])

def totals_56(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 56 from the bottom of the sheet"""
    text1 = "CONTRIBUTION TO PROFIT AND OVERHEAD"
    formula1 = f"=I{adjust(row, -1)}-I{adjust(row, -15)}"
    rate_fabrication = xlsx.settings.consumables['FABRICATION'].rate
    rate_paint = xlsx.settings.consumables['PAINT'].rate
    labor = get_labor(xlsx)
    value1 = -(sum([section_info[section].value for section in section_info]) +
                    section_info['FABRICATION'].value * rate_fabrication +
                    section_info['PAINT'].value * rate_paint +
                    labor)

    xlsx.write(adjust(row, 0), 6, text1, xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['currencyBold'], value1)

def totals_58(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 58 from the bottom of the sheet"""
    _ = section_info
    text1 = "CONTRIBUTION MARGIN"
    formula1 = (f"=IF(I{adjust(row, -3)}=0,0,SUM(I{adjust(row, -3)}"
                f"-I{adjust(row, -17)})/I{adjust(row, -3)})")

    value1 =  0.0

    xlsx.write(adjust(row, 0), 6, text1, xlsx.styles['rightJust2'])
    xlsx.write(adjust(row, 0), 8, formula1, xlsx.styles['percent1'], value1)


def totals_61(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 61 from the bottom of the sheet"""
    _ = section_info
    text1 = "Pricing Policy References: "
    text2 = "Discounts / Minimum contribution margins: "

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic2'])
    xlsx.write(adjust(row, 0), 5, text2, xlsx.styles['generic2'])

def totals_62(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 62 from the bottom of the sheet"""
    _ = section_info
    text1 = "Boat MSRP = C / .61 / 0.7"
    text2 = ("Government/Commercial Discounts - Max discount 30% / "
             "Minimum margin 35%")

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, text2, xlsx.styles['generic1'])

def totals_63(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 63 from the bottom of the sheet"""
    _ = section_info
    text1 = "Options MSRP = C / .8046 / .48"
    text2 = ("Guide / Lodge Program - Commercial Markup- Max discount "
             "30% / Minimum margin 35%")

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, text2, xlsx.styles['generic1'])

def totals_64(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 64 from the bottom of the sheet"""
    _ = section_info
    text1 = "Trailers MSRP = C / 0.80 / 0.7"
    text2 = ("Guide / Lodge Program - Recreational Retail Price list- "
             "Max discount 20%")

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, text2, xlsx.styles['generic1'])

def totals_65(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 65 from the bottom of the sheet"""
    _ = section_info
    text1 = "Inboard Motors MSRP = C / 0.85 / 0.7"
    text2 = ("Non-Commercial Direct Sales - Max discount 26% / Minimum "
             "margin 38.5%")

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 5, text2, xlsx.styles['generic1'])

def totals_66(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 66 from the bottom of the sheet"""
    _ = section_info
    text1 = "Big Ticket Items MSRP = C / (range from 0.80 – 0.85) / 0.7"

    xlsx.write(adjust(row, 0), 2, text1, xlsx.styles['generic1'])
    xlsx.sheet.write_rich_string(
        adjust(row, 0), 5, xlsx.styles['generic1'], 'Voyager - ',
        xlsx.styles['red'],
        "SEE MIKE ON ALL VOYAGER OR CUSTOM DEALER REFERRAL PRICING")

def totals_67(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 67 from the bottom of the sheet"""
    _ = section_info
    text1 = "GSA Pricing"
    text2 = "1 - 2 boats"
    text3 = "30% discount"

    xlsx.write(adjust(row, 0), 5, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 7, text2, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text3, xlsx.styles['generic1'])

def totals_68(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 68 from the bottom of the sheet"""
    _ = section_info
    text1 = "3 boats"
    text2 = "30.5% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])

def totals_69(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 69 from the bottom of the sheet"""
    _ = section_info
    text1 = "4 boats"
    text2 = "31% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])

def totals_70(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 70 from the bottom of the sheet"""
    _ = section_info
    text1 = "5 boats"
    text2 = "31.5% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])

def totals_71(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 71 from the bottom of the sheet"""
    _ = section_info
    text1 = "6 - 10 bts"
    text2 = "32% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])

def totals_72(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 72 from the bottom of the sheet"""
    _ = section_info
    text1 = "11 - 20 bts"
    text2 = "32.25% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])

def totals_73(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 73 from the bottom of the sheet"""
    _ = section_info
    text1 = "20+ boats"
    text2 = "32.5% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 8, text2, xlsx.styles['generic1'])


def totals_74(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 74 from the bottom of the sheet"""
    _ = section_info
    text1 = "Outboard motors"
    text2 = "Government agencies - 15% discount"

    xlsx.write(adjust(row, 0), 5, text1, xlsx.styles['generic1'])
    xlsx.write(adjust(row, 0), 7, text2, xlsx.styles['generic1'])

def totals_75(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 75 from the bottom of the sheet"""
    _ = section_info
    text1 = "(approval req. for greater discount, no more than addl. 3%)"
    text2 = "GSA Pricing - 18% discount"

    xlsx.merge_range(adjust(row, 0), 5, row + 2, 6, text1, xlsx.styles['italicsNote'])
    xlsx.write(adjust(row, 0), 7, text2, xlsx.styles['generic1'])

def totals_76(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 76 from the bottom of the sheet"""
    _ = section_info
    text1 = "Guides / Lodges - 10% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])

def totals_77(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 77 from the bottom of the sheet"""
    _ = section_info
    text1 = "Commercial sales - 5% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])

def totals_78(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 78 from the bottom of the sheet"""
    _ = section_info
    text1 = "Voyager - 5% discount"

    xlsx.write(adjust(row, 0), 7, text1, xlsx.styles['generic1'])

def totals_80(xlsx: XlsxBom, section_info: dict[str, SectionInfo],
              row: int)-> None:
    """fill out line at row 80 from the bottom of the sheet"""
    _ = section_info
    style1 = xlsx.styles['generic1']
    style2 = xlsx.styles['generic2']

    xlsx.write(adjust(row, 0), 2,
               'Cost estimate check list - complete prior to sending quote i'
               'or submitting bid', style2)
    xlsx.write(adjust(row, 1), 2,
               'Verify all formulas are correct and all items are included '
               'in cost total', style1)
    xlsx.write(adjust(row, 2), 2,
               'Verify aluminum calculated with total lbs included. Include '
               'metal costing sheet separate if completed', style1)
    xlsx.write(adjust(row, 3), 2,
               'Verify paint costing equals paint description', style1)
    xlsx.write(adjust(row, 4), 2,
               'Cost estimate includes all components on sales quote', style1)
    xlsx.write(adjust(row, 5), 2,
               'Pricing policy discounts and minimum margins are met', style1)
    xlsx.write(adjust(row, 6), 2,
               'Vendor quotes received and included in costing folder', style1)
    xlsx.write(adjust(row, 7), 2,
               'Labor hours reviewed and correct to best knowledge of project',
               style1)
    xlsx.write(adjust(row, 8), 2,
               'Name of peer who reviewed prior to submission to customer',
               style1)
    xlsx.write(adjust(row, 9), 2,
               'Customer signed sales quotation', style1)
    xlsx.write(adjust(row, 10), 2,
               'Customer provided terms and conditions, including payment '
               'schedule', style1)


def generate_totals(xlsx: XlsxBom,
                    section_info: dict[str, SectionInfo]) -> None:
    """generate header on costing sheet"""
    skip = [10, 11, 18, 21, 22, 24, 31, 33,
            39, 41, 51, 53, 55, 57, 59, 60, 79]
    # pylint: disable=unused-variable
    config.offset = section_info['TRAILER'].subtotal + 2
    for row in range(0, 81): # 81
        if row in skip:
            continue
        # pylint: disable=eval-used
        eval(f"totals_{row:02}(xlsx, section_info, row + config.offset)")
    totals_column_b(xlsx, section_info)

if __name__ == "__main__":
    pass
