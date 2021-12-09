#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
"""
Generate Costing Sheet Headers at top of sheet
"""
from .costing_data import XlsxBom

# WRITING HEADING FUNCTIONS ===================================================

def generate_header(xlsx: XlsxBom) -> None:
    """generate header on costing sheet"""

    xlsx.sheet.set_row(1, 26.25)

    xlsx.write('B2', 'Customer:', xlsx.styles['headingCustomer1'])
    xlsx.write('C2', None, xlsx.styles['headingCustomer2'])
    xlsx.write('G2', 'Salesperson:')
    xlsx.merge_range('H2:I2', None, xlsx.styles['bgYellow2'])

    xlsx.write('B4', 'Boat Model:')
    xlsx.write('C4', xlsx.file_name_info['folder'], xlsx.styles['bgYellow1'])
    xlsx.write('B5', 'Beam:')
    xlsx.write('C5', xlsx.bom.beam, xlsx.styles['bgYellow1'])
    xlsx.write('B6', 'Length:')
    xlsx.write('C6', xlsx.file_name_info['size'], xlsx.styles['bgYellow1'])

    xlsx.write('H4', 'Original Date Quoted:', xlsx.styles['rightJust1'])
    xlsx.write('I4', None, xlsx.styles['bgYellow1'])

    xlsx.merge_range('E5:G5',
                     'Indicate changes here',
                     xlsx.styles['bgGreen2']
                    )
    xlsx.merge_range('E6:G6',
                     'Indicate changes here',
                     xlsx.styles['bgPurple2']
                    )

    xlsx.merge_range('E7:G7',
                     'Indicate changes here',
                     xlsx.styles['bgCyan2']
                    )
    xlsx.merge_range('E8:G8',
                     'Indicate changes here',
                     xlsx.styles['bgOrange2']
                    )

    xlsx.write('H5', 'Rev1', xlsx.styles['rightJust1'])
    xlsx.write('H6', 'Rev2', xlsx.styles['rightJust1'])
    xlsx.write('H7', 'Rev3', xlsx.styles['rightJust1'])
    xlsx.write('H8', 'Rev4', xlsx.styles['rightJust1'])

    xlsx.write('I5', None, xlsx.styles['bgGreen1'])
    xlsx.write('I6', None, xlsx.styles['bgPurple1'])
    xlsx.write('I7', None, xlsx.styles['bgCyan1'])
    xlsx.write('I8', None, xlsx.styles['bgOrange1'])


if __name__ == "__main__":
    pass
