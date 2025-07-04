import os
import sys
import subprocess as sub
from typing import List, Dict, Set, Any, Optional
from openpyxl import load_workbook
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableColumnProperties, TableCellProperties
from odf.number import NumberStyle, Number

import mergeClass as mc


def main():
    # Preliminaries
    args = sys.argv
    if len(args) ==2:
        odsname = sys.argv[1]
    else:
        odsname = "Data.ods"

    print("Input file: ", odsname)
    if not odsname.endswith('.ods'):
        print(f" illegal filetype {odsname}  (must be .ods)")
        quit()

    sub.run(['cp', odsname, odsname+'BACKUP'])

    # Read, sort, writeback out
    merge = mc.ClaimsDataMerger("Data.ods")

    odsname = merge.ods_filename

    doc = load(odsname)
# Verify the save by reading back

    print('document loaded!')

    sortOK = merge.sort_ods_by_date_of_service()

    if sortOK:
        print(' document sorted')
    else:
        print(' doc sort failed')
        quit()

    merge.save_ods(doc)

if __name__=='__main__':
    main()
