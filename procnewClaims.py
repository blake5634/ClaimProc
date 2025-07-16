#!/usr/bin/env python3
"""
Medical Claims Data Merger
Version: 2.1
Created: 2025-07-04
Last Updated: 2025-07-04

Reads medical claim data from Excel files and merges into LibreOffice Calc (default_ODS_name)
without duplicating existing entries while preserving formatting.
"""

import os
import sys
from typing import List, Dict, Set, Any, Optional
from openpyxl import load_workbook
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableColumnProperties, TableCellProperties
from odf.number import NumberStyle, Number

import mergeClass as mc

def main():
    """Main function to run the claims merger."""
    if len(sys.argv) < 2:
        print("Medical Claims Data Merger v2.1 (thanks Claude.ai)")
        print("Usage: python claims_merger.py <excel_file1> [excel_file2] ...")
        print("Example: python claims_merger.py claims_2024.xlsx")
        return

    merger = mc.ClaimsDataMerger(mc.default_ODS_name)
    total_added = 0

    for excel_file in sys.argv[1:]:
        added = merger.merge_excel_to_ods(excel_file)
        total_added += added
        print("Processed excel file")
        # try:
        #     added = merger.merge_excel_to_ods(excel_file)
        #     total_added += added
        #     print(f"Processed: {excel_file}")
        # except Exception as e:
        #     print(f"Error processing (merge_excel_to_ods()) {excel_file}: {str(e)}")

    print(f"\nTotal claims added: {total_added}")


if __name__ == "__main__":
    main()

