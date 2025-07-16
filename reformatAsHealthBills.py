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
import formatting.formatODS as fods



        #
        # self.streamlined_headers = [
        #     "Claim number", "D.O.S", "Received", "Processed",
        #     "Status", "Provider", "Patient", "Billed",
        #     "Your Rate", "We Paid", "You owe",
        #     "Deductible", "Copay", "Coins.", "Other Ins.",
        #     "Medication", "Prescription", "NDC number", "Day supply",
        #     "Quantity", "Pharmacy", "Pharm. number"
        # ]

def reformatODSclaims(doc,mobj):
    #   0,     5,      10         0        6
    #  DOS, Provider, Amt Owed, Claim#, Patient
    oldcolindeces = [1, 5, 10,  0, 6]
    newcolindeces = [0, 2,  5, 11, 12]
    fmts          = ['d','s','$','s','s']

    nColsOut = len(oldcolindeces)
    #
    #  Get all rows and select the cols we want
    #
    tables = doc.spreadsheet.getElementsByType(Table)  #(I think "tables" are "Tabs")
    rows = tables[0].getElementsByType(TableRow)

    new_rows = []
    for row_idx, row in enumerate(rows):
        oldRowCells = row.getElementsByType(TableCell)

        print(f'got {len(oldRowCells)} cells from row {row_idx}.')
        if len(oldRowCells) < nColsOut:
            continue

        # eliminate blank lines and header
        if str(oldRowCells[0]).strip() != '' and str(oldRowCells[0]).strip().lower() != 'claim number':
            emptyCell = TableCell()
            emptyCell.addElement(P(text='')) # start with row of blank cells

            new_TableRowList = []
            for i in range(13):    # fill new row list with empty cells
                tcell = TableCell()
                tcell.addElement(P(text=''))
                print(f'----> adding cell {i} to row LIST: [{tcell}]')
                new_TableRowList.append(tcell)
            #
            # now replace key columns with real data
            for i,old_idx in enumerate(oldcolindeces):  # reorder the relevant columns
                #  todo reformat at least to floats for $ amts.
                if old_idx >= 0:
                    if   fmts[i] == 'd':
                        fods.format_ods_cell(doc, oldRowCells[old_idx],cell_type='date')  # in-place reformat
                    elif fmts[i] == '$':
                        fods.format_ods_cell(doc, oldRowCells[old_idx],cell_type='currency')

                    new_TableRowList[newcolindeces[i]] = oldRowCells[old_idx]

            # Build the new row
            new_TableRow = TableRow()
            i=0
            for c in new_TableRowList:
                print(f'----> adding cell {i} to row: ', c)
                new_TableRow.addElement(c)
                i+=1

            # Store the actual row object with its date for sorting
            new_rows.append(new_TableRow)
        else:
            print(' . . .   blank or header row')
    return new_rows

def main():
    # Preliminaries

    #   Get ods input file
    args = sys.argv
    if len(args) ==2:
        odsname = sys.argv[1]
    else:
        odsname = mc.default_ODS_name

    print("Input file: ", odsname)
    if not odsname.endswith('.ods'):
        print(f" illegal filetype {odsname}  (must be .ods)")
        quit()


    # create a merge object (we won't actually merge)
    merge = mc.ClaimsDataMerger(odsname) # maybe don't need this?

    odsname = merge.ods_filename

    doc = load(odsname)
    print('document loaded!\n\n')

    #
    #   Do the reformatting
    #
    newrows = reformatODSclaims(doc,mc)


    #
    #   Make the new output spreadsheet
    #       (for manual pasting into HealthBills20xx.ods)
    #
    newsheetName = 'NewSheetforHealthBills.ods'
    print('Starting new document for output ', newsheetName)
    newmg = mc.ClaimsDataMerger(newsheetName)
    newdoc = OpenDocumentSpreadsheet()
    newdoc.name = newmg.ods_filename

    # build the output header
    outheaders = ['Date of Serv','Date of Bill','Provider', 'Description', 'Organiz.', 'Responsib.', 'Amt', 'payment', 'E/P', 'Date', 'Deductable','Claim Number','Patient']

    hrow = TableRow()
    for h in outheaders:
        cell = TableCell()
        cell.addElement(P(text=h))
        hrow.addElement(cell)

    newtable = Table(name='claimUpdates')

    # build the new table
    newtable.addElement(hrow)
    for row in newrows:
        newtable.addElement(row)

    newdoc.spreadsheet.addElement(newtable)
    newmg.save_ods(newdoc)

if __name__=='__main__':
    main()
