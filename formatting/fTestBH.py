
import odspy
from odspy import Document, Sheet, Cell
import sys
import subprocess as sub
from typing import List, Dict, Set, Any, Optional
# from openpyxl import load_workbook
from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TableColumnProperties, TableCellProperties
from odf.number import NumberStyle, Number

def format_ods_cell(cell, bold=False, cell_type=None, date_format=None, currency_format=None):
    """
    Format an ODS cell using odspy.
    
    Parameters:
    -----------
    cell : odspy cell object
        The cell to format
    bold : bool, optional
        Whether to make the font bold (default: False)
    cell_type : str, optional
        Type to change cell to: 'date', 'currency', or None for no change
    date_format : str, optional
        Date format string (e.g., 'YYYY-MM-DD', 'MM/DD/YYYY')
        If None, uses default date format
    currency_format : str, optional
        Currency format string (e.g., '$#,##0.00', '€#,##0.00')
        If None, uses default currency format
    
    Returns:
    --------
    None (modifies cell in place)
    
    Example:
    --------
    # Make cell bold
    format_ods_cell(cell, bold=True)
    
    # Change to date format
    format_ods_cell(cell, cell_type='date', date_format='MM/DD/YYYY')
    
    # Change to currency with bold
    format_ods_cell(cell, bold=True, cell_type='currency', currency_format='$#,##0.00')
    """
    
    # Apply bold formatting
    if bold:
        cell.set_style_property('fo:font-weight', 'bold')
    
    # Change cell type and format
    if cell_type == 'date':
        # Set cell type to date
        cell.set_attribute('office:value-type', 'date')
        
        # Apply date format if specified
        if date_format:
            # Create a date style with the specified format
            style_name = f"date_style_{id(cell)}"  # Unique style name
            
            # Common date format mappings
            format_mapping = {
                'YYYY-MM-DD': '%Y-%m-%d',
                'MM/DD/YYYY': '%m/%d/%Y',
                'DD/MM/YYYY': '%d/%m/%Y',
                'YYYY/MM/DD': '%Y/%m/%d',
                'MMM DD, YYYY': '%b %d, %Y',
                'MMMM DD, YYYY': '%B %d, %Y'
            }
            
            # Use mapped format or the provided format directly
            format_code = format_mapping.get(date_format, date_format)
            cell.set_style_property('number:format-source', 'language')
            cell.set_style_property('style:data-style-name', style_name)
    
    elif cell_type == 'currency':
        # Set cell type to currency
        cell.set_attribute('office:value-type', 'currency')
        cell.set_attribute('office:currency', 'USD')  # Default to USD
        
        # Apply currency format if specified
        if currency_format:
            style_name = f"currency_style_{id(cell)}"  # Unique style name
            
            # Extract currency symbol if present
            if currency_format.startswith('$'):
                cell.set_attribute('office:currency', 'USD')
            elif currency_format.startswith('€'):
                cell.set_attribute('office:currency', 'EUR')
            elif currency_format.startswith('£'):
                cell.set_attribute('office:currency', 'GBP')
            
            cell.set_style_property('style:data-style-name', style_name)
        else:
            # Default currency format
            cell.set_style_property('style:data-style-name', 'currency')


# Alternative simpler version if the above doesn't work with your odspy version
def format_ods_cell_simple(cell, bold=False, cell_type=None):
    """
    Simplified version that focuses on basic formatting.
    
    Parameters:
    -----------
    cell : odspy cell object
        The cell to format
    bold : bool, optional
        Whether to make the font bold
    cell_type : str, optional
        'date' or 'currency' to change cell type
    """
    
    # Apply bold formatting
    if bold:
        try:
            cell.set_style_property('fo:font-weight', 'bold')
        except:
            # Alternative method if the above doesn't work
            cell.style.font_weight = 'bold'
    
    # Change cell value type
    if cell_type == 'date':
        try:
            cell.set_attribute('office:value-type', 'date')
        except:
            cell.value_type = 'date'
            
    elif cell_type == 'currency':
        try:
            cell.set_attribute('office:value-type', 'currency')
            cell.set_attribute('office:currency', 'USD')
        except:
            cell.value_type = 'currency'


def main():


    # Preliminaries
    args = sys.argv
    if len(args) ==2:
        odsname = sys.argv[1]
    else:
        odsname = mc.default_ODS_name

    print("Input file: ", odsname)
    if not odsname.endswith('.ods'):
        print(f" illegal filetype {odsname}  (must be .ods)")
        quit()


    doc = load(odsname)

    tables = doc.spreadsheet.getElementsByType(Table)  #(I think "tables" are "Tabs")

    if not tables:
        print("Error: No tables found in ODS document.")
        return False

    table = tables[0]
    rows = table.getElementsByType(TableRow)

    if len(rows) < 2:
        print("No data rows to format.")
        return True
