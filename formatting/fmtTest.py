from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TextProperties
from odf.number import DateStyle, CurrencyStyle
from odf import teletype

import formatODS as fods

def debug_cell_structure(cell, label):
    """
    Debug function to inspect cell structure
    """
    print(f"\n=== {label} ===")
    print(f"Cell tag: {cell.tagName}")
    print(f"Attributes:")

    # Handle attributes safely - some might be namespace tuples
    for attr in cell.attributes.keys():
        try:
            if isinstance(attr, tuple):
                # This is a namespace tuple like ('urn:...', 'value-type')
                attr_name = attr[1]  # Get the local name
                value = cell.attributes[attr]
                print(f"  {attr_name} (ns): {value}")
            else:
                # Simple attribute name
                value = cell.getAttribute(attr)
                print(f"  {attr}: {value}")
        except Exception as e:
            print(f"  {attr}: <error getting value: {e}>")

    # Try to get some specific attributes we care about
    try:
        valuetype = cell.getAttribute('valuetype')
        if valuetype:
            print(f"  valuetype (direct): {valuetype}")
    except:
        pass

    try:
        datevalue = cell.getAttribute('datevalue')
        if datevalue:
            print(f"  datevalue (direct): {datevalue}")
    except:
        pass

    print(f"Child nodes:")
    for i, child in enumerate(cell.childNodes):
        print(f"  Child {i}: {child.tagName if hasattr(child, 'tagName') else type(child)}")
        if hasattr(child, 'tagName') and child.tagName == 'text:p':
            print(f"    Text content: '{teletype.extractText(child)}'")

    print(f"Extracted text: '{teletype.extractText(cell)}'")


def test_cell_formatting():
    """
    Test the cell formatting functions on fmtTest.ods
    """
    # Load the test spreadsheet
    doc = load("fmtTest.ods")

    # Get the first table (sheet)
    tables = doc.getElementsByType(Table)
    if not tables:
        print("No tables found in the document")
        return

    table = tables[0]
    rows = table.getElementsByType(TableRow)

    if len(rows) < 3:
        print(f"Expected at least 3 rows, found {len(rows)}")
        return

    # Get cells B1, B2, B3 (second column of first three rows)
    cell_b1 = None
    cell_b2 = None
    cell_b3 = None

    # Get B1 (row 0, column 1)
    if len(rows) > 0:
        cells_row1 = rows[0].getElementsByType(TableCell)
        if len(cells_row1) > 1:
            cell_b1 = cells_row1[1]

    # Get B2 (row 1, column 1)
    if len(rows) > 1:
        cells_row2 = rows[1].getElementsByType(TableCell)
        if len(cells_row2) > 1:
            cell_b2 = cells_row2[1]

    # Get B3 (row 2, column 1)
    if len(rows) > 2:
        cells_row3 = rows[2].getElementsByType(TableCell)
        if len(cells_row3) > 1:
            cell_b3 = cells_row3[1]

    print("BEFORE FORMATTING:")
    # if cell_b1:
    #     debug_cell_structure(cell_b1, "B1 BEFORE")
    # if cell_b2:
    #     debug_cell_structure(cell_b2, "B2 BEFORE")
    # if cell_b3:
    #     debug_cell_structure(cell_b3, "B3 BEFORE")

    # Apply formatting changes
    print("\n" + "="*50)
    print("APPLYING FORMATTING...")

    # B1: Change to date
    if cell_b1:
        fods.format_ods_cell(doc, cell_b1, cell_type='date')
        print("B1: Applied date formatting")

    # B2: Change to currency
    if cell_b2:
        fods.format_ods_cell(doc, cell_b2, cell_type='currency')
        print("B2: Applied currency formatting")

    # B3: Make bold
    if cell_b3:
        fods.format_ods_cell(doc, cell_b3, bold=True)
        print("B3: Applied bold formatting")

    print("\nAFTER FORMATTING:")
    # if cell_b1:
    #     debug_cell_structure(cell_b1, "B1 AFTER")
    # if cell_b2:
    #     debug_cell_structure(cell_b2, "B2 AFTER")
    # if cell_b3:
    #     debug_cell_structure(cell_b3, "B3 AFTER")

    # Save the modified spreadsheet
    doc.save("fmtTest_modified.ods")
    print("\nSaved modified spreadsheet as 'fmtTest_modified.ods'")
    print("Test completed successfully")


def inspect_document_styles(doc):
    """
    Inspect what styles are available in the document
    """
    print("=== Document Styles Inspection ===")

    # Check automatic styles
    auto_styles = doc.automaticstyles.getElementsByType(Style)
    print(f"Automatic styles found: {len(auto_styles)}")
    for style in auto_styles[:5]:  # Show first 5
        print(f"  - {style.getAttribute('name')} (family: {style.getAttribute('family')})")

    # Check regular styles
    styles = doc.styles.getElementsByType(Style)
    print(f"Regular styles found: {len(styles)}")
    for style in styles[:5]:  # Show first 5
        print(f"  - {style.getAttribute('name')} (family: {style.getAttribute('family')})")

    # Check number styles
    date_styles = doc.styles.getElementsByType(DateStyle)
    currency_styles = doc.styles.getElementsByType(CurrencyStyle)
    number_styles = doc.styles.getElementsByType(NumberStyle)

    print(f"Date styles found: {len(date_styles)}")
    for style in date_styles[:3]:
        print(f"  - {style.getAttribute('name')}")

    print(f"Currency styles found: {len(currency_styles)}")
    for style in currency_styles[:3]:
        print(f"  - {style.getAttribute('name')}")

    print(f"Number styles found: {len(number_styles)}")
    for style in number_styles[:3]:
        print(f"  - {style.getAttribute('name')}")


def simple_cell_access_test():
    """
    Simple test to verify we can access cells correctly
    """
    try:
        doc = load("fmtTest.ods")
        tables = doc.getElementsByType(Table)
        print(f"Found {len(tables)} table(s)")

        if tables:
            table = tables[0]
            rows = table.getElementsByType(TableRow)
            print(f"Found {len(rows)} row(s)")

            for i, row in enumerate(rows[:3]):  # First 3 rows
                cells = row.getElementsByType(TableCell)
                print(f"Row {i+1} has {len(cells)} cell(s)")
                for j, cell in enumerate(cells[:2]):  # First 2 columns
                    text = teletype.extractText(cell)
                    print(f"  Cell {chr(65+j)}{i+1}: '{text}'")

    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    print("Running document styles inspection...")
    try:
        doc = load("fmtTest.ods")
        inspect_document_styles(doc)
    except Exception as e:
        print(f"Style inspection failed: {e}")

    print("\nRunning simple access test...")
    simple_cell_access_test()

    print("\n" + "="*50)
    print("Running formatting test...")
    try:
        test_cell_formatting()
    except Exception as e:
        print(f"Formatting test failed: {e}")
        import traceback
        traceback.print_exc()

