from odf.opendocument import OpenDocumentSpreadsheet, load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TextProperties
from odf.number import DateStyle, CurrencyStyle
from odf import teletype


def format_ods_cell(doc, cell, bold=False, cell_type=None):
    """
    Format an ODS cell using odfpy.

    Parameters:
    -----------
    doc : OpenDocumentSpreadsheet
        The document object
    cell : TableCell
        The cell to format
    bold : bool, optional
        Whether to make the font bold (default: False)
    cell_type : str, optional
        Type to change cell to: 'date', 'currency', or None for no change

    Returns:
    --------
    None (modifies cell in place)
    """

    # Handle bold formatting
    if bold:
        style_name = f"bold_style_{id(cell)}"
        cell_style = Style(name=style_name, family="table-cell")
        text_props = TextProperties(fontweight="bold")
        cell_style.addElement(text_props)
        doc.styles.addElement(cell_style)
        cell.setAttribute('stylename', style_name)

    # Handle cell type changes
    if cell_type == 'date':
        # First remove any existing value-type attributes
        try:
            cell.removeAttribute('valuetype')
        except:
            pass
        try:
            # Remove the namespace version too
            for attr_key in list(cell.attributes.keys()):
                if isinstance(attr_key, tuple) and attr_key[1] == 'value-type':
                    del cell.attributes[attr_key]
        except:
            pass

        # Now set the new type
        cell.setAttribute('valuetype', 'date')

        # Try to parse existing text and set date value
        current_text = teletype.extractText(cell)
        if current_text:
            # Remove leading single quote if present (forces text mode in LibreOffice)
            if current_text.startswith("'"):
                current_text = current_text[1:]

            try:
                from datetime import datetime
                date_obj = None

                # Try common date formats - prioritize US format mm/dd/yy
                date_formats = ['%m/%d/%y', '%m/%d/%Y', '%m-%d-%y', '%m-%d-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d']
                for fmt in date_formats:
                    try:
                        date_obj = datetime.strptime(current_text.strip(), fmt)
                        break
                    except ValueError:
                        continue

                if date_obj:
                    iso_date = date_obj.strftime('%Y-%m-%d')
                    cell.setAttribute('datevalue', iso_date)

                    # Don't set display text - let LibreOffice handle it with default formatting
                    # Clear old content but don't add new P element
                    for child in list(cell.childNodes):
                        cell.removeChild(child)

                else:
                    # Default to today's date
                    from datetime import date
                    today = date.today()
                    today_iso = today.strftime('%Y-%m-%d')
                    cell.setAttribute('datevalue', today_iso)

                    # Clear old content
                    for child in list(cell.childNodes):
                        cell.removeChild(child)

            except Exception as e:
                print(f"Warning: Could not parse date from '{current_text}': {e}")

    elif cell_type == 'currency':
        # First remove any existing value-type attributes
        try:
            cell.removeAttribute('valuetype')
        except:
            pass
        try:
            # Remove the namespace version too
            for attr_key in list(cell.attributes.keys()):
                if isinstance(attr_key, tuple) and attr_key[1] == 'value-type':
                    del cell.attributes[attr_key]
        except:
            pass

        # Now set the new type
        cell.setAttribute('valuetype', 'currency')
        cell.setAttribute('currency', 'USD')

        # Try to parse existing text as currency value
        current_text = teletype.extractText(cell)
        if current_text:
            # Remove leading single quote if present (forces text mode in LibreOffice)
            if current_text.startswith("'"):
                current_text = current_text[1:]

            try:
                import re
                numeric_text = re.sub(r'[^\d.-]', '', current_text)
                if numeric_text:
                    value = float(numeric_text)
                    cell.setAttribute('value', str(value))

                    # Don't set display text - let LibreOffice handle currency formatting
                    # Clear old content
                    for child in list(cell.childNodes):
                        cell.removeChild(child)

                else:
                    cell.setAttribute('value', '0')

                    # Clear old content
                    for child in list(cell.childNodes):
                        cell.removeChild(child)

            except Exception as e:
                print(f"Warning: Could not parse currency from '{current_text}': {e}")
                cell.setAttribute('value', '0')

