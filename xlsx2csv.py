#!/usr/bin/env python3
"""
Convert XLSX to CSV using LibreOffice
Usage: python xlsx_to_csv.py file.xlsx
Creates: file.csv
"""

import subprocess
import sys
import os

def xlsx_to_csv(xlsx_file):


    """Convert XLSX to CSV using LibreOffice."""
    print(f"DEBUG: Current working directory: {os.getcwd()}")
    print(f"DEBUG: Input file: {xlsx_file}")
    print(f"DEBUG: Input file absolute path: {os.path.abspath(xlsx_file)}")


    """Convert XLSX to CSV using LibreOffice."""
    if not os.path.exists(xlsx_file):
        print(f"Error: {xlsx_file} not found")
        return False
    
    try:
        result = subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'csv',
            '--outdir', os.path.dirname(xlsx_file) or '.',
            xlsx_file
        ], check=True, capture_output=True)

        print(f"LibreOffice stdout: {result.stdout}")
        if result.stderr:
            print(f"LibreOffice stderr: {result.stderr}")

        # Wait for process to fully complete
        result.wait() if hasattr(result, 'wait') else None

        # Give filesystem time to sync
        import time
        time.sleep(1)

        csv_file = os.path.splitext(xlsx_file)[0] + '.csv'

        # Verify file actually exists before returning
        if os.path.exists(csv_file):
            print(f"Created: {csv_file}")
            return True
        else:
            print(f"Error: Expected {csv_file} was not created")
            return False

    except subprocess.TimeoutExpired:
        print("Error: LibreOffice conversion timed out")
        return False
    except:
        print("Error: LibreOffice conversion failed")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python xlsx_to_csv.py file.xlsx")
    else:
        xlsx_to_csv(sys.argv[1])
