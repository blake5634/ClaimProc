#!/usr/bin/env python3
"""
Simple Excel File Cleaner using LibreOffice
Usage: python clean_excel.py input_file.xlsx
Creates: input_file_cleaned.xlsx
"""

import subprocess
import sys
import os
import time

def clean_excel_file(input_file):
    """Convert Excel file using LibreOffice to fix formatting issues."""
    
    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return False
    
    # Create output filename
    base_name = os.path.splitext(input_file)[0]
    output_file = f"{base_name}_cleaned.xlsx"
    
    print(f"Converting '{input_file}' using LibreOffice...")
    
    try:
        # Use LibreOffice in headless mode to convert
        cmd = [
            'libreoffice', '--headless', '--convert-to', 'xlsx',
            '--outdir', os.path.dirname(input_file) or '.',
            input_file
        ]
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        
        if result.returncode != 0:
            print(f"LibreOffice failed: {result.stderr}")
            return False
        
        # Wait for file to be written
        time.sleep(2)
        
        # LibreOffice creates file with same base name
        lo_output = os.path.join(
            os.path.dirname(input_file) or '.',
            os.path.splitext(os.path.basename(input_file))[0] + '.xlsx'
        )
        
        if os.path.exists(lo_output) and lo_output != input_file:
            # Rename to cleaned version
            if os.path.exists(output_file):
                os.remove(output_file)
            os.rename(lo_output, output_file)
            print(f"Success! Created: {output_file}")
            return True
        elif os.path.exists(input_file):
            # LibreOffice might have updated the original file
            import shutil
            shutil.copy2(input_file, output_file)
            print(f"Success! Created: {output_file}")
            return True
        else:
            print("Error: Could not find LibreOffice output")
            return False
            
    except subprocess.TimeoutExpired:
        print("Error: LibreOffice conversion timed out")
        return False
    except FileNotFoundError:
        print("Error: LibreOffice not found. Please install LibreOffice.")
        return False
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

def main():
    if len(sys.argv) != 2:
        print("Usage: python clean_excel.py input_file.xlsx")
        print("Creates a cleaned version: input_file_cleaned.xlsx")
        return
    
    input_file = sys.argv[1]
    
    if clean_excel_file(input_file):
        print("File cleaning completed successfully!")
    else:
        print("File cleaning failed!")

if __name__ == "__main__":
    main()
