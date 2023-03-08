from pathlib import Path
import openpyxl # pip install openpyxl
import os

def main():
    root_dir = os.path.dirname(__file__)

    # Creates a list of files that ends in '.xlsx'
    xlsx_filelist = [filename for filename in os.listdir(root_dir) if filename.endswith(".xlsx")]
    # Creates a list of files that ends in '.xls'
    xls_filelist = [filename for filename in os.listdir(root_dir) if filename.endswith(".xls")]
    
    for xlsxfile in xlsx_filelist:
        # Pre-determines the new file name
        output = f"{xlsxfile.split('.')[0]}.xls"
        # Prevents duplicate files from being created
        # after first successful conversion
        if output not in xls_filelist:
            workbook = openpyxl.load_workbook(xlsxfile)
            workbook.save(output)

if __name__ == '__main__':
    main()