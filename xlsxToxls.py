from pathlib import Path
import openpyxl
import os

def main():

    root_dir = os.path.dirname(__file__)

    os.chdir(os.path.abspath(root_dir))
    pdir = Path(root_dir)
    # Creates a list of files that ends in '.xlsx'
    xlsx_filelist = [filename for filename in pdir.iterdir() if filename.suffix == ".xlsx"]
    # Creates a list of files that ends in '.xls'
    xls_filelist = [filename for filename in pdir.iterdir() if filename.suffix == ".xls"]

    for xlsxfile in xlsx_filelist:
        # Pre-determines the new file name
        output = f"{xlsxfile.name.split('.')[0]}.xls"
        # Prevents duplicate files from being created
        # after first successful conversion
        if output not in xls_filelist:
            workbook = openpyxl.load_workbook(xlsxfile)
            workbook.save(output)

if __name__ == '__main__':
    main()