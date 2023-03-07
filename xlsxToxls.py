from pathlib import Path
import openpyxl
import os

def main():

    root_dir = os.path.dirname(__file__)

    os.chdir(os.path.abspath(root_dir))
    pdir = Path(root_dir)
    filelist = [filename for filename in pdir.iterdir() if filename.suffix == ".xlsx"]


    for xlsxfile in filelist:
        workbook = openpyxl.load_workbook(xlsxfile)
        outfile = f"{xlsxfile.name.split('.')[0]}.xls"
        workbook.save(outfile)

if __name__ == '__main__':
    main()