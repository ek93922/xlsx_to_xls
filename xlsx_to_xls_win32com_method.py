import win32com.client as win32 # pip install pywin32
import os

def main():
    # Set root directory
    root_dir = os.path.dirname(__file__)

    # Creates a list of files that ends in '.xlsx'
    xlsx_filelist = [filename for filename in os.listdir(root_dir) if filename.endswith(".xlsx")]
    # Creates a list of files that ends in '.xls'
    xls_filelist = [filename for filename in os.listdir(root_dir) if filename.endswith(".xls")]


    for file in xlsx_filelist:
        file_dir = os.path.join(root_dir, file)
        # Define name of soon to be output file
        output_name = f"{file.split('.')[0]}.xls"
        # Checks if the output name already exists in the directory.
        # Prevents duplicate files from being created.
        if output_name not in xls_filelist:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            output_dir = os.path.join(root_dir, output_name)
            wb = excel.Workbooks.Open(file_dir)
            wb.SaveAs(output_dir, FileFormat = 56) # FileFormat [56 is xls, 51 is xlsx]
            wb.Close()
            excel.Application.Quit()

if __name__ == '__main__':
    main()
