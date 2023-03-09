# xlsx_to_xls

Converts xlsx files to xls in the same directory as the Python Script.
  - Will not create duplicate xls file if there's already one in the folder from previous conversion.

[SaveAs Documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.saveas)

[FileFormat Documentation](https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat)

------------------------------------------
* win32com method uses Excel application installed on the pc and converts .xlsx file to .xls
  - **Pros**
      - Since it uses Excel application directly to convert, no error message on launch of newly generated .xls file
  - **Cons**
      - It requires Excel to be installed. (Must be newer than 2003) 
      - Script can run with Office03 **but** excel file will break. (Can be done manually by the user)
     
------------------------------------------

* openpyxl method uses openpyxl module to convert the .xlsx file to .xls
  - **Pros**
      - Does not require Excel to be installed
  - **Cons**
      - Opening the file prompts the user with an error message and the user must select "Yes" to proceed.
