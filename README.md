# xlsx_to_xls

Converts xlsx files to xls in the same directory as the Python Script.
  - Will not create duplicate xls file if there's already one in the folder from previous conversion.

------------------------------------------
* win32com method uses Excel application installed on the pc and converts .xlsx file to .xls
  - **Pros**
      - Since it uses Excel application directly to convert, no error message on launch of newly generated .xls file
  - **Cons**
      - It requires Excel to be installed. 
      - Script can run the script with Office03 **but** excel file will break. 
     
------------------------------------------

* openpyxl method uses openpyxl module to convert the .xlsx file to .xls
  - **Pros**
      - Does not require Excel to be installed
  - **Cons**
      - Opening the file prompts the user with an error message and the user must select "Yes" to proceed.
