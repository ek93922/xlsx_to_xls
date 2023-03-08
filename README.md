# xlsx_to_xls

Converts xlsx files to xls in the same directory as the Python Script.
  - Will not create duplicate xls file if there's already one in the folder from previous conversion.

------------------------------------------
* win32com method uses Excel application installed on the pc and converts .xlsx file to .xls
  - **Pros**
      - Since it uses Excel application directly to convert, no error message on launch of newly generated .xls file
  - **Cons**
      - It requires Excel to be installed. (Haven't tried forward capabilities. ie. using Office03 to convert the file from .xlsx to .xls)
      - (Theorectically can't happen since Office03 can't open .xlsx file)
------------------------------------------

* openpyxl method uses openpyxl module to convert the .xlsx file to .xls
  - **Pros**
      - Does not require Excel to be installed
  - **Cons**
      - Opening the file prompts with error message and the user must select "Yes" to proceed.
