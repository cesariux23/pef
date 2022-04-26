FOR %%f IN (*.xlsx) DO (
    soffice.exe --headless --convert-to xls %%f
)