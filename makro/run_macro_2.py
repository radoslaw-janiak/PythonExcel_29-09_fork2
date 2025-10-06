import openpyxl

book = openpyxl.load_workbook("xl/macro.xlsm", keep_vba=True)
book["Arkusz1"]["A1"].value = "Kliknij przycisk!"
book.save("macro_openpyxl.xlsm")