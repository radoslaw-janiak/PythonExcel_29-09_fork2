import xlwings as xw, psutil
# działa tylko na windows
# Nowy plik Excela
wb = xw.Book()

# Import modułu VBA
wb.api.VBProject.VBComponents.Import(r"C:\Users\radek\PycharmProjects\PythonExcelXlwings\makro.bas")

# Zapisz jako plik z makrami (.xlsm)
wb.save(r"C:\Users\radek\PycharmProjects\PythonExcelXlwings\plik_z_makrem.xlsm")

# Uruchom makro
makro = wb.macro("HelloWorld")
makro()

wb.close()