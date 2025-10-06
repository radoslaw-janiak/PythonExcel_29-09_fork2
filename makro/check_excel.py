import xlwings as xw

app = xw.App(visible=True, add_book=False)
wb = app.books.add()
wb.sheets[0]["A1"].value = "Hello from xlwings on macOS"
wb.save("xlwings_test.xlsx")
app.quit()