# xlwings  - praca z plikami excel, wymaga excela !!!

# import xlwings as xw
# import pandas as pd
# import numpy as np
#
# df = pd.DataFrame(data=np.random.randn(100, 5),
#                   columns=[f'Próba {i}' for i in range(1, 6)])
#
# print(df)
# print(df.head()) # 5 pierwszych
# print(df.tail()) # 5 ostatnich
#
# xw.view(df)

import xlwings as xw
import pandas as pd
import numpy as np

df = pd.DataFrame(data=np.random.randn(100, 5),
                  columns=[f"Próba {i}" for i in range(1, 6)])

print(df)
print(df.head())  # 5 pierwszych
print(df.tail())  # 5 ostatnich
#     Próba 1   Próba 2   Próba 3   Próba 4   Próba 5
# 0  0.847841  1.625375 -0.107138  0.327772  2.169133
# 1  0.608697  1.099411 -0.169195 -1.909709  0.890208
# 2 -1.466427  0.859255  0.928376  0.178363 -1.210645
# 3  0.985196  0.596124 -0.618687  0.364990 -0.799930
# 4 -0.980503  0.108145  0.789764 -0.294158  0.177767

# xw.view(df)

book = xw.Book()
print(book.name)
print(book.sheets)
#  Zeszyt1
# Sheets([<Sheet [Zeszyt1]Arkusz1>])

sheet1 = book.sheets[0]
print(sheet1.range('A1'))
# <Range [Zeszyt1]Arkusz1!$A$1>

sheet1.range('A1').value = [[1, 2],
                            [3, 4]]

sheet1.range('A4').value = "Witaj!"

print(sheet1['A1'])  # <Range [Zeszyt1]Arkusz1!$A$1>
print(sheet1['A1'].value)  # 1.0

print(sheet1["A1:B2"])  # <Range [Zeszyt1]Arkusz1!$A$1:$B$2>
print(sheet1["A1:B2"].value)  # [[1.0, 2.0], [3.0, 4.0]]

import xlwings as xw

app = xw.App(visible=True, add_book=False)
wb = app.books.add()
wb.sheets[0]["A1"].value = "Hello from xlwings on macOS"
wb.save("xlwings_test.xlsx")
app.quit()