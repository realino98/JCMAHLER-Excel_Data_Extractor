import pandas as pd
import numpy as np

path = r"C:\Users\fedel\Desktop\excelData\PhD_data.xlsx"

x1 = np.random.randn(100, 2)
df1 = pd.DataFrame(x1)

x2 = np.random.randn(100, 2)
df2 = pd.DataFrame(x2)

writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
df1.to_excel(writer, sheet_name = 'x1')
df2.to_excel(writer, sheet_name = 'x2')
writer.save()
writer.close()