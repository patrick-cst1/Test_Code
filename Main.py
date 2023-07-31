import xlwings as xw
import pandas as pd
import glob

Raw_path = glob.glob(r"D:\Test_Code\Raw\ITD*.xlsx")
df = pd.read_excel(r"D:\Test_Code\ITD.xlsx")
df2 = pd.read_excel(Raw_path)
df_combined = pd.concat([df, df2], axis=0)


#load workbook
app = xw.App(visible=False)
wb = xw.Book(r"D:\Test_Code\ITD.xlsx")  
ws = wb.sheets['Sheet1']

# last_row = int(ws.range('A' + str(ws.cells.last_cell.row)).end('up').row) + 1
#Update workbook at specified range
# ws.range('A' + str(last_row)).options(index=False).value = df
ws.range('A1').options(index=False).value = df_combined

#Close workbook
wb.save()
wb.close()
app.quit()