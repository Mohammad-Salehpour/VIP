# import pandas as pd
# import xlwings as xw
#
# # Opening an excel file
# wb = xw.Book(r'C:\Users\Mohammad Salehpour\PycharmProjects\pythonProject\VIP\active_distinct.xlsx')
#
# # Viewing available
# # sheets in it
# wks = xw.sheets
# print("Available sheets :\n", wks)
#
# # Selecting a sheet
# ws = wks[0]
#
# # Selecting a value
# # from the selected sheet
# val = ws.range("A1").value
# print("A value in sheet1 :", val)
# print(wb.sheets)
# ws.range("A1").value = df
# last_row = ws.range(1,1).end('down').row
# print("The last row is {row}.".format(row=last_row))
# print("The DataFrame df has {rows} rows.".format(rows=df.shape[0]))

# import pandas as pd
# import numpy as np
#
# path = r'C:\Users\Mohammad Salehpour\PycharmProjects\pythonProject\VIP\active_distinct.xlsx'
# # xl = pd.ExcelFile(path)
# # print(path)
# # df = xl.parse("dummydata")
# # df = pd.read_excel(r'C:\Users\Mohammad Salehpour\PycharmProjects\pythonProject\VIP\active_distinct.xlsx')
# # print(len(df))
#
# i = 0
# for df in pd.ExcelFile.parse(r'C:\Users\Mohammad Salehpour\PycharmProjects\pythonProject\VIP\active_distinct.xlsx', chunksize=1999):
#     df.to_excel('./partition/file_{:02d}.xlsx'.format(i), index=False)
#     i += 1
# # #
# # import pandas as pd
# # import numpy as np
#
#
#
# chunkSize = 1999
# i = 0
# for chunk in np.array_split(df, 1999):
#     chunk.to_excel('./partition/file_{:02d}.xlsx'.format(i), index=False)
#     i += 1



# for chunk in np.array_split(df, len(df) // chunkSize):
#     chunk.to_excel('./partition/file_{:02d}.xlsx'.format(i), index=True)
#
import pandas as pd

df = pd.read_excel(r'C:\Users\Mohammad Salehpour\PycharmProjects\pythonProject\VIP\active_distinct.xlsx')

rows_per_file = 1999

n_chunks = len(df) // rows_per_file

for i in range(n_chunks):
    start = i*rows_per_file
    stop = (i+1) * rows_per_file
    sub_df = df.iloc[start:stop]
    sub_df.to_excel(f"./partition/Chunked_part{i+1}.xlsx", index=True)
if stop < len(df):
    sub_df = df.iloc[stop:]
    sub_df.to_excel(f"./partition/Chunked_part{i+1}.xlsx", index=True)