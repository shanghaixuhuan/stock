import xlwings as xw
import time
import pymongo
from openpyxl import load_workbook

path = "data.xlsx"
date = time.strftime("%Y-%m-%d", time.localtime())

app = xw.App(visible=True,add_book=False)
wb = app.books.add()
ws = wb.sheets[0]

ws.range('a1').value = "=EM_SECTOR(\"2000032254\", \"" + date + "\")"
while ws.range('a100').value == None:
    time.sleep(1)

print("股票完成导入")

rows = ws["A1048576"].end('up').row

# for i in range(2, 2 + data_num):
#     if((i-1) % 500 == 0):
#         print("第" + str(i-1) + "条数据完成导入")
#     share_num = ws["B" + str(i)].value
#     ws['C' + str(i)].value = "=EM_S_SEST_NETPROFITFY1(\"" + share_num + "\",\"" + date + "\")"
#     ws['D' + str(i)].value = "=EM_S_SEST_NETPROFITFY2(\"" + share_num + "\",\"" + date + "\")"
#     ws['E' + str(i)].value = "=EM_S_SEST_NETPROFITFY3(\"" + share_num + "\",\"" + date + "\")"
#     ws['F' + str(i)].value = "=EM_S_SEST_NETPROFITF12(\"" + share_num + "\",\"" + date + "\")"
#     ws['G' + str(i)].value = "=EM_S_SEST_NETPROFITYOY(\"" + share_num + "\",\"" + date + "\")"

ws['C2:C' + str(rows)].value = "=EM_S_SEST_NETPROFITFY1(" + "B2" + ",\"" + date + "\")"
ws['D2:D' + str(rows)].value = "=EM_S_SEST_NETPROFITFY2(" + "B2" + ",\"" + date + "\")"
ws['E2:E' + str(rows)].value = "=EM_S_SEST_NETPROFITFY3(" + "B2" + ",\"" + date + "\")"
ws['F2:F' + str(rows)].value = "=EM_S_SEST_NETPROFITF12(" + "B2" + ",\"" + date + "\")"
ws['G2:G' + str(rows)].value = "=EM_S_SEST_NETPROFITYOY(" + "B2" + ",\"" + date + "\")"

while ws['C' + str(rows)].value == "Refreshing" or ws['D' + str(rows)].value == "Refreshing" or ws['E' + str(rows)].value == "Refreshing" or \
        ws['F' + str(rows)].value == "Refreshing" or ws['G' + str(rows)].value == "Refreshing":
    time.sleep(1)

wb.save(path)
wb.close()
app.quit()

print("相关数据完成导入")

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["shares"]
mycol = mydb["shares"]
wb = load_workbook(filename=path,data_only=True)
ws = wb.active
data_amount = ws.max_row - 1
for i in range(2, 2 + data_amount):
    mycol.insert({"stock_name" : ws['A' + str(i)].value, "date" : date, "stock_code" : ws['B' + str(i)].value,
                  "FY1" : ws['C' + str(i)].value, "FY2": ws['D' + str(i)].value, "FY3" : ws['E' + str(i)].value,
                  "F12" : ws['F' + str(i)].value, "YOY" : ws['G' + str(i)].value,})

print("导入MongoDB完成")