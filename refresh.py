import xlwings as xw
import time

app = xw.App(visible=True,add_book=False)
wb = app.books.add()
ws = wb.sheets[0]

while ws.range('a100').value == None:
    time.sleep(1)