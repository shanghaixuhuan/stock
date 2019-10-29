from openpyxl import Workbook
import time

path = "data.xlsx"
wb = Workbook()
ws = wb.active
date = time.strftime("%Y-%m-%d", time.localtime())
ws['A1'] = "=EM_SECTOR(\"2000032254\", \"" + date + "\")"
wb.save(filename=path)

