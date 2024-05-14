from os import listdir
from os.path import isfile, join
import openpyxl
import random
from openpyxl import Workbook


#deleting some of the rows from master list to create sudo attended list
pathMarch = 'attended.xlsx'
book = openpyxl.load_workbook(pathMarch)
sheet = book.active

randomRemoveRow = random.sample(range(3, sheet.max_row), 20)
print(randomRemoveRow)

for r in randomRemoveRow:
    sheet.delete_rows(r)
book.save('attended.xlsx')
#set up complete