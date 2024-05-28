import os
import sys
from os import listdir
from os.path import isfile, join
import openpyxl
import re
import csv

address = "C3_Nursing_ASN.xlsx"
book = openpyxl.load_workbook(address)
sheet = book['sheet1']
sheet.delete_rows(1,100)  