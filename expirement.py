import os
import sys
from os import listdir
from os.path import isfile, join
import openpyxl
import re
import csv

possibleColTitle = ["id", "first name", "name", "student id"]

searchField = ['id', 'student id']
myPath = os.getcwd()
logDir = 'checkInLogs'
logDir = join(myPath, logDir)

def findTitleRow(sheet) -> int:
    for r in range(1, 5):
        for c in range(1, sheet.max_column+1):
            if (sheet.cell(r, c).value != None) and type(sheet.cell(r,c).value) == type('str') and (sheet.cell(r, c).value.lower() in possibleColTitle):
                return r
    raise ValueError("No Proper Title Row Detected, Ensure The Tile Row Contains One Of The Possible Col Title")

def findColByTitles(sheet, titles: list[str]):
    titleRow = findTitleRow(sheet)
    for c in range(1, sheet.max_column+1):
        if (type(sheet.cell(titleRow, c).value) == type('str')) and (sheet.cell(titleRow, c).value.lower() in titles):
            return c
    raise ValueError("Make Sure \"ID\" or \"student id\" Is One Of the Col Title, or Update findColByName Function")


def dateTimeStr2Tuple(dateTime: str):
    regex = r"(\d+)/(\d+)/(\d+) (\d+):(\d+):(\d+)"
    match = re.search(regex, dateTime)
    dateTimeTuple = tuple(match.group(i+1) for i in range(6))
    return dateTimeTuple

def removeLeadZero(str: str):
    str.lstrip('0')
    if str == '':
        str = '0'
    return str

def findDataStartingRow(sheet, anchorDateTime:str) -> int:
    titleRow = findTitleRow(sheet)
    col = findColByTitles(sheet, ['time', 'timestamp'])
    maxRow = sheet.max_row
    anchorDateTime = dateTimeStr2Tuple(anchorDateTime)
    for row in range(titleRow+1, maxRow+1):
        currentDateTime = sheet.cell(row, col).value
        currentDateTime = dateTimeStr2Tuple(currentDateTime)
        if currentDateTime[:3] == anchorDateTime[:3]:
            hourI = 3
            minI = 4
            currentHour = int(removeLeadZero(currentDateTime[hourI]))
            anchorHour = int(removeLeadZero(anchorDateTime[hourI]))
            currentMin = int(removeLeadZero(currentDateTime[minI]))
            anchorMin = int(removeLeadZero(anchorDateTime[minI]))
            if (currentHour == anchorHour and currentMin >= anchorMin) or (currentHour > anchorHour):
                return row
    raise ValueError("No value start after input date time, reconfirm the excel and/or date time")

def findDataEndingRow(sheet, anchorDateTime: str) -> int:
    titleRow = findTitleRow(sheet)
    col = findColByTitles(sheet, ['time', 'timestamp'])
    maxRow = sheet.max_row
    endRow = 0
    anchorDateTime = dateTimeStr2Tuple(anchorDateTime)
    for row in range(titleRow+1, maxRow+1):
        currentDateTime = sheet.cell(row, col).value
        currentDateTime = dateTimeStr2Tuple(currentDateTime)
        if currentDateTime[:3] == anchorDateTime[:3]:
            endRow = row
    if endRow == 0: raise ValueError("no entry in check log with given date")
    return endRow

def translateCSV2XLSX(csvAddress: str) -> str:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    with open(csvAddress) as csvFile:
        csvData = csv.reader(csvFile, delimiter=',')
        for row in csvData:
            sheet.append(row)
    xlsxAddress = csvAddress.replace(".csv", ".xlsx")
    workbook.save(xlsxAddress)
    return xlsxAddress


path = "Checkin Log D-J.xlsx"
book = openpyxl.load_workbook(path)
sheet = book.active
cell = sheet.cell(2, 5).value
a = 3
print(  )