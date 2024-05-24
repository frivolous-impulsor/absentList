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
minorDir = join(myPath, logDir)

def isExactOneArg() -> bool:
    return len(sys.argv) == 2

def getMasterFile() -> str:
    if isExactOneArg():
        return sys.argv[1]
    raise ValueError("Ensure Exactly One Master Sheet After the Program!")

def getMinorFiles() -> list[str]:
    return [join(minorDir, f) for f in listdir(minorDir) if (isfile(join(minorDir, f)) and bool(re.search('.xlsx', f)) and f != getMasterFile())]

def translateCSV2XLSX(csvAddress: str):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    
    with open(csvAddress) as csvFile:
        csvData = csv.reader(csvFile, delimiter=',')
    for row in csvData:
        sheet.append(row)
    xlsxAddress = csvAddress.replace(".csv", ".xlsx")
    workbook.save(xlsxAddress)

def ensureXLSX(fileName):
    excelExt = '.xlsx'
    csvExt = '.csv'
    if excelExt in fileName and not (csvExt in fileName):
        return 0
    if csvExt in fileName and not (excelExt in fileName):
        translateCSV2XLSX(fileName)
        return 1
    raise TypeError("checkInLogs consists files outside xlsx or csv! remove them")
    
    
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

def findDataStartingRow(sheet):
    "todo"


def setIDsDict(path: str, idDict: dict, status: bool):
    book = openpyxl.load_workbook(path)
    sheet = book.active
    maxRow = sheet.max_row
    titleRow = findTitleRow(sheet)
    titleCol = findColByTitles(sheet, searchField)

    for r in range(titleRow+1, maxRow+1):
        id = sheet.cell(r, titleCol).value
        idDict[id] = status


def diffIDs(masterPath: str, minorPaths: list[str]):
    IDdict = {}
    setIDsDict(masterPath, IDdict, False)
    for minor in minorPaths:
        setIDsDict(minor, IDdict, True)

    absentIDs = []
    for key in IDdict.keys():
        if not IDdict[key]:
            absentIDs.append(key)
    return absentIDs

resultIDs = diffIDs("master.xlsx", getMinorFiles())
print(resultIDs)