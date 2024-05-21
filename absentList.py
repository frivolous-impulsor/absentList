import os
from os import listdir
from os.path import isfile, join
import openpyxl
import re
from openpyxl import Workbook

possibleColTitle = ["id", "first name", "name", "student id"]

searchField = 'ID'
myPath = os.getcwd()
thisFile = __file__
minorfiles = [f for f in listdir(myPath) if (isfile(join(myPath, f)) and bool(re.search('.xlsx', f)) and f != "master.xlsx")]

def findTitleRow(sheet):
    for r in range(1, 5):
        for c in range(1, sheet.max_column+1):
            if (sheet.cell(r, c).value != None) and (sheet.cell(r, c).value.lower() in possibleColTitle):
                return r
    raise ValueError("No Proper Title Row Detected, Ensure The Tile Row Contains One Of The Possible Col Title")

def findColByName(sheet, colName: str):
    IDcol = 0
    colName = colName.lower()
    titleRow = findTitleRow(sheet)
    for c in range(1, sheet.max_column+1):
        if sheet.cell(titleRow, c).value.lower() == colName:
            return c
    raise ValueError("Make Sure \"ID\" or \"student id\" Is One Of the Col Title, or Update findColByName Function")



def setIDsDict(path: str, idDict: dict, status: bool):
    book = openpyxl.load_workbook(path)
    sheet = book.active
    maxRow = sheet.max_row
    titleRow = findTitleRow(sheet)
    Col = findColByName(sheet, searchField)

    for r in range(titleRow+1, maxRow+1):
        id = sheet.cell(r, Col).value
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

resultIDs = diffIDs("master.xlsx", minorfiles)
print(resultIDs)