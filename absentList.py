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

def isExactOneArg() -> bool:
    return len(sys.argv) == 2

def getMasterName() -> str:
    if isExactOneArg():
        return sys.argv[1]
    raise ValueError("Ensure Exactly One Master Sheet After the Program!")



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

def ensureXLSX(fileName) -> str:
    excelExt = '.xlsx'
    csvExt = '.csv'
    if excelExt in fileName and not (csvExt in fileName):
        return fileName
    if csvExt in fileName and not (excelExt in fileName):
        return translateCSV2XLSX(fileName)
    raise TypeError("checkInLogs consists files outside xlsx or csv! remove them")



def getLogFiles() -> list[str]:
    masterName = getMasterName()
    logFiles = set()
    for fileName in listdir(logDir):
        if masterName == fileName: raise FileExistsError("master file should NOT exist in check in log directory, remove it")
        fileAddress = join(logDir, fileName)
        if isfile(fileAddress):
            xlsxAddress = ensureXLSX(fileAddress)
            logFiles.add(xlsxAddress)
            
        else:
            raise FileExistsError(f"non-file object in checkInLogs, remove {fileName}")
    return logFiles


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

def main():
    masterFileAddress = getMasterName()
    logFileAddresses = getLogFiles()
    resultIDs = diffIDs(masterFileAddress, logFileAddresses)
    print(resultIDs)

main()