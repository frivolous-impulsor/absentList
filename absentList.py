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

beginTime = "5/22/2024 13:00:00"

def getMasterName() -> str:
    if len(sys.argv) > 1 and type(sys.argv[1]) == type("str"):
        masterName = sys.argv[1]
        if len(masterName) > 2 and masterName[0] == '.':
            masterName = masterName[2:]
        if masterName in listdir(myPath) and isfile(join(myPath, masterName)):
            return masterName
        else:
            raise ValueError("the master sheet should be in current(same as the py script) directory")
    else:
        raise ValueError("ensure the first argument is a valid name of a xlsx file being the master sheet")



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
    raise ValueError("no colume with listed title found on title row")

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
    raise ValueError(f"No value start after {anchorDateTime}, reconfirm the excel and/or date time")

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
        
        



def setIDsDict(path: str, idDict: dict, isLog: bool):
    book = openpyxl.load_workbook(path)
    sheet = book.active
    idCol = findColByTitles(sheet, searchField)
    if isLog:
        try: startRow = findDataStartingRow(sheet, beginTime)
        except ValueError: return
        endRow = findDataEndingRow(sheet, beginTime)
    else:
        startRow = findTitleRow(sheet)
        endRow = sheet.max_row
    for r in range(startRow, endRow+1):
        id = sheet.cell(r, idCol).value
        if id != None and id.isnumeric():
            idDict[id] = isLog



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
    return resultIDs

def test():
    result = main()

    book = openpyxl.load_workbook("Absence List - C2_HealthSciences_ASN.xlsx")
    sheet = book["Absent"]

    answer = []
    for row in range(1, sheet.max_row+1):
        id = sheet.cell(row, 3).value
        answer.append(id)

    resultCOMPanswer = []
    answerCOMPresult = []

    for id in answer:
        if id not in result:
            resultCOMPanswer.append(id)

    for id in result:
        if id not in answer:
            answerCOMPresult.append(id)

    print(resultCOMPanswer)
    print(answerCOMPresult)

test()
        