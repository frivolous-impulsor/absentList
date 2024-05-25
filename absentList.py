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
    print(sorted(resultIDs))
    print(len(resultIDs))

main()


""" results = ['000753041', '001134985', '001139572', '001147237', '001152170', '001210943', '001219510', '001300699', '001315037', '001335315', '001415044', '001419807', '001437219', '400009243', '400009502', '400016397', '400022002', '400025264', '400030771', '400066785', '400071530', '400083999', '400107526', '400132382', '400136304', '400139358', '400164152', '400168872', '400172765', '400177619', '400197345', '400212431', '400224393', '400232595', '400288553', '400288749', '400311990', '400316036', '400318202', '400320751', '400320769', '400322709', '400325095', '400330742', '400341942', '400341947', '400352961', '400353316', '400353406', '400353411', '400353418', '400353740', '400357272', '400390217', '400425629', '400425630', '400425641', '400425657', '400425725', '400425729', '400425757', '400425767', '400425770', '400425780', '400425781', '400463214', '400467650', '400474557', '400475682', '400476638', '400480075', '400480132', '400485328']

answerPath = "Absent list_C3_Nursing_ASN.xlsx"
answerBook = openpyxl.load_workbook(answerPath)
answerSheet = answerBook["Absent"]
answerIDs = []
for row in range(1, answerSheet.max_row+1):
    answerId = answerSheet.cell(row, 3).value
    answerIDs.append(answerId)
answerIDs = answerIDs[2:]
print(answerIDs)

for answerId in answerIDs:
    if answerId not in results:
        print(answerId) """