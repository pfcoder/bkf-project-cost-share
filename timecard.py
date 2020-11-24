from openpyxl import load_workbook
from dateutil import relativedelta
from datetime import datetime
import os
import sys

basePath = "timecard_input/"
outputPath = "timecard_output/"

def init():

    # load all source excel from specified dirs
    sources = os.listdir(basePath)
    print(sources)
    # process all excels under this folder
    for i in range(0, len(sources)):
        fileName = sources[i]
        if not fileName.endswith('.xlsx') or fileName.startswith('~'):
            continue

        print("start process:{}".format(fileName))

        wb = load_workbook(basePath + fileName)
        processSource(wb, fileName)


def loadPayInfo(sheet):
    # print("load payInfo of {}".format(sheet))
    result = {}
    for i in range(3, sheet.max_row + 1):
        name = sheet.cell(row=i, column=1).value
        result[name] = []
        for j in range(2, sheet.max_column + 1):
            result[name].append(sheet.cell(row=i, column=j).value)

    return result


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


def updateResultDict(prjOrContractResult, employeeResult, prjType, name, code, hours):
    if code not in prjOrContractResult:
        prjOrContractResult[code] = {
            "hours": hours,
            "cost": [0.0] * 7
        }
    else:
        prjOrContractResult[code]["hours"] += hours

    # update employee store
    prjTypeKey = "prjDict"
    if prjType == 1:
        prjTypeKey = "contractDict"

    if code not in employeeResult[name][prjTypeKey]:
        employeeResult[name][prjTypeKey][code] = hours
    else:
        employeeResult[name][prjTypeKey][code] += hours


def processSource(wb, fileName):
    # result dict
    projectResultDict = {}
    contractResultDict = {}

    # some count store, dict key is employee name, store total working hours
    employeeDict = {}

    timeRecordsSheet = wb[wb.sheetnames[0]]
    contractPayInfosSheet = wb[wb.sheetnames[1]]
    projectPayInfosSheet = wb[wb.sheetnames[2]]

    # load contract pay info
    contractPayInfo = loadPayInfo(contractPayInfosSheet)
    # print(contractPayInfo)

    # load project pay info
    projectPayInfo = loadPayInfo(projectPayInfosSheet)
    # print(projectPayInfo)

    # go through timeRecordsSheet row by row
    for i in range(2, timeRecordsSheet.max_row + 1):
        name = timeRecordsSheet.cell(row=i, column=1).value;
        print(".", end='')

        # check project or contract
        prjCell = timeRecordsSheet.cell(row=i, column=25)
        contractCell = timeRecordsSheet.cell(row=i, column=18)
        if (not isEmptyCell(prjCell)) and (not isEmptyCell(contractCell)):
            print("发现项目和合同同时存在：{} {} {} 请改正".format(name, prjCell.value, contractCell.value))
            sys.exit(0)

        hours = float(timeRecordsSheet.cell(row=i, column=3).value)

        if name not in employeeDict:
            employeeDict[name] = {
                "totalHours": hours,
                "prjDict": {},
                "contractDict": {}
            }
        else:
            employeeDict[name]["totalHours"] += hours

        if not isEmptyCell(prjCell):
            code = prjCell.value
            updateResultDict(projectResultDict, employeeDict, 0, name, code, hours)

        if not isEmptyCell(contractCell):
            code = contractCell.value
            updateResultDict(contractResultDict, employeeDict, 1, name, code, hours)

    print(employeeDict)
    print(projectResultDict)
    print(contractResultDict)

    # store error record
    missPrjName = set()
    missContractName = set()

    # now start calculate
    # go through employee dict
    for name in employeeDict:
        prjDict = employeeDict[name]['prjDict']
        contractDict = employeeDict[name]['contractDict']
        employeeMonthHours = employeeDict[name]['totalHours']
        for code in prjDict:
            prjHours = prjDict[code]
            rate = prjHours / employeeMonthHours

            if name not in projectPayInfo:
                missPrjName.add(name)
                continue

            employeePrjPayInfo = projectPayInfo[name]
            compute = [p * rate for p in employeePrjPayInfo]

            projectResultDict[code]['cost'] = [round(i + j, 2) for i, j in zip(projectResultDict[code]['cost'], compute)]

        for code in contractDict:
            prjHours = contractDict[code]
            rate = prjHours / employeeMonthHours

            if name not in contractPayInfo:
                missContractName.add(name)
                continue

            employeeContractPayInfo = contractPayInfo[name]
            compute = [p * rate for p in employeeContractPayInfo]

            contractResultDict[code]['cost'] = [round(i + j, 2) for i, j in zip(contractResultDict[code]['cost'], compute)]

    print(projectResultDict)
    print(contractResultDict)

    if len(missPrjName) > 0:
        print("{} 缺少以下研发人员薪资信息：{}".format(fileName, missPrjName))

    if len(missContractName) > 0:
        print("{} 缺少以下实施人员薪资信息：{}".format(fileName, missContractName))

    #if len(missPrjName) > 0 or len(missContractName) > 0:
    #    sys.exit(0)

    # write out result excel
    resultPrjSheetName = "项目汇总"
    resultContractSheetName = "实施汇总"

    if resultPrjSheetName in wb.sheetnames:
        del wb[resultPrjSheetName]

    if resultContractSheetName in wb.sheetnames:
        del wb[resultContractSheetName]

    updateTarget(wb, resultPrjSheetName, projectResultDict)
    updateTarget(wb, resultContractSheetName, contractResultDict)

    wb.save(outputPath + fileName)


def updateTarget(wb, title, dict):
    targetSheet = wb.create_sheet(title=title)
    targetSheet.append(["代号", "总工时", "个人工资", "单位养老", "单位失业", "单位工伤", "单位生育", "单位医疗", "单位公积金"])

    for code in dict:
        row = [code, dict[code]['hours']]
        row += dict[code]['cost']
        targetSheet.append(row)


init()
