import os
import os.path

import datetime
import pandas as myPandas

import pptx as myPythonPptx
from pptx import Presentation

# Google sheet API & Credential
import gspread

pathToRawExcelFileTemplate = "E:\\Downloads\\Defect Improvement Status - IP34" # Defect Improvement Status - IP34 - 220428.xlsx
pathToRawPptFile = "E:\\Downloads\\Defect Improvement Status - IP34.pptx"
pathToOutputFile = f"DailyIssue.xlsx"
componentToFind = ["Vehicle"]
# "[Vehicle]Daily defect report"

TEAM = "Team"
PIC = "P.I.C"
TICKET_DESC = "Ticket desc."
ISSUE_CATE = "Issue Category"
JIRA_NUM = "JIRA Num."
m_TicketsDict = dict()
m_TicketsDict[TEAM] =            []
m_TicketsDict[PIC] =            []
m_TicketsDict[TICKET_DESC] =        []
m_TicketsDict[ISSUE_CATE] =        []
m_TicketsDict[JIRA_NUM] =           []

# Google Spreadsheet Config
# If modifying these scopes, delete the file token.json.

# SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '[Vehicle]Daily defect report'

def exportToGoogleSheet():
    sa = gspread.service_account()
    sheet = sa.open(SPREADSHEET_ID)

    currentDay = datetime.datetime.today()
    monthName = datetime.datetime.strptime(str(currentDay.month), "%m")
    shortMonthName = monthName.strftime("%b")
    workSheetName = str(currentDay.day) + '-' + shortMonthName;
    # sheet.del_worksheet(workSheetName)
    sheet.add_worksheet(workSheetName, 100, 25)
    
    toDayWorkSheet = sheet.worksheet(workSheetName)
    print("Start push data to Google Sheet")

    dataframe = myPandas.DataFrame(m_TicketsDict)
    toDayWorkSheet.update([dataframe.columns.values.tolist()] + dataframe.values.tolist())
    #wks.update('D2:E3', [['Engineering', 'Tennis'], ['Business', 'Pottery']])
    toDayWorkSheet.update('F1', 'Root Cause')
    toDayWorkSheet.update('F1:K1', [['Root Cause', 'Original Target', 'Changed Target', 'Reason for changing target', 'Blocker', 'CR Related']])

    # Formarting google sheet:
    toDayWorkSheet.format('A1:K1', {'textFormat': {'bold': True}})
    numberOfRows = len(m_TicketsDict[JIRA_NUM]) + 1
    # GREEN:   RGB(0.58 0.77 0.5)
    # ORANGNE: RGB(1.0 0.85 0.4)

    colToBeColored = 'A1:C' + str(numberOfRows)
    toDayWorkSheet.format(colToBeColored, {
        "backgroundColor": {
        "red": 0.58, "green": 0.77, "blue": 0.5
        },
        "horizontalAlignment": "LEFT"
    })

    colToBeColored = 'D1:D' + str(numberOfRows)
    toDayWorkSheet.format(colToBeColored, {
        "backgroundColor": {
        "red": 1.0, "green": 0.85, "blue": 0.4
        },
        "horizontalAlignment": "CENTER"
    })

    colToBeColored = 'E1:E' + str(numberOfRows)
    toDayWorkSheet.format(colToBeColored, {
        "backgroundColor": {
        "red": 0.58, "green": 0.77, "blue": 0.5
        },
        "horizontalAlignment": "CENTER"
    })

    colToBeColored = 'F1:K' + str(numberOfRows)
    toDayWorkSheet.format(colToBeColored, {
        "backgroundColor": {
        "red": 1.0, "green": 0.85, "blue": 0.4
        },
        "horizontalAlignment": "CENTER"
    })

# Make 1 become "01"

def insertZeroToNumber(inp):
    if (inp < 10):
        return "0" + str(inp)
    else:
        return str(inp)

def getTicketFromRawDataByDate(list_component_to_find, currentDay):
    strDay = insertZeroToNumber(currentDay.day)
    strMonth = insertZeroToNumber(currentDay.month)
    strYear = str(currentDay.year % 100) # from 2022 to 22

    toDayExcelFileName = pathToRawExcelFileTemplate + " - " + strYear + strMonth + strDay + ".xlsx"
    toDaySheetName = "RawData_" + strMonth + strDay

    rawXls = myPandas.ExcelFile(toDayExcelFileName)
    
    sheetToday = myPandas.read_excel(rawXls, toDaySheetName)
    
    lastLine = sheetToday.index[-1]
    print("Max index: " + str(lastLine))
    count = 0
    for i in range(0, lastLine):
        currentLineDataFrame = sheetToday.iloc[i]
        for component in list_component_to_find:
            if component in currentLineDataFrame['Component/s']:
                m_TicketsDict[TEAM].append('Must filled')
                m_TicketsDict[PIC].append(currentLineDataFrame['Assignee'])
                m_TicketsDict[TICKET_DESC].append(currentLineDataFrame['Summary'])
                m_TicketsDict[ISSUE_CATE].append('Must filled')
                m_TicketsDict[JIRA_NUM].append(currentLineDataFrame['Issue key'])
                print(currentLineDataFrame)
                count += 1
    if (count < 2):
        print("Today has: " + str(count) + " ticket")
    else:
        print("Today has: " + str(count) + " tickets")
    return

def exportToExcel():
    dataframe = myPandas.DataFrame(m_TicketsDict)
    currentDay = datetime.datetime.today()
    todaySheetName = "RawData_" + insertZeroToNumber(currentDay.day + 1) + insertZeroToNumber(currentDay.month)
    try:
        with myPandas.ExcelWriter(pathToOutputFile) as writer:
            dataframe.to_excel(writer, sheet_name=todaySheetName)
    except:
      print("\nError: ---> Please close output file " + pathToOutputFile)
    return

def createPptFromTemplate():
    myPptx = Presentation(pathToRawPptFile)
    print(myPptx)
    return myPptx

def printComment():
    nTickets = len(m_TicketsDict[JIRA_NUM])
    COMMENT_CALL_TO_FILL = f'Hi all, hôm nay có {nTickets} tickets, nhờ mọi người điền nhé: https://docs.google.com/spreadsheets/d/1fVlAgdk6OARRty9LOnZVb2BSMCAjy1j9x_BHg3sbhVs/edit?usp=sharing'
    print(COMMENT_CALL_TO_FILL)

    currentDay = datetime.datetime.today()
    monthName = datetime.datetime.strptime(str(currentDay.month), "%m")
    shortMonthName = monthName.strftime("%b")
    fDay = str(currentDay.day) + ' ' + shortMonthName + '.'
    COMMENT_SEND_EMAIL = f'Dear chị Vy, em gửi daily defect report ({fDay}) của Vehicle ạ.'
    print(COMMENT_SEND_EMAIL)

def main():
    getTicketFromRawDataByDate(componentToFind, datetime.datetime.today())
    # createPptFromTemplate()
    exportToExcel()
    exportToGoogleSheet()
    # testClient()
    printComment()
    print("Completed.")

if __name__ == "__main__":
    main()