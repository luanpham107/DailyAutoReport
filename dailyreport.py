import os
import os.path

import datetime
import pandas as myPandas

import pptx as myPythonPptx
from pptx import Presentation

# Google sheet API & Credential
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient import discovery
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

pathToRawExcelFileTemplate = "E:\\Downloads\\Defect Improvement Status - IP34" # Defect Improvement Status - IP34 - 220428.xlsx
pathToRawPptFile = "E:\\Downloads\\Defect Improvement Status - IP34.pptx"
pathToOutputFile = f"DailyIssue.xlsx"
componentToFind = ["Vehicle"]
# "[Vehicle]Daily defect report"


PIC = "P.I.C"
TICKET_DESC = "Ticket desc."
ISSUE_CATE = "Issue Category"
JIRA_NUM = "JIRA Num."
m_TicketsDict = dict()
m_TicketsDict[PIC] =            []
m_TicketsDict[TICKET_DESC] =        []
m_TicketsDict[ISSUE_CATE] =        []
m_TicketsDict[JIRA_NUM] =           []

# Google Spreadsheet Config
# If modifying these scopes, delete the file token.json.

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1fVlAgdk6OARRty9LOnZVb2BSMCAjy1j9x_BHg3sbhVs'
SAMPLE_RANGE_NAME = '6-May!A2:E'

def exportToGoogleSheet():
    print("Start push data to Google Sheet")
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = discovery.build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[4]))
        values = [
            [
                # Cell values ...
            ],
            # Additional rows ...
        ]

        body = {
            'values': values
        }
        result = sheet.values().update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME,
            valueInputOption='RAW', body=body).execute()

        print('{0} cells updated.'.format(result.get('updatedCells')))
    except HttpError as err:
        print(err)
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

def main():
    getTicketFromRawDataByDate(componentToFind, datetime.datetime.today())
    # createPptFromTemplate()
    exportToExcel()
    # exportToGoogleSheet()
    # testClient()
    print("Completed.")

if __name__ == "__main__":
    main()