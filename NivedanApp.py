import os
import socket
import PySimpleGUI as sg
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import date, datetime
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1tVRojVxTdw20U0ZV2jooMkI4Bjd1CozLPaEhEGCOZA4"

def authenticate_google_sheets():
    credentials = None
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    return credentials

def write_data_to_sheet(credentials, values):
    if not check_internet():
        sg.popup_error('No internet Connection❗Please connect Wi-Fi')
    service = build("sheets", "v4", credentials=credentials)
    sheets = service.spreadsheets()
    range_name = "Sheet1!A:G"  # Adjust the range as needed
    current_date = datetime.now().strftime("%d-%m-%Y")
    row_data = [current_date, values['Pushp No'], values['Yuvakendra'], values['Name'], values['Type'], values['Vishay'], values['Count']]
    body = {"values": [row_data]}
    sheets.values().append(spreadsheetId=SPREADSHEET_ID, range=range_name,
                           body=body, valueInputOption="RAW").execute()

def check_internet():
    try:
        socket.create_connection(("www.google.com",80))
        return True
    except OSError:
        return False
def main():
    sg.theme('DarkTeal9')

    layout = [
        [sg.Text('Pushp No:',size=(15,1)),sg.InputText(key='Pushp No')],
        [sg.Text('Yuva-kendra: ',size=(15,1)),sg.Combo(['LS-Boys Hostel','Blossom','Vaniyavaad'],default_value='yuvakendra',key='Yuvakendra')],
        [sg.Text('Name of Yuvan: ', size=(15, 1)), sg.InputText(key='Name')],
        [sg.Text('VishayType: ', size=(15, 1)), sg.Combo(['Guna-Darshan','Charitra-Darshan','Mantavya-Darshan','Debate','Utsav-Darshan','Anadhyayan'],default_value='vishay',key='Type')],
        [sg.Text('Sankhya: ', size=(15, 1)), sg.Spin([i for i in range(0,101)],initial_value=0,key='Count')],
        [sg.Text('Vishay:', size=(15, 1)), sg.InputText(key="Vishay")],
        [sg.Submit(),sg.Button('Clear'), sg.Button('Exit')]
    ]

    window = sg.Window("Nivedanam", layout)

    credentials = authenticate_google_sheets()
    def clear_input():
        for key in values:
            window[key]('')
        return None
      
    while True:
        event, values = window.read()
        if event =='Clear':
            clear_input()
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Submit':
            write_data_to_sheet(credentials, values)
            sg.popup('Data Saved!')
            window.find_element('Name').update('')
            window.find_element('Vishay').update('')
            window.find_element('Count').update('')
            window.find_element('Type').update('')

    window.close()

if __name__ == "__main__":
    main()
