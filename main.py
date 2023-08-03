input_service_cred_json = 'dependable-star-322813-d4f8ef55004d.json'
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
import csv
import pandas as pd
import datetime
from datetime import date
from datetime import datetime
import pyautogui
import os
import time
scope = ['https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name(input_service_cred_json, scope)
# https://developers.google.com/drive/api/v3/quickstart/python
service = build('drive', 'v3', credentials=credentials)
service_sheet = build('sheets', 'v4', credentials = credentials)

def start_zoom():
    os.startfile(r"C:\Users\ashug\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Zoom\Zoom.lnk")
    joinmeetbtn=pyautogui.locateCenterOnScreen(r"Photos\join meet button.png")
    while (joinmeetbtn == None):
            joinmeetbtn=pyautogui.locateCenterOnScreen(r"Photos\join meet button.png")
    pyautogui.moveTo(joinmeetbtn)
    pyautogui.click()
    return joinmeetbtn


def meet_id(meetid):
    meetingbtn=pyautogui.locateCenterOnScreen(r"Photos\Meeting ID.png")
    while meetingbtn == None:
        meetingbtn=pyautogui.locateCenterOnScreen(r"Photos\Meeting ID.png")
    pyautogui.moveTo(meetingbtn)
    pyautogui.click()
    pyautogui.write(meetid)
            
    joinbtn = pyautogui.locateOnScreen(r"Photos\join button.png")
    while (joinbtn == None):
            joinbtn=pyautogui.locateCenterOnScreen(r"Photos\join button.png")
    pyautogui.moveTo(joinbtn)
    pyautogui.click()
    return joinbtn
    

def pass_tab(passcode):
    passcodebtn = pyautogui.locateOnScreen(r"Photos\meetingpasscode.png")
    while (passcodebtn == None):
        passcodebtn = pyautogui.locateOnScreen(r"Photos\meetingpasscode.png")
    pyautogui.moveTo(passcodebtn)
    pyautogui.click()
    pyautogui.write(passcode)

    join_meetbtn = pyautogui.locateOnScreen(r"Photos\meet_btn.png")
    while (join_meetbtn == None):
        join_meetbtn = pyautogui.locateOnScreen(r"Photos\meet_btn.png")
    pyautogui.moveTo(join_meetbtn)
    pyautogui.click()
    return join_meetbtn


def navigate_meet():
    #record_prompt = pyautogui.locateOnScreen(r"Photos\record screen button.png")
    #while (record_prompt == None):
        #record_prompt = pyautogui.locateOnScreen(r"Photos\record screen button.png")
    #pyautogui.moveTo(record_prompt)
    #pyautogui.click()
    time.sleep(5)
    pyautogui.keyDown('alt')
    pyautogui.press('f')
    pyautogui.keyUp('alt')
    time.sleep(2)
    pyautogui.keyDown('alt')
    pyautogui.press('a')
    pyautogui.keyUp('alt')
    time.sleep(1)
    pyautogui.keyDown('alt')
    pyautogui.press('v')
    pyautogui.keyUp('alt')
    #pyautogui.press('alt')
    time.sleep(2)
    return None

def join_breakout():
    breakout = pyautogui.locateOnScreen(r"Photos\joining breakout.png")
    while (breakout == None):
        breakout = pyautogui.locateOnScreen(r"Photos\joining breakout.png")
    time.sleep(1)
    pyautogui.moveTo(breakout)
    return join_breakout

def join_breakout_first():
    afterbreakout = pyautogui.locateOnScreen(r"Photos\afterbreakoutpopup.png")
    while (afterbreakout == None):
        afterbreakout = pyautogui.locateOnScreen(r"Photos\afterbreakoutpopup.png")
    pyautogui.moveTo(afterbreakout)
    pyautogui.click()
    time.sleep(2)
    return afterbreakout

def record():
    pyautogui.keyDown('alt')
    pyautogui.press('f')
    pyautogui.keyUp('alt')
    recordbtn = pyautogui.locateOnScreen(r"C:\Users\Victor\Pictures\Screenshots\recordbtn.png")
    while (recordbtn == None):
        recordbtn = pyautogui.locateOnScreen(r"C:\Users\Victor\Pictures\Screenshots\recordbtn.png")
    pyautogui.moveTo(recordbtn)
    pyautogui.click()
    time.sleep(2)
    con = pyautogui.locateOnScreen(r"Photos\continuewoaudio.png")
    while (con == None):
            con = pyautogui.locateOnScreen(r"Photos\continuewoaudio.png")
    pyautogui.moveTo(con)
    pyautogui.click()
    
def start_caption():
    caption = pyautogui.locateOnScreen(r"Photos\captions.png")
    while (caption == None):
        caption = pyautogui.locateOnScreen(r"Photos\captions.png")
    pyautogui.moveTo(caption)
    pyautogui.click()
    time.sleep(3)
    pyautogui.press('tab', presses=2)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press("tab")
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('down', presses = 2)
    pyautogui.press('enter')
    time.sleep(2)
    return caption
    

def save_transcript():
    transcript = pyautogui.locateOnScreen(r"Photos\save transcript.png")
    while (transcript == None):
        transcript = pyautogui.locateOnScreen(r"Photos\save transcript.png")
    pyautogui.moveTo(transcript)
    pyautogui.click()
    return transcript

def out_breakout():
    out = pyautogui.locateOnScreen(r"Photos/exit breakout.png")
    while (out == None):
        out = pyautogui.locateOnScreen(r"Photos/exit breakout.png")
    pyautogui.moveTo(out)
def save_recording():
    save_record = pyautogui.locateOnScreen(r"Photos\save recording.png")
    while (save_record == None):
        save_record = pyautogui.locateOnScreen(r"Photos\save recording.png")
    pyautogui.moveTo(save_record)
    pyautogui.click()
    return save_record

def record_multiple():
    while True:
        join_breakout()
        time.sleep(10)
        record()
        time.sleep(4)
        start_caption()
        time.sleep(4)
        #save_transcript()
        time.sleep(3)
        
        save_recording()
    return None

def leave_meeting():
    pyautogui.press('alt')
    pyautogui.keyDown('alt')
    pyautogui.press('q')
    pyautogui.keyUp('alt')
    pyautogui.press('enter')
    return None

def join_meet(meetid,passcode):
    start_zoom()
    time.sleep(10)
    meet_id(meetid)
    time.sleep(10)
    pass_tab(passcode)
    time.sleep(15)
    navigate_meet()
    time.sleep(5)
    join_breakout()
    time.sleep(20)
    #join_breakout_first()
    record()
    start_caption()
    save_transcript()
    save_recording()
    record_multiple()

    
def get_spreadsheet():
    page_token = None
    response = service.files().list(q="mimeType = 'application/vnd.google-apps.spreadsheet' and name = 'Schedule The/Nudge' and trashed = false",supportsAllDrives = True,
                                            fields='nextPageToken, '
                                                   'files(id, name)',pageToken=page_token, includeItemsFromAllDrives = True ).execute()
    spreadsheet = response.get('files', [])
    spreadsheet_id = spreadsheet[0].get('id')
    return spreadsheet_id

def sheet_values(spreadsheet_id):
    sheet = service_sheet.spreadsheets()
    result = sheet.values().get(spreadsheetId = spreadsheet_id, range = 'Sheet1!A1:D7').execute()
    values = result.get('values', [])
    return values

def data_manipulation(values,date):
    df = pd.DataFrame(values)
    df.columns = df.iloc[0]
    df = df[1:]
    df.reset_index(drop=True, inplace=True)
    date_index = []
    for i in range(0,len(df['Date'])):
        if df['Date'][i] == date:
            date_index.append(i)
    dict_meetId = {}
    for i in date_index:
        dict_meetId["Meeting ID"] = df["Meeting ID"][i]
    dict_passcode = {}
    for i in date_index:
        dict_passcode["Meeting Passcode"] = df["Meeting Passcode"][i]
    meeting_id = dict_meetId.get("Meeting ID")
    meeting_passcode = dict_passcode.get("Meeting Passcode")
    return meeting_id,meeting_passcode

def get_data(date):
    id = get_spreadsheet()
    values = sheet_values(id)
    data = data_manipulation(values,date)
    return data

def meeting_id(data):
    return data[0]

def meeting_password(data):
    return data[1]

def get_date():
    datetime_obj = datetime.now()
    date = datetime_obj.date()
    date = str(date)
    return date

def join_meet_final():
    date = get_date()
    meetid = meeting_id(get_data(date))
    passcode = meeting_password(get_data(date))
    join_meet(meetid,passcode)
join_meet_final()