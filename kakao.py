from asyncio.windows_events import NULL
import ctypes
import time
from tkinter import END
import win32api
import win32con
import win32gui
import random
import datetime
import re
from firebase_admin import credentials, db, initialize_app
import os

import pandas as pd
from pywinauto import clipboard

PROGRAM_ON_OFF = True
KAKAO_ROOM_NAME = '송상윤'
MAX_REVEIVED_COUNT = 3
TODAY_RECEIVE_COUNT = 0

START_RECEVED_HOUR = 10
END_RECEIVED_HOUR = 22
current_hour = datetime.datetime.now().hour
TODAY_TIME_OUT_INTERVAL_SECOND = (END_RECEIVED_HOUR - START_RECEVED_HOUR) * 60 * 60
TODAY_COUNT_OUT_INTERVAL_SECOND = 10
DETECTION_INTERVER_SECOND = 3

cred = credentials.Certificate(str(os.getcwd()) + '\\mykey.json')
initialize_app(cred,{
    'databaseURL' : 'https://kakao-auto-program-default-rtdb.firebaseio.com/' 
})
firebase_db = db.reference("controll")

    # 예외처리필요
    #"[당일 마감 모집]",
DETECTION_SENTENCES = [
    "양도받으실분있으신가요",
    "양도받을분있나용",
    "양도받을분있나요",
    "양도받을분있으신가요",
    "양도받을분있으신가용",
    "양도받을분구해요",
    "양도받을분구해용",
    "양도받을분구합니다",
    "양도받을사람구해요",
    "양도받을사람구해용",
    "양도받을사람구합니다",
    "양도받을분있으실까요",
    "양도받을분있으실까용",
    "양도받아주실분있으실까요",
    "양도받아주실분있으실까용",
    
    "양도받으실분있으실까요",
    "양도받으실분있으실까용",
    "양도받으실분있을까요",
    "양도받으실분있을까용",
    "양도받으실분있나요",
    "양도받으실분구합니다",  
    "양도받으실분구해요",
    "양도받으실분구해용",

    "양도받으실분계세요", 
    "양도받으실분계세용", 
    "양도받으실분계신가요", 
    "양도받으실분계신가용", 
    "양도받으실분계신지요", 
    "양도받으실분계신지용", 
    "양도받으실분계실까요",  
    "양도받으실분계실까용",  

    "양도받아주실분있나요",
    "양도받아주실분있나용",
    "양도받아주실분있으실까요",
    "양도받아주실분있으실까용",

    "양도하려고하는데하실분계실까요",
    "양도하려고하는데하실분계실까용",
    "양도합니다",
    "양도해요",
    "양도하려합니다",
    "양도자찾습니다",
    "양도자찾습니다:)",
    "양수하실분계실까요",
    "양수하실분계실까용"
    ]

AGREE_MULTI_ANSWER_IST = [
    "넵모두진행할게요",
    "모두해봉게여",
    "넵다해보렏습니다",
    "네다해볼게요",
    "넵다진행하겟슴ㅁ미다",
    "모두ㄱㅏ능합니다",
    "모두ㅜ가능해요",
    "넵 다진행할게요",
    "모두하겟습미다"
]

AGREE_SINGLE_ANSWER_LIST = [
    "넵제가하겟습미다",
    "네제가발행할게여",
    "넵해보렏습니다~",
    "넵진행요~",
    "해볼개요~",
    "해볼게요",
    "네진행해볼게요",
    "넵 진ㄴ행하겟슴ㅁ미다",
    "넵진행할게요"
    ]

SHEET_CHANGE_ANSWER_LIST = [
    "시트도바꿔놓을게요~",
    "시트 변경도 해놓을게요~", 
    "시트는 제가 변경할게요~", 
    "시트는 제가 바꿔놓겠습니다~", 
    "시트도 바꿔 진행할게요~", 
    "시트도 변경해 진행할게요~", 
    "시트도 변경하여 진행하겠습니다~", 
    "시트도 함께 바꿔놓을게요~"
    ]

PBYTE256 = ctypes.c_ubyte * 256
_user32 = ctypes.WinDLL("user32")
GetKeyboardState = _user32.GetKeyboardState
SetKeyboardState = _user32.SetKeyboardState
PostMessage = win32api.PostMessage
SendMessage = win32gui.SendMessage
FindWindow = win32gui.FindWindow
IsWindow = win32gui.IsWindow
GetCurrentThreadId = win32api.GetCurrentThreadId
GetWindowThreadProcessId = _user32.GetWindowThreadProcessId
AttachThreadInput = _user32.AttachThreadInput

MapVirtualKeyA = _user32.MapVirtualKeyA
MapVirtualKeyW = _user32.MapVirtualKeyW

MakeLong = win32api.MAKELONG
w = win32con


def send_kakao_talk(text):
    hwndMain = win32gui.FindWindow( None, KAKAO_ROOM_NAME)
    hwndEdit = win32gui.FindWindowEx( hwndMain, None, "RichEdit50W", None)

    print("================카톡 전송함================ \n전송 내용 : ", text, "\n===========================================\n")
    
    win32api.SendMessage(hwndEdit, win32con.WM_SETTEXT, 0, text)
    time.sleep(1)
    PressEnter(hwndEdit)


def get_chat_contents():
    hwndMain = win32gui.FindWindow( None, KAKAO_ROOM_NAME)
    hwndListControl = win32gui.FindWindowEx( hwndMain, None, "EVA_VH_ListControl_Dblclk", None)

    postKeyEx(hwndListControl, ord('A'), [w.VK_CONTROL], False)
    postKeyEx(hwndListControl, ord('C'), [w.VK_CONTROL], False)
    clipborad_text = clipboard.GetData()
    return clipborad_text


def postKeyEx(hwnd, key, shift, specialkey):
    if IsWindow(hwnd):

        ThreadId = GetWindowThreadProcessId(hwnd, None)

        lparam = MakeLong(0, MapVirtualKeyA(key, 0))
        msg_down = w.WM_KEYDOWN
        msg_up = w.WM_KEYUP

        if specialkey:
            lparam = lparam | 0x1000000

        if len(shift) > 0:
            pKeyBuffers = PBYTE256()
            pKeyBuffers_old = PBYTE256()

            SendMessage(hwnd, w.WM_ACTIVATE, w.WA_ACTIVE, 0)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, True)
            GetKeyboardState(ctypes.byref(pKeyBuffers_old))

            for modkey in shift:
                if modkey == w.VK_MENU:
                    lparam = lparam | 0x20000000
                    msg_down = w.WM_SYSKEYDOWN
                    msg_up = w.WM_SYSKEYUP
                pKeyBuffers[modkey] |= 128

            SetKeyboardState(ctypes.byref(pKeyBuffers))
            time.sleep(0.01)
            PostMessage(hwnd, msg_down, key, lparam)
            time.sleep(0.01)
            PostMessage(hwnd, msg_up, key, lparam | 0xC0000000)
            time.sleep(0.01)
            SetKeyboardState(ctypes.byref(pKeyBuffers_old))
            time.sleep(0.01)
            AttachThreadInput(GetCurrentThreadId(), ThreadId, False)

        else:
            SendMessage(hwnd, msg_down, key, lparam)
            SendMessage(hwnd, msg_up, key, lparam | 0xC0000000)


def PressEnter(hwnd):
    win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
    time.sleep(1)
    win32api.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)


def get_last_chat():
    chat_contents = get_chat_contents()  

    list_chat_contents = chat_contents.split('\r\n')   # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨

    data_frame_chat_contents = pd.DataFrame(list_chat_contents)

    data_frame_chat_contents[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')

    return data_frame_chat_contents.index[-2], data_frame_chat_contents.iloc[-2, 0]


def get_campaign_number(chat_contents):
    return len(re.findall(r'\d{4}.', chat_contents))

def is_multi_compaign(chat_contents):
    campaign_number = get_campaign_number(chat_contents)

    if (campaign_number > 1):
        return True
    else:
        return False


def detect_keyword_sentence(last_raw_chat_index, last_raw_chat_contents):

    chat_conents = get_chat_contents() 

    list_chat_contents = chat_conents.split('\r\n')  # \r\n 으로 스플릿 __ 대화내용 인용의 경우 \r 때문에 해당안됨
    
    data_frame_chat_contents = pd.DataFrame(list_chat_contents)

    data_frame_chat_contents[0] = data_frame_chat_contents[0].str.replace('\[([\S\s]+)\] \[(오전|오후)([0-9:\s]+)\] ', '')

    if data_frame_chat_contents.iloc[-2, 0] == last_raw_chat_contents:
        print("================채팅 없었음================ \n마지막 채팅 인덱스 : ", last_raw_chat_index, "\n마지막 채팅 내용 : ", last_raw_chat_contents, "\n===========================================\n")
    else:
        print("================채팅 있었음================ \n마지막 채팅 인덱스 : ", last_raw_chat_index, "\n마지막 채팅 내용 : ", last_raw_chat_contents, "\n===========================================\n")

        last_raw_chat_contents = data_frame_chat_contents.iloc[last_raw_chat_index+1 : , 0]

        last_raw_chat_contents_line_str = str(last_raw_chat_contents.values[0].replace(" ", "").replace("\n", ""))
        
        is_compaign_detected = False
        for i in range(len(DETECTION_SENTENCES)):
            if (DETECTION_SENTENCES[i] in last_raw_chat_contents_line_str):
                is_compaign_detected = True
                break
            else:
                is_compaign_detected = False

        if is_compaign_detected:
            if (is_multi_compaign(last_raw_chat_contents_line_str)):
                answer = random.choice(AGREE_MULTI_ANSWER_IST) 
            else:
                answer = random.choice(AGREE_SINGLE_ANSWER_LIST) 
            send_kakao_talk(answer)

            time.sleep(4)
            
            answer = random.choice(SHEET_CHANGE_ANSWER_LIST) 
            send_kakao_talk(answer)  

            received_campaign_number = get_campaign_number(last_raw_chat_contents_line_str)
            plus_today_receive_count_to_firebase(received_campaign_number)

            print("================키워드 확인!================\n 마지막 채팅 인덱스 : ", last_raw_chat_index, "\n마지막 채팅 내용 : ", last_raw_chat_contents, "\n===========================================\n")

        else:
            print("================키워드 미확인!================\n 마지막 채팅 인덱스 : ", last_raw_chat_index, "\n마지막 채팅 내용 : ", last_raw_chat_contents, "\n===========================================\n")
            
    return data_frame_chat_contents.index[-2], data_frame_chat_contents.iloc[-2, 0]


def plus_today_receive_count_to_firebase(received_campaign_number):
    global TODAY_RECEIVE_COUNT
    TODAY_RECEIVE_COUNT = int(firebase_db.child("TODAY_RECEIVE_COUNT").get()) + received_campaign_number
    
    firebase_db.update({'TODAY_RECEIVE_COUNT' : TODAY_RECEIVE_COUNT})


def zero_today_receive_count_to_firebase():
    firebase_db.update({'TODAY_RECEIVE_COUNT' : 0})


def firebase_listener(event):    

    global PROGRAM_ON_OFF
    global KAKAO_ROOM_NAME
    global MAX_REVEIVED_COUNT
    global START_RECEVED_HOUR
    global END_RECEIVED_HOUR
    global TODAY_RECEIVE_COUNT

    print("================전역변수변경================ \n")
    print("data : ", event.data, "\n")

    for key, value in event.data.items():
        if key == "PROGRAM_ON_OFF":
            print("PROGRAM_ON_OFF : ", PROGRAM_ON_OFF, " -> ", value, "\n")
            PROGRAM_ON_OFF = value

        if key == "KAKAO_ROOM_NAME":
            print("KAKAO_ROOM_NAME : ", KAKAO_ROOM_NAME, " -> ", value, "\n")
            KAKAO_ROOM_NAME = value

        if key == "MAX_REVEIVED_COUNT":
            print("MAX_REVEIVED_COUNT : ", MAX_REVEIVED_COUNT, " -> ", value, "\n")
            MAX_REVEIVED_COUNT = value

        if key == "START_RECEVED_HOUR":
            print("START_RECEVED_HOUR : ", START_RECEVED_HOUR, " -> ", value, "\n")
            START_RECEVED_HOUR = value

        if key == "END_RECEIVED_HOUR":
            print("END_RECEIVED_HOUR : ", END_RECEIVED_HOUR, " -> ", value, "\n")
            END_RECEIVED_HOUR = value

        if key == "TODAY_RECEIVE_COUNT":
            print("TODAY_RECEIVE_COUNT : ", TODAY_RECEIVE_COUNT, " -> ", value, "\n")
            TODAY_RECEIVE_COUNT = value

    print("===========================================\n")

 
# 로직 플로우 차트 : https://drive.google.com/file/d/1UVD8wxVqSgF89pPqWXAi4iAYhLtzF2D9/view?usp=sharing
def main():

    firebase_db.listen(firebase_listener)

    print("================가동 시작================")

    last_raw_chat_index, last_raw_chat_contents = get_last_chat() 

    global TODAY_RECEIVE_COUNT
    while True:
        if PROGRAM_ON_OFF == True:
            if (current_hour >= START_RECEVED_HOUR and current_hour < END_RECEIVED_HOUR):
                if (TODAY_RECEIVE_COUNT < MAX_REVEIVED_COUNT):
                    print("================양도 받는중================\n", TODAY_RECEIVE_COUNT, "2 : ", MAX_REVEIVED_COUNT)
                    last_raw_chat_index, last_raw_chat_contents = detect_keyword_sentence(last_raw_chat_index, last_raw_chat_contents)
                    time.sleep(DETECTION_INTERVER_SECOND)
                else:
                    print("================양도 최대 일 갯수 초과================\n", TODAY_RECEIVE_COUNT, "2 : ", MAX_REVEIVED_COUNT)
                    time.sleep(TODAY_COUNT_OUT_INTERVAL_SECOND)
                    continue
            else:
                print("================양도 최대 시간 초과================\n")
                TODAY_RECEIVE_COUNT = 0
                zero_today_receive_count_to_firebase()
                time.sleep(TODAY_TIME_OUT_INTERVAL_SECOND)
                continue
        else:
            print("================프로그램 꺼짐================\n")
            time.sleep(300)
            continue

        

if __name__ == '__main__':
    main()