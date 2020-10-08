import time
import webbrowser
import schedule
import pyautogui
import pathlib
from openpyxl import load_workbook
import re
import logging
import requests
from datetime import datetime
from secrets import *
logging.basicConfig(format='%(message)s', level=logging.DEBUG)
logging.disable()
parent = pathlib.Path(__file__).resolve().parent
mypath = pathlib.Path(parent)


# TODO
# add functionality for consenting to being recorded
# implement Start Date and End Date only join meetings during that interval
# add functionality to connect a prerecorded video to the meeting
# Notify user over text when they have joined a meeting and send a screenshot of desktop
# add name of class to zoomdata so can text user what class they have joined
def loadexcelfile():
    excelpath = mypath / 'docs' / 'schedule.xlsx'
    wb = load_workbook(excelpath)
    sheet = wb['Sheet1']
    numofcols = sheet.max_column
    numofrows = sheet.max_row
    zoomdata = []
    loopdata = []
    for x in range(2, numofrows + 1):
        if sheet.cell(row=x, column=1).value is not None:
            zoomdata.append([])
        for y in range(1, numofcols + 1):
            cellvalue = sheet.cell(row=x, column=y).value
            if y == 1 and cellvalue is None:
                break
            else:
                zoomdata[x - 2].append(cellvalue)
    return zoomdata


def createschedule(zoomdata):
    #[1] is link
    #[2] is password, prob not needed many classes have code included
    #[3] is time
    #[4-10] monday - sunday
    zoomlinks = []
    zoompasses = []
    zoomtimes = []
    meetingnames = []
    dayslist = []
    for x in range(len(zoomdata)):
        meetingnames.append(zoomdata[x][0])
        dayslist.append([])
        zoomlinks.append(zoomdata[x][1].strip())
        if zoomdata[x][2] is not None:
            zoompasses.append(zoomdata[x][2])
        else:
            zoompasses.append(-1)
            # if no password we have only link then regex for pass and code to make a better link
            # p = re.compile("(\d{11}).*pwd=(.*)&")
            p = re.compile("(\d{10,11}).*pwd=(.{32})")
            m = p.search(zoomlinks[x])
            if m:
                # creates a direct join link to zoom application
                # another option is to be able to forgo application entirely with
                # link this https://www.zoom.us/wc/join/94165984842?pwd=ME1OemMrdHdUSElGaXdobkN4Z2NzQT09
                # which opens browser version of zoom
                # could have an option but might be confusing to many
                if m.group(1) and m.group(2):
                    zoomlinks[x] = "zoommtg://www.zoom.us/join?action=join&confno=" + \
                        m.group(1) + "&pwd=" + m.group(2)
                # zoomlinks[x] = "https://www.zoom.us/wc/join/" + m.group(1) + "?pwd=" + m.group(2)

        zoomtimes.append(str(zoomdata[x][3]))
        for i in range(4, 11, 1):
            if zoomdata[x][i] is not None:
                dayslist[x].append(getday(i))
    for x in range(len(zoomdata)):
        for i in range(len(dayslist[x])):
            setschedule(dayslist[x][i], zoomtimes[x], [
                zoomlinks[x], zoompasses[x], meetingnames[x]])
    logging.debug(zoomlinks)


def getday(daynum):
    days = ["MONDAY", "TUESDAY", "WEDNESDAY",
            "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
    return days[daynum - 4]


def setschedule(day, time, zoomdata):
    splittime = time.split(":")
    time = f"{splittime[0]}:{splittime[1]}"
    if day.upper() == 'MONDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().monday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'TUESDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().tuesday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'WEDNESDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().wednesday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'THURSDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().thursday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'FRIDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().friday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SATURDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().saturday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SUNDAY':
        print(
            f"Setting schedule for {zoomdata[2].upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().sunday.at(str(time)).do(joinzoommeeting, zoomdata)


def joinzoommeeting(info):
    # info[0] classcode info [1] password if there is one
    # print(info[0],info[1])
    print("Trying to join meeting")
    pyautogui.hotkey('winleft', 'm')
    try:
        if info[1] != -1:
            webbrowser.open("https://www.zoom.us/j/" + str(info[0]))
            loc = pyautogui.locateCenterOnScreen(
                str(mypath / 'passwordbox.png'))
            while loc is None:
                loc = pyautogui.locateCenterOnScreen(
                    str(mypath / 'passwordbox.png'))
            time.sleep(1)
            pyautogui.click(pyautogui.locateCenterOnScreen(
                str(mypath / 'closesymbol.png')))
            pyautogui.click(loc)
            pyautogui.write(info[1])
        else:
            webbrowser.open(info[0])
            time.sleep(3)
        time.sleep(3)
        while pyautogui.locateCenterOnScreen(str(mypath / 'waiting.png')) is not None:
            time.sleep(1)
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'joinmeeting.png')))
        time.sleep(5)
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'joincomaud.png')))
        time.sleep(3)
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'mute.png')))
        time.sleep(2)
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'fullscreen.png')))
        winlist = pyautogui.getAllTitles()
        win = pyautogui.getWindowsWithTitle('Zoom Meeting')
        if 'Zoom Meeting' in winlist:
            win[0].maximize()
            sendmessage(
                f'\U00002705 You have joined your meeting: {info[2]}')
        elif 'Waiting for Host' in winlist:
            sendmessage(
                f'\U00002705 Waiting for host to start the meeting: {info[2]}')
        else:
            sendmessage(
                f'\U0000274C ERROR: may have not joined meeting: {info[2]}')
        senddesktopscreenshot()

    except IndexError:
        pyautogui.alert("Error data is not correctly entered")
        print("The code/password is not present")


def senddesktopscreenshot():
    if 'api_key' in globals():
        im = pyautogui.screenshot('scrshot.png')
        url = f'https://api.telegram.org/bot{api_key}/sendPhoto'
        data = {'chat_id': chat_id}
        files = {'photo': open('scrshot.png', 'rb')}
        r = requests.post(url, files=files, data=data)


def sendmessage(mymessage):
    if 'api_key' in globals():
        requests.get(f'https://api.telegram.org/bot{api_key}/sendMessage',
                     params={'chat_id': {chat_id},
                             'text': {mymessage}})


def main():
    zoomdata = loadexcelfile()
    createschedule(zoomdata)
    newlines = ''
    for x in range(40):
        newlines += "." + "\n"
    sendmessage(newlines)
    sendmessage("Bot has started")
    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        logging.debug(f"Current Time = {current_time}")
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()
