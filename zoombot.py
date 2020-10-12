import time
import webbrowser
import schedule
import pyautogui
import pathlib
import os
from openpyxl import load_workbook
import re
import logging
import requests
import configparser
import pyscreenshot as imggrab
from datetime import datetime
from secrets import *
from subprocess import Popen, PIPE
from telegram.ext import Updater, CommandHandler, MessageHandler, Filters
chat_id = ''
api_key = ''

# Enable logging
logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    level=logging.INFO)
logger = logging.getLogger(__name__)
# logging.disable()
parent = pathlib.Path(__file__).resolve().parent
mypath = pathlib.Path(parent)
config = configparser.ConfigParser()

# TODO
# add functionality for consenting to being recorded
# implement Start Date and End Date only join meetings during that interval
# add functionality to connect a prerecorded video to the meeting
# add config to add telegram userid, can be found by messaging @userinfobot
# api_key can stay in globals within secrets


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
    schedule_message = ""
    for x in range(len(zoomdata)):
        for i in range(len(dayslist[x])):
            setschedule(dayslist[x][i], zoomtimes[x], [
                zoomlinks[x], zoompasses[x], meetingnames[x]])
        splittime = zoomtimes[x].split(":")
        time = f"{splittime[0]}:{splittime[1]}"
        t = datetime.strptime(time, "%H:%M")
        timevalue_12hour = t.strftime("%I:%M %p")
        schedule_message += f"Scheduling {meetingnames[x].upper()} meeting on {', '.join([x.capitalize() for x in dayslist[x]])} joining at {str(timevalue_12hour)} " + "\n\n"
    logging.debug(zoomlinks)
    sendmessage(schedule_message)


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
                f'\U00002705 You have successfully joined your meeting: {info[2].upper()}')
        elif 'Waiting for Host' in winlist:
            sendmessage(
                f'\U000023F3 Waiting for host to start the meeting: {info[2].upper()}')
        else:
            sendmessage(
                f'\U0000274C ERROR: may have not joined meeting: {info[2].upper()}')
        senddesktopscreenshot()
        print("Finished joining Meeting")
    except IndexError:
        pyautogui.alert("Error data is not correctly entered")
        print("The code/password is not present")


def senddesktopscreenshot():
    if 'api_key' in globals():
        img = imggrab.grab()
        saveas = ('scrshot.png')
        img.save(saveas)
        #im = pyautogui.screenshot('scrshot.png')
        url = f'https://api.telegram.org/bot{api_key}/sendPhoto'
        data = {'chat_id': chat_id}
        files = {'photo': open(str(mypath / 'scrshot.png'), 'rb')}
        r = requests.post(url, files=files, data=data)


def sendmessage(mymessage):
    if 'api_key' in globals():
        requests.get(f'https://api.telegram.org/bot{api_key}/sendMessage',
                     params={'chat_id': {chat_id},
                             'text': {mymessage}})


def makeconfig():
    if not os.path.exists('config.ini'):
        print("Making config file")
        config['Telegram Info'] = {'userid': '0',
                                   'api_key': '0'}
        with open('config.ini', 'w') as configfile:
            config.write(configfile)
    else:
        config.read('config.ini')


def openzoom(update, context):
    if str(update.message.chat_id) == str(chat_id):
        update.message.reply_text("Trying to Open Zoom")
        proc = Popen(r'C:\Users\James\AppData\Roaming\Zoom\bin\zoom.exe')
        time.sleep(3)
        senddesktopscreenshot()


def screenshot(update, context):
    if str(update.message.chat_id) == str(chat_id):
        senddesktopscreenshot()


def help(update, context):
    """send help message."""
    if str(update.message.chat_id) == str(chat_id):
        update.message.reply_text('''There are a few commands with this bot\nIf you type /screen you will be sent a picture of your desktop
If you type /openzoom it will openzoom''')
    else:
        update.message.reply_text('''Unauthenticated User''')

def cs(update, context):
    if str(update.message.chat_id) == str(chat_id):
        update.message.reply_text("Trying to Open cs accepter")
        pyautogui.press('winleft')
        time.sleep(.5)
        pyautogui.write('accepter')
        time.sleep(.5)
        pyautogui.press('enter')
        time.sleep(2)
        d = pyautogui.getWindowsWithTitle('Counter-Strike: Global Offensive')
        if len(d) > 0:
            d[0].maximize()
        senddesktopscreenshot()

def shutit(update,context):
    if str(update.message.chat_id) == str(chat_id):
        os.system(f'shutdown /s /t 0')

def main():
    makeconfig()
    global chat_id
    global api_key
    chat_id = config['Telegram Info']['userid']
    api_key = config['Telegram Info']['api_key']
    updater = Updater(api_key, use_context=True)

    # Get the dispatcher to register handlers
    dp = updater.dispatcher

    # on different commands - answer in Telegram
    dp.add_handler(CommandHandler("openzoom", openzoom))
    dp.add_handler(CommandHandler("screen", screenshot))
    dp.add_handler(CommandHandler("cs", cs))
    #dp.add_handler(CommandHandler("shutdown", shutit))
    # on noncommand i.e message - echo the message on Telegram
    dp.add_handler(MessageHandler(Filters.text & ~Filters.command, help))

    zoomdata = loadexcelfile()
    newlines = ''
    for x in range(40):
        newlines += "." + "\n"
    sendmessage(newlines)
    sendmessage("Bot has started")
    createschedule(zoomdata)
    updater.start_polling()

    while True:
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        logging.debug(f"Current Time = {current_time}")
        schedule.run_pending()
        time.sleep(1)


if __name__ == '__main__':
    main()
