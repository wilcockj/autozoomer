import time
import webbrowser
import schedule
import pyautogui
import pathlib
import openpyxl
import re
from secrets import *
from converttime import convertpsttoutc
parent = pathlib.Path(__file__).resolve().parent
mypath = pathlib.Path(parent)


# TODO
# implement Start Date and End Date only join meetings during that interval
# add functionality to connect a prerecorded video to the meeting
# Notify user over text when they have joined a meeting and send a screenshot of desktop

def loadexcelfile():
    excelpath = mypath / 'docs' / 'zoom.xlsx'
    wb = openpyxl.load_workbook(excelpath)
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
            p = re.compile("(\d{11}).*pwd=(.*)&")
            m = p.search(zoomlinks[x])
            if m.group(2):
                # creates a direct join link to zoom application
                # another option is to be able to forgo application entirely with
                # a link link this https://www.zoom.us/wc/join/94165984842?pwd=ME1OemMrdHdUSElGaXdobkN4Z2NzQT09
                # which opens browser version of zoom
                # could have an option but might be confusing to many
                zoomlinks[x] = "zoommtg://www.zoom.us/join?action=join&confno=" + \
                    m.group(1) + "&pwd=" + m.group(2)
                #zoomlinks[x] = "https://www.zoom.us/wc/join/" + m.group(1) + "?pwd=" + m.group(2)

        zoomtimes.append(zoomdata[x][3])
        for i in range(4, 11, 1):
            if zoomdata[x][i] is not None:
                dayslist[x].append(getday(i))
    for x in range(len(zoomdata)):
        for i in range(len(dayslist[x])):
            makeschedule(meetingnames[x], dayslist[x][i], zoomtimes[x], [
                         zoomlinks[x], zoompasses[x]])


def getday(daynum):
    days = ["MONDAY", "TUESDAY", "WEDNESDAY",
            "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
    return days[daynum - 4]


def makeschedule(meetingname, day, time, zoomdata):
    # something odd is going on with the time sometimes I have to convert to utc other times time is localtime by default
    if day.upper() == 'MONDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().monday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'TUESDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().tuesday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'WEDNESDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().wednesday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'THURSDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().thursday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'FRIDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().friday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SATURDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().saturday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SUNDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().sunday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)
    '''
    if day.upper() == 'MONDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().monday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'TUESDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().tuesday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'WEDNESDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().wednesday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'THURSDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().thursday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'FRIDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().friday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SATURDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().saturday.at(str(time)).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SUNDAY':
        print(
            f"Setting schedule for {meetingname.upper()} on {day.capitalize()} joining conference at {str(time)}")
        schedule.every().sunday.at(str(time)).do(joinzoommeeting, zoomdata)
    '''


def joinzoommeeting(info):
    # info[0] classcode info [1] password if there is one
    # print(info[0],info[1])
    print("Trying to join meeting")
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
    except IndexError:
        pyautogui.alert("Error data is not correctly entered")
        print("The code/password is not present")


def main():
    zoomdata = loadexcelfile()
    createschedule(zoomdata)
    while True:
        time.sleep(1)
        schedule.run_pending()


if __name__ == '__main__':
    main()
