import time
import webbrowser
import schedule
import pyautogui
import pathlib
import openpyxl
from secrets import *
from converttime import convertpsttoutc
parent = pathlib.Path(__file__).resolve().parent
mypath = pathlib.Path(parent)


# TODO
# implement Start Date and End Date only join meetings during that interval
# add functionality to connect a prerecorded video to the meeting
# Notify user over text when they have joined a meeting and send a screenshot of desktop

def loadexcelfile():
    excelpath = mypath / 'zoom.xlsx'
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
            #print(sheet.cell(row=x, column=y).value)
    return zoomdata


def createschedule(zoomdata):
    #[1] is link
    #[2] is password, prob not needed many classes have code included
    #[3] is time
    #[4-10] monday - sunday
    zoomlinks = []
    zoompasses = []
    zoomtimes = []
    dayslist = []
    for x in range(len(zoomdata)):
        dayslist.append([])
        zoomlinks.append(zoomdata[x][1])
        if zoomdata[x][2] is not None:
            zoompasses.append(zoomdata[x][2])
        else:
            zoompasses.append(-1)
        zoomtimes.append(zoomdata[x][3])
        for i in range(4, 11, 1):
            if zoomdata[x][i] is not None:
                dayslist[x].append(getday(i))
    '''
    print(dayslist)
    print(zoomlinks)
    print(zoompasses)
    '''
    for x in range(len(zoomdata)):
        for i in range(len(dayslist[x])):
            makeschedule(dayslist[x][i], zoomtimes[x],
                         [zoomlinks[x], zoompasses[x]])
    # iterate through list in range(len(zoomdata))
    # for zoom links need to add a check to joinzoommeeting to see if it is link or code and password format


def getday(daynum):
    days = ["MONDAY", "TUESDAY", "WEDNESDAY",
            "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
    return days[daynum - 4]


def makeschedule(day, time, zoomdata):
    print(f"Making Schedule {day}")
    print(str(time))
    if day.upper() == 'MONDAY':
        schedule.every().monday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'TUESDAY':
        schedule.every().tuesday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'WEDNESDAY':
        schedule.every().wednesday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'THURSDAY':
        schedule.every().thursday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'FRIDAY':
        schedule.every().friday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SATURDAY':
        schedule.every().saturday.at(convertpsttoutc(
            str(time))).do(joinzoommeeting, zoomdata)
    elif day.upper() == 'SUNDAY':
        schedule.every().sunday.at(convertpsttoutc(str(time))).do(joinzoommeeting, zoomdata)


def joinzoommeeting(info):
    # info[0] classcode info [1] password if there is one
    # print(info[0],info[1])
    try:
        if info[1] != -1:
            webbrowser.open("https://oregonstate.zoom.us/j/" + str(info[0]))
            loc = pyautogui.locateCenterOnScreen(
                str(mypath / 'passwordbox.png'))
            while loc is None:
                loc = pyautogui.locateCenterOnScreen(
                    str(mypath / 'passwordbox.png'))
        else:
            webbrowser.open(info[0])
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'closesymbol.png')))
        if info[1] != -1:
            pyautogui.click(loc)
            pyautogui.write(info[1])
        time.sleep(3)
        while pyautogui.locateCenterOnScreen(mypath / 'wating.png') is not None:
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
