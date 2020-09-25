import time
import webbrowser
import schedule
import pyautogui
import pathlib
from secrets import *
parent = pathlib.Path(__file__).resolve().parent
mypath = pathlib.Path(parent)

# passcode is 678065
# TODO make functionality so that maybe through a gui
# or some other sort of config you can import classes and
# generate the code
# once I can make my schedule secret I will make public


def joinzoommeeting(info):
    # info[0] classcode info [1] password if there is one
    # print(info[0],info[1])
    try:
        webbrowser.open("https://oregonstate.zoom.us/j/" + str(info[0]))
        loc = pyautogui.locateCenterOnScreen(str(mypath / 'passwordbox.png'))
        while loc is None:
            loc = pyautogui.locateCenterOnScreen(
                str(mypath / 'passwordbox.png'))
        pyautogui.click(pyautogui.locateCenterOnScreen(
            str(mypath / 'closesymbol.png')))
        pyautogui.click(loc)
        pyautogui.write(info[1])
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
# print(convertpsttoutc("16:20"))
# here I will make arrays first element is classcode second is the class password or -1 if there isnt
# remove after 9/25


def joinzoomwebinar():
    webbrowser.open("https://oregonstate.zoom.us/w/99451242990?tk=DZbeEnoo9h6O8uzCbavHjUkrjxta2UyLyxDJZhDRAtE.DQIAAAAXJ8GJ7hZFMFpNV1dYMlJrU3NMOFlxdDh1eFdnAAAAAAAAAAAAAAAAAAAAAAAAAAAA&uuid=WN_rgFMBjX3T76Ue-rftu6lsw")
    time.sleep(5)
    pyautogui.click(pyautogui.locateCenterOnScreen(
        str(mypath / 'closesymbol.png')))
    time.sleep(2)
    pyautogui.click(pyautogui.locateCenterOnScreen(
        str(mypath / 'joinmeeting.png')))
    time.sleep(5)
    pyautogui.click(pyautogui.locateCenterOnScreen(
        str(mypath / 'joincomaud.png')))
    time.sleep(3)
    pyautogui.click(pyautogui.locateCenterOnScreen(str(mypath / 'mute.png')))
    time.sleep(2)
    pyautogui.click(pyautogui.locateCenterOnScreen(
        str(mypath / 'fullscreen.png')))


schedule.every().monday.at("11:57").do(
    joinzoommeeting, sus102)  # set 1 SUS102 MWF
schedule.every().monday.at("14:57").do(
    joinzoommeeting, mth264)  # set 2 MTH264 MWF
schedule.every().tuesday.at("08:57").do(
    joinzoommeeting, ph212)  # set 3 PH 212 Tues Thurs
# schedule.every().tuesday.at("09:58").do(joinzoommeeting, eng104)  # set 4 Eng 104 Tues Thurs
schedule.every().tuesday.at("13:57").do(
    joinzoommeeting, ph212rec)  # PH 212 studio
schedule.every().tuesday.at("15:57").do(
    joinzoommeeting, hst104)  # set 5 HST 104 Tues Thurs
schedule.every().tuesday.at("17:27").do(
    joinzoommeeting, mth264rec)  # MTH264 rec
schedule.every().wednesday.at("11:57").do(joinzoommeeting, sus102)  # set 1
schedule.every().wednesday.at("14:57").do(joinzoommeeting, mth264)  # set 2
schedule.every().wednesday.at("15:58").do(
    joinzoommeeting, sus102lab)  # SUS 102 lab
schedule.every().thursday.at("08:57").do(joinzoommeeting, ph212)  # set 3
# schedule.every().thursday.at("09:58").do(joinzoommeeting, eng104)  # set 4
# schedule.every().thursday.at("11:57").do(joinzoommeeting, [])  # PH212 lab ? so optional or meeting group
schedule.every().thursday.at("15:57").do(joinzoommeeting, hst104)  # set 5
schedule.every().friday.at("11:57").do(joinzoommeeting, sus102)  # set 1
schedule.every().friday.at("14:57").do(joinzoommeeting, mth264)  # set 2
schedule.every().friday.at("10:55").do(joinzoomwebinar)
# schedule.every().second.do(joinzoommeeting,sus102)
# 11:55am, 2:55pm monday
# 9:00am, 9:57am ,1:55pm, 3:55pm, 5:25pm tues
# 11:55am, 2:55pm, 3:55pm weds
# 8:55am, 9:55am, 11:55am,3:55pm thurs
# 11:55am, 2:55pm friday
# webbrowser.open("https://zoom.us/")
# joinzoommeeting(hst104)


def main():
    while True:
        time.sleep(1)
        schedule.run_pending()


if __name__ == '__main__':
    main()
