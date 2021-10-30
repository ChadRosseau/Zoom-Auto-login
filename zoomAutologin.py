# Necessary pip imports: "pip install pyobjc pyautogui vlc playsound"

# Necessary imports
from AppKit import NSWorkspace
from email.header import decode_header
from getpass import getpass
from playsound import playsound
from pyautogui import press, typewrite, hotkey, keyDown, keyUp

import base64
import datetime as dt
import email
import imaplib
import importlib
import linecache
import math
import os
import platform
import pyperclip
import re
import smtplib
import subprocess
import sys
import time
import webbrowser

# Create classTime global arrays
classTimes = []
warningTimes = []
joinTimes = []

warningMessages = [
    "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n",
    "                                ZOOM AUTO-CONNECT PROGRAM\n",
    "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n",
    "                                PLEASE READ THE FOLLOWING\n",
    "1. ENSURE ENTERED TIMES ARE IN 24H TIME. 12H TIMES WILL NOT WORK DURING THE AFTERNOON.",
    "2. For this program, you will be asked for your email address and password.\nThese are not recorded anywhere, they are used simply to retrieve the Google Calendar emails from your inbox.",
    "3. For this program, you will be asked to enter 'offsets' for a warning notification and for joining the meeting.\nAn example is '5' for 5 minutes before the class time.\nPlease note that the resulting time you select should not be before the time that Google Calendar sends you the email containing the meet information.\nIf it does, the program will not work as desired.\nThis setting can be adjusted in your Google Calendar, to get earlier emails for your meets if necessary.\n",
    "------------------------------------------------------------------------------------------------------------------------------------------------------------------------------\n",
]

for i in warningMessages:
        print(i)
userAgrees = input("DO YOU AGREE TO THE ABOVE (Y/N)\n\n")
if re.search("^no?$", userAgrees, re.IGNORECASE):
    sys.exit("User Agreement not accepted")
# Account Credentials
FROM_EMAIL = input("Input your email: \n")
FROM_PWD = getpass("Input your email password (Characters will not be shown for security reasons): \n")
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

# Attempt first login
try: 
    mail = imaplib.IMAP4_SSL(SMTP_SERVER)
    mail.login(FROM_EMAIL,FROM_PWD)
    mail.select('inbox')
except Exception as e:
    sys.exit(e)

# Account data
print("Input class times in HH:MM format (Ex. '05:10'). When done, press enter with no input.\nClasstimes do NOT need to be in order.")
classInput = "NULL"
while classInput != "":
    classInput = input("Add time: ")
    if classInput != "":
        classTimes.append(classInput)
print(f"Inputted times: ", end="")
for time in classTimes:
    print(f"{time} ", end="")
print()

warningOffset = int(input("Input no. of minutes before class you want to recieve a warning notification: ('0' for no warning)\n"))
joinOffset = int(input("Input no. of minutes before class you want to automatically join: \n"))
print("Credentials received, please wait for class")

# Class information
global neededEmailSubject
global neededEmailCode
global neededEmailPassword

# Store original app being used
activeAppName = ""





def main():
    # Notify user of script beginning
    notify("Script started", "Class auto-login script has begun")

    # Get current Datetime
    t = dt.datetime.now()
    # Populate warning and joining time arrays
    editTimes(warningOffset, joinOffset)

    editTimes(warningOffset, joinOffset)
    for time in joinTimes:
        print(f"{time} ", end="")
    print()

    time = f"{str(t.hour)}:{str(t.minute)}"
    checkTimes(time)

    # Save the current time to a variable ('t')
    while True:
        delta = dt.datetime.now()-t
        if delta.seconds >= 60:
            # Update 't' variable to new time
            t = dt.datetime.now()
            time = f"{str(t.hour)}:{str(t.minute)}"
            # Assign time to variable
            checkTimes(time)

# -------------------------------------------------
#
# Utility to read email from Gmail Using Python
#
# ------------------------------------------------

def PrintException():
    exc_type, exc_obj, tb = sys.exc_info()
    f = tb.tb_frame
    lineno = tb.tb_lineno
    filename = f.f_code.co_filename
    linecache.checkcache(filename)
    line = linecache.getline(filename, lineno, f.f_globals)
    print('EXCEPTION IN ({}, LINE {} "{}"): {}'.format(filename, lineno, line.strip(), exc_obj))

def read_email_from_gmail():
    # Try access email
    try:
        # Login to user email and go to inbox
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        # Search with no filters
        type, data = mail.search(None, 'ALL')
        mail_ids = data[0]

        # Get id's for each email
        id_list = mail_ids.split()   
        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        # Tracking Variables
        foundCalendarEmail = False

        # For every email between the latest email and the first email.
        for i in range(latest_email_id,first_email_id, -1):

            # Break if there is already a calendar email
            if foundCalendarEmail == True:
                break

            # Fetch the email's data 
            typ, data = mail.fetch(str(i), '(RFC822)')

            # Finds email with Google Calendar in subject
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1].decode())
                    email_subject = msg['subject']
                    email_from = msg['from']
                    email_body = msg.get_payload()
                    print(f"Searching email {i} from: {email_from}")
                    if "Google Calendar" in email_from or "Chad Rossouw" in email_from:
                        print(f"\n\nGoogle Calendar email found\nFrom: {email_from}\nSubject: {email_subject}\n")
                    	# print email_body[0]
                    	# print 'Body : ' + email_body + '\n'
                        baseEmailSender = email_from
                        baseEmailSubject = email_subject
                        foundCalendarEmail = True

        if foundCalendarEmail == True:
            # Get email information by splitting up data
            global neededEmailSubject
            neededEmailSubject = baseEmailSubject.split(" @", 1)[0]
            if isinstance(email_body, list) == True:
                baseEmailBody = purify(str(email_body[1]), "=\n")
            else:
                baseEmailBody = purify(str(email_body), "=\n")

            if "https://us04web.zoom.us/j/" in baseEmailBody:
                baseEmailBody = baseEmailBody.split('https://us04web.zoom.us/j/')[1]
            else:
                baseEmailBody = baseEmailBody.split('https://zoom.us/j/')[1]
            
            global neededEmailCode
            neededEmailCode = baseEmailBody.split("?", 1)[0]

            global neededEmailPassword
            baseEmailPassword = baseEmailBody.split("?pwd=", 1)[1]
            neededEmailPassword = baseEmailPassword.split('09')[0] + '09'
            if neededEmailPassword[0] + neededEmailPassword[1] == "3D":
                neededEmailPassword = neededEmailPassword.lstrip("3D")

            print(f"Meeting ID: {neededEmailCode}")
            print(f"PWD: {neededEmailPassword}\n")
        else:
            print("No Google Calendar email found")

        

    # Show exception
    except:
        PrintException()


# Function to remove unwanted characters from a string
def purify(string, separator):
    if separator in string:
        stringArray = string.split(separator)
        newString = ""
        for i in stringArray:
            newString = newString + i
        return newString
    else:
        return string


# Defines a notification function
def notify(title, text):
    os.system("""
              osascript -e 'display notification "{}" with title "{}"'
              """.format(text, title))


def checkTimes(time):
    for x in warningTimes:
        if time == x:
            read_email_from_gmail()
            notify(f"{neededEmailSubject} @ {classTime}", f"Meeting starts in {warningOffset} minutes.")
    i = 0
    for x in joinTimes:
        if time == x:
            read_email_from_gmail()
            joinMeeting("zoom", "zoom.us", classTimes[i], neededEmailCode, neededEmailPassword)
        i += 1


# Function to check if an app is already open
def checkAppOpen(appName):
    appOpen = False
    openApps = NSWorkspace.sharedWorkspace().runningApplications()
    for i in range(len(openApps)):
        returnStr = str(openApps[i])
        if appName.lower() in returnStr:
            appOpen = True
    return appOpen


# Function to launch app if closed, and then to navigate to app
def findApp(appName, appId):
    if checkAppOpen(appName) == False:
        # Navigate to applications directory where Zoom should be installed
        os.chdir("/Applications") 
        # Open zoom using command line
        subprocess.call(["open", "%s.app" % appId])
        # Notify user zoom was manually opened
        print("%s opened with script" % appName.capitalize())
        # Allow time for app launch
        time.sleep(7)
    # Navigate to zoom app through command line
    subprocess.call(["open", "-a", f"{appId}"])


# Function to join a meeting within zoom, using zoom shortcuts by default.
def joinMeetingInput(meetId, pwd):
    currentOS = platform.system()
    if currentOS == 'Darwin':
        commandKey = 'command'
    else:
        commandKey = "ctrl"
    # Exit any other menu open in zoom
    press('esc')
    time.sleep(0.5)
    # Default shortcut to join meeting
    hotkey(commandKey, 'j')
    # Allow for menu open
    time.sleep(0.5)
    # Write in the meeting id and enter
    typewrite(meetId)
    press('enter')
    # Allow for processing time
    time.sleep(2.5)
    # Write in password and enter
    pyperclip.copy(pwd)
    typewrite(pwd)
    press('enter')


# Function to populate warningTimes and joinTimes arrays, and to correct classTimes if needed.
def editTimes(warningOffset, joinOffset):
    try:
        # For every time in classTimes, generate corresponding times for the warning and joining times based on user offset preferences.
        for i in range(len(classTimes)):
            # If the 0 is missing in front, place it in.
            if len(classTimes[i]) == 4:
                classTimes[i] = f"0{classTimes[i]}"
            # Warningtime assignment and population
            if warningOffset != 0:
                warningTime = offsetTime(classTimes[i], warningOffset)
                warningTimes.append(warningTime)
            # Jointime assignment and population
            joinTime = offsetTime(classTimes[i], joinOffset)
            joinTimes.append(joinTime)
    except:
        print("You have incorrectly inputted times. Re-run program with correct times.")


# Function used to get custom times using offsets from user.
def offsetTime(time, offset):
    # Convert given hours and minutes into integers
    currentHours = int(time[0]+time[1])
    currentMins = int(time[3]+time[4])
    # If the offset will result in a change in hour
    if offset > currentMins:
        # Find new hour by subtracting the offset, divided by 60 to convert it into an hour offset, rounded down, from the original hour.
        newHours = currentHours - math.floor(offset / 60)
        # Find the new minutes by subtracting the remainder of that division from the current minutes.
        newMins = currentMins - (offset % 60)
    else:
        # No need for hour change
        newHours = currentHours
        # Find the new minutes by subtracting offset from currentMins
        newMins = currentMins - offset
    # Get final new time, format into string, and then return.
    newTime = f"{str(newHours)}:{str(newMins)}"
    return newTime


# Function to manage all zoom-related instructions.
def joinMeeting(appName, appId, classTime, meetId, pwd):
    # Notify user that meeting is aboutt to be joined, with sound prompt.
    notify(f"{neededEmailSubject} @ {classTime}", "Joining meeting. Please remove hands from device.")
    time.sleep(0.5)
    playsound('/System/Library/Sounds/ping.aiff')
    time.sleep(3)
    # Store the current app
    activeAppName = NSWorkspace.sharedWorkspace().activeApplication()['NSApplicationName']
    # Navigate to calling application
    findApp(appName, appId)
    time.sleep(0.5)
    # Join meeting with keystrokes
    joinMeetingInput(meetId, pwd)
    # Attempt to navigate back to initial application being used.
    try:
        subprocess.call(["open", "-a", f"{activeAppName}"])
    except:
        PrintException()


main()
#joinMeeting("zoom", "zoom.us", neededEmailCode, "dconline")