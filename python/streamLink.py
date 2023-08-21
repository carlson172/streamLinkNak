#!/usr/bin/env python
# coding: utf-8

# streamlink v 0.6
#
# streamlink is to set links for redirection to youtube based on a Google spreadsheet
# mail with link wil sent out when new links can be set. Mail receiver and content will be also stored in the spreadsheet
# 
# # Google sheet link update
# 
#   import libs

import pandas as pd
import json
import csv
import pickle
from google.oauth2 import service_account
from gspread_formatting import *
import pygsheets
import sys
import os
import subprocess
import getpass
from pytube import extract
import configparser
from datetime import date, datetime
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type

### Get Config Parameters from config.ini

username = getpass.getuser()
config = configparser.ConfigParser()
userDir = ("/home/" + username + "/")
config.read( userDir + "streamLinkNak/config.ini")
startcol = config.getint("streamLink", "startcol")
endcol = config.getint("streamLink", "endcol")
user = config.get("streamLink", "user")
host = config.get("streamLink", "host")
url = config.get("streamLink", "url")

#if username == "bash":
#    username = getpass.getuser()
goSheetCred = ( userDir + config.get("streamLink", "goSheetCred"))
goMailCred = ( userDir + config.get("streamLink", "goMailCred") )

#get current date (time set to 00:00:00)
today = datetime.today().replace(microsecond=0, second=0, minute=0, hour=0)
print("Today: ", today)


# connect to google
with open(goSheetCred) as source:
   info = json.load(source)
credentials = service_account.Credentials.from_service_account_info(info)
client = pygsheets.authorize(service_account_file=goSheetCred)

# access spreadsheet
sheet = client.open_by_url(url)
#print(sheet)

# open worksheets
wks = sheet.worksheet_by_title('links')
mailsheet = sheet.worksheet_by_title('Mail')
emailReceiver = sheet.worksheet_by_title('EmailRec')
emailContentSheet = sheet.worksheet_by_title('AutoMail')
#print(wks)
#print(mailsheet)

#get email receivers from google sheet
maxRows = wks.rows # To view the number of rows
emailReceiverRows = emailReceiver.rows

emailRow = emailReceiver.get_col(2)
emailContentRows = emailContentSheet.rows

emailReceiverString = ""
for emailCurrRow in range(1, emailReceiverRows):
    if (len(emailRow[emailCurrRow]) != 0):
        emailReceiverString = ( emailRow[emailCurrRow] + ", " + emailReceiverString)

#get email Fail receivers from google sheet
emailRow = emailReceiver.get_col(3)
emailContentRows = emailContentSheet.rows

emailFailReceiverString = ""
for emailCurrRow in range(1, emailReceiverRows):
    if (len(emailRow[emailCurrRow]) != 0):
        emailFailReceiverString = ( emailRow[emailCurrRow] + ", " + emailFailReceiverString)

     
#get shell commands for each "slot"
churchCommunity = wks.get_row(1)
linkcode = wks.get_row(2) 
channelLink = wks.get_row(3)
keyRow = wks.get_row(4)

#here are all the dates stored
first_column = wks.get_col(1)

#variable presets for the loop
search = True
#first date afer header range
rowNumber=6

#now walk over row, pick the first column and compare the date with the selected date
while search:
    rowNumber +=1
    if rowNumber >= maxRows:
        search = False;
    else:        
        searchDate = datetime.strptime(first_column[rowNumber], '%d.%m.%Y')
        #print("searchDate: ", searchDate)
        search = searchDate < today
        
if rowNumber < maxRows:
    print("selected Date: " + first_column[rowNumber])
else:
    sys.exit('Date is not found')

#check if today was a transmission. If so no need to set the links today
if searchDate == today:
    #today is nothing todo
    print("Today was a transmission, wait till tomorrow to set new link")
    exit()

#create mail content from google sheet and build in date of next transmission
emailContentCol = emailContentSheet.get_col(1)

#Subject stored in the first 3 lines
emailSubjectString = ""
for emailContentCurrRow in range(0, 3):
    if (len(emailContentCol[emailContentCurrRow]) != 0):
        match str(emailContentCol[emailContentCurrRow]):
            case "DATUM":
                #print("Datum")
                emailSubjectString = ( emailSubjectString + first_column[rowNumber])
            case "NL":
                emailSubjectString = ( emailSubjectString + "\n")
            case _:
                emailSubjectString = ( emailSubjectString + emailContentCol[emailContentCurrRow] )

#mail content storded in the following lines
emailContentString = ""
for emailContentCurrRow in range(3, emailContentRows):
    if (len(emailContentCol[emailContentCurrRow]) != 0):
        match str(emailContentCol[emailContentCurrRow]):
            case "DATUM":
                emailContentString = ( emailContentString + first_column[rowNumber])
            case "NL":
                emailContentString = ( emailContentString + "\n")
            case _:
                emailContentString = ( emailContentString + emailContentCol[emailContentCurrRow] )


#print( "emailContentString: " + emailContentString)
#senderTestList = ( "user@host")
#send_message(service, senderTestList, emailSubjectString, emailContentString, [ ])
#exit()

#for some reasons the rowNumber is one higher than in the first_column function....
link_sel_link = wks.get_row(rowNumber+1)
mailToSend = False
failForMail = ""

#go over all (4) columns and pick up the entry in the selected row
for x in range(startcol, endcol):
    #if selected element is empty exit...
    if(len(link_sel_link[x]) == 0):
        sys.exit('field is 0. Cannot be used')
    print("Use " + searchDate.strftime('%d.%m.%Y') + " and try for " + churchCommunity[x], end="")
    id = extract.video_id(link_sel_link[x])
    #print( "videoID: " + id )
    
    storedId = "/usr/bin/ssh -i " + userDir + ".ssh/" + keyRow[x] + " " + user + "@" + host + " nactube_ctl list streams "
    try:
        ret = subprocess.check_output(storedId, shell=True).decode('utf-8').strip().split('\n')
    except subprocess.CalledProcessError as err:
        print(err)
        ret = "error!"

    #print(" return from query: " + ret[0] + "--") 
    #print( "videoID: " + id + "--")
    
    if (ret[0] == id):
        print(" already set")
    else:
    #now construct the shell command
        ausgang = "ssh -i " + userDir + ".ssh/" + keyRow[x] + " " + user + "@" + host + " nactube_ctl update path " + channelLink[x] + " set stream " + id
        #print(ausgang)
        
        #execute it
        ret = os.system(ausgang)
        #ret = 0

        # for future extensions..
        #all good?
        if x == 1:
            col = "B"
        if x == 2:
            col = "C"
        if x == 3:
            col = "D"
        if x == 4:
            col = "E"
        if x == 5:
            col = "F"
        dataRange = wks.range(col + "" + str(rowNumber) + ":" + col + "" + str(rowNumber))
        #fmt = cellFormat(
        #    backgroundColor=color(0, 1, 0),
        #    textFormat=textFormat(bold=False, foregroundColor=color(1, 1, 1)),
        #    horizontalAlignment='CENTER'
        #    )
        
        #now check the returncode of the os command
        if (ret == 0):
            # ok, then give feedback and mark the mail to be send
            print(" stream set to " + id)
            mailToSend = True
            #wks.batch_update(fmt)
            #dataRange.format_cell_range(wks, dataRange, fmt)
            #wks.cellFormat(dataRange, { "backgroundColor": {"red": 0, "green": 1, "blue": 0}})
            #dataRange.setBackground("green")
        else:
            # error, prepare the mail to the fail list to be send
            print("Fehler: " + ausgang)
            failForMail = churchCommunity[x]


#setup mail access -- start

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
SCOPES = ['https://mail.google.com/']
our_email = config.get("streamLink", "senderAddress")

def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError:
                print("Credentials could not be refreshed, possibly the authorization was revoked by the user.")
                os.unlink(goMailCred)
                return
        else:
            flow = InstalledAppFlow.from_client_secrets_file(goMailCred, SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

# get the Gmail API service
service = gmail_authenticate()

# Adds the attachment with the given filename to the given message
def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
        fp = open(filename, 'rb')
        msg = MIMEText(fp.read().decode(), _subtype=sub_type)
        fp.close()
    elif main_type == 'image':
        fp = open(filename, 'rb')
        msg = MIMEImage(fp.read(), _subtype=sub_type)
        fp.close()
    elif main_type == 'audio':
        fp = open(filename, 'rb')
        msg = MIMEAudio(fp.read(), _subtype=sub_type)
        fp.close()
    else:
        fp = open(filename, 'rb')
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(fp.read())
        fp.close()
    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

def build_message(destination, obj, body, attachments=[]):
    if not attachments: # no attachments given
        message = MIMEText(body)
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
    else:
        message = MIMEMultipart()
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message.attach(MIMEText(body))
        for filename in attachments:
            add_attachment(message, filename)
    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}

def send_message(service, destination, obj, body, attachments=[]):
    return service.users().messages().send(
      userId="me",
      body=build_message(destination, obj, body, attachments)
    ).execute()

#setup mail access -- end



#mailContent for error case
mailFailContent = ("Link für Gemeinde " + failForMail + " konnte nicht gesetzt werden!!!\n"
                   "Bitte überprüfen.\n\nVielen Dank\n")
#Subject for error case
mailFailSubject = ("streamLink Fehler für " + first_column[rowNumber])

if mailToSend:
    #Mail with the links go out
    print("Will send out email with links..... ")
    send_message(service, emailReceiverString, emailSubjectString, emailContentString, [ ])

if failForMail != "":
    #error mail goes out.
    print("Error: will send out error mail....")
    send_message(service, emailFailReceiverString, mailFailSubject, mailFailContent, []) 
                 
print("Thanks for using streamLink")


