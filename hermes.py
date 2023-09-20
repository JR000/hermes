from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import *
from tkinter.messagebox import *
import re
import os
import mimetypes

import base64
from email.message import EmailMessage

from email.utils import formataddr


from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from apiclient import errors, discovery
import mimetypes
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase

SCOPES = ['https://www.googleapis.com/auth/gmail.send']



def send_message(creds=None, to="", fromName="", fromEmail="", subject="", msgHtml="", attachments=[]):
    try:
        service = build('gmail', 'v1', credentials=creds)
        message = MIMEMultipart('alternative')

        message['To'] = to
        message['From'] = formataddr((fromName, fromEmail))
        message['Subject'] = subject

        message.attach(MIMEText(msgHtml, 'html'))   


        for attachmentFile in attachments:
            content_type, encoding = mimetypes.guess_type(attachmentFile)

            if content_type is None or encoding is not None:
                content_type = 'application/octet-stream'
            main_type, sub_type = content_type.split('/', 1)
            with open(attachmentFile, 'rb') as fp:
                msg = MIMEBase(main_type, sub_type)
                msg.set_payload(fp.read())
            filename = os.path.basename(attachmentFile)
            msg.add_header('Content-Disposition', 'attachment', filename=filename)
            message.attach(msg)


        encoded_message = base64.urlsafe_b64encode(message.as_bytes()) \
            .decode()

        create_message = {
            'raw': encoded_message
        }
        send_message = (service.users().messages().send
                        (userId="me", body=create_message).execute())
        print(F'Message Id: {send_message["id"]}')
    except HttpError as error:
        print(F'An error occurred: {error}')
        send_message = None
    return send_message

def send_messages(emails=[], fromName="", fromEmail="", subject="", msgHtml="", attachments=[]):
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    for email in emails:
        send_message(creds, to=email, fromName=fromName, fromEmail=fromEmail, subject=subject, msgHtml=msgHtml, attachments=attachments)

    



window = Tk()
window.title("Hermes [OKDF-FF]")
window.geometry('350x370')

Label(window, text="Sender's name:").grid(column=0, row=0, sticky=W)
nameTxt = Entry(window,width=30)
nameTxt.grid(column=1, row=0, sticky=W)

Label(window, text="Sender's email:").grid(column=0, row=1, sticky=W)
emailTxt = Entry(window,width=30)
emailTxt.grid(column=1, row=1, sticky=W)

Label(window, text="Subject:").grid(column=0, row=2, sticky=W)
subjectTxt = Entry(window,width=30)
subjectTxt.grid(column=1, row=2, sticky=W)


Label(window, text="Emails:").grid(column=0, row=3, sticky=NW)
emails = Text(window, width=30, height=5)
emails.grid(column=1, row=3, sticky=W)
def load_emails():
    f = askopenfile()
    emails.delete('1.0', END)
    
    for email in f.readlines():
        emails.insert('1.0', email + '\n')  
Button(window, text="Load list", command=load_emails).grid(column=1, row=4, sticky=W)

Label(window, text="Attachments:").grid(column=0, row=5, sticky=NW)

attachments = set()
attachmentsTxt = Text(window, width=30, height=5)
attachmentsTxt.grid(column=1, row=5, sticky=W)

def add_attachment():
    f = askopenfilenames()
    for fname in f:
        if fname in attachments:
            continue
        attachments.add(fname)
        attachmentsTxt.insert(END, fname + '\n')

Button(window, text="Add file", command=add_attachment).grid(column=1, row=6, sticky=W)

Label(window, text="HTML-content:").grid(column=0, row=7, sticky=NW)
htmlFileStr = StringVar()
htmlFileStr.set('<None>')
Label(window, textvariable=htmlFileStr).grid(column=1, row=7, sticky=W)


def load_html():
    f = askopenfilenames()
    htmlFileStr.set(f[0])
Button(window, text="Load file", command=load_html).grid(column=1, row=8, sticky=W)

def process_emails():
    EMAIL_REGEX = re.compile(r"[^@]+@[^@]+\.[^@]+")
    result = []
    for email in emails.get('1.0', END).splitlines():
        if EMAIL_REGEX.match(email):
            result.append(email)
    emails.delete('1.0', END)
    emails.insert('1.0', '\n'.join(result))
    return result

def process_attachments():
    result = []
    for attachment in attachmentsTxt.get('1.0', END).splitlines():
        if os.path.isfile(attachment):
            result.append(attachment)
    attachmentsTxt.delete('1.0', END)
    attachmentsTxt.insert('1.0', '\n'.join(result))
    return result
def send_email():
    emails = process_emails()
    attachments = process_attachments()    

    if not os.path.isfile(htmlFileStr.get()):
        showerror("Error", "Can't find the content file")
        htmlFileStr.set('<None>')
        return

    msgHtml = ""
    with open(htmlFileStr.get(), encoding='utf-8') as file:
        msgHtml = file.read()
    send_messages(emails=emails, fromName=nameTxt.get(), fromEmail=emailTxt.get(), subject=subjectTxt.get(), msgHtml=msgHtml, attachments=attachments)

Button(text="Send emails", command=send_email).grid(column=0, row=9, columnspan=3)




window.mainloop()
