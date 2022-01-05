# modified from https://www.thepythoncode.com/article/use-gmail-api-in-python

import os
#import pickle
import base64
from apiclient import errors
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import google.auth.exceptions
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from mimetypes import guess_type as guess_mime_type

#todo: make class so that importing this module won't automatically run oauth

# Request all access (permission to read/send/receive emails, manage the inbox, and more)
# https://developers.google.com/gmail/api/auth/scopes
SCOPES = ['https://mail.google.com/'] #TODO: reduce to a less invasive scope
import yaml

with open("config.yaml", 'r') as stream:
    try:
        settings = yaml.safe_load(stream)
        our_email = settings['from_email']
    except:
        print('Issue with config.yaml')
        raise

def gmail_authenticate(): #TODO: move this to its own file, so other modules only have to import quickstart and not the entire email_sender
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        #with open("token.json", "rb") as token:
            #creds = pickle.load(token) i don't think pickling is necessary here

    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except google.auth.exceptions.RefreshError:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        #with open("token.json", "wb") as token:
            #pickle.dump(creds, token)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
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
    if not attachments:  # no attachments given
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


def send_message(service, user_id, message):
    try:
        message = service.users().messages().send(userId=user_id, body=message).execute()
        #print('Message Id: %s' % message['id'])
        return message
    except Exception as e:
        print('An error occurred: %s' % e)
        return None


def send_message_easy(service, destination, obj, body, attachments=[]):
    try:
        return service.users().messages().send(
            userId="me",
            body=build_message(destination, obj, body, attachments)
            ).execute()
    except:
        print('Something went wrong when sending the email')
        raise
        return None


def reply_message(service, threadId, destination, message_id,
                  prev_references, obj, body, attachments=[]):
    # https://stackoverflow.com/questions/31626069/mime-headers-not-making-it-through-gmail-api
    if not attachments:  # no attachments given
        message = MIMEText(body)
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message['In-Reply-To'] = message_id
        message['References'] = f'{prev_references} {message_id}'
        raw = base64.urlsafe_b64encode(message.as_string().encode()).decode()
        # message['Content-Type'] = 'multipart/alternative; boundary="000000000000e4b26205c6b3f2e8"'
    else:
        message = MIMEMultipart()
        message['to'] = destination
        message['from'] = our_email
        message['subject'] = obj
        message['In-Reply-To'] = message_id
        message['References'] = f'{prev_references} {message_id}'
        message.attach(MIMEText(body))
        for filename in attachments:
            add_attachment(message, filename)
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    # raw = build_message_raw(destination, obj, body, attachments)
    request_body = {'raw': raw, 'threadId': threadId}
    send_message(service, 'me', request_body)
    # draft = service.users().drafts().create(userId="me", body=message).execute()

# send_message_easy(service, target_email, "subject test", 'body test')