# Skrivet av Mikael Padwick för Åva Båtsällskap utan någon som helst support eller garanti av funktion.
# Ver 1.2.4

import os
import sys
sys.path.append(os.getcwd()+"/.env/lib/python3.11/site-packages")
from imbox import Imbox # pip install imbox
import traceback
from datetime import date, datetime
import subprocess
import time
from smtplib import SMTP_SSL, SMTP_SSL_PORT
#from pathlib import Path

host = "mailcluster.loopia.se"
port="993"
username = 'user@domain.com'
password = 'password'
printer = 'Garage-Skrivare'
Subject = '#print'
suffixes = ['pdf','xlsx']

runpath = '/opt/autoprint'
download_folder = runpath+'/download'
log_folder = runpath+'/log'

#-------------------------------------------------------------------------------------
def my_sendmail(sendto):
    try:
        # Craft the email by hand
        from_email = f"Skrivaren i garaget <{username}>"  # or simply the email address
        to_emails = [sendto]
        body = f"Ditt dokument kommer nu att skrivas ut på skrivaren i garaget på Åva Båtsällskap"
        headers = f"From: {from_email}\r\n"
        headers += f"To: {', '.join(to_emails)}\r\n"
        headers += f"Subject: Du har skicka ett dokument för utskrift\r\n"
        email_message = headers + "\r\n" + body # Blank line needed between headers and body
        # Connect, authenticate, and send mail
        smtp_server = SMTP_SSL(host, port=SMTP_SSL_PORT)
        smtp_server.set_debuglevel(1)  # Show SMTP server interactions
        smtp_server.login(username, password)
        smtp_server.sendmail(from_email, to_emails, email_message.encode('utf-8'))
        # Disconnect
        smtp_server.quit()
        my_log(f"Sending acknowledgement email to {sendto}")
    except:
        print(traceback.print_exc())
        my_log("Failed sending acknowledgement email.")

def my_log(text):
    log_file = f"{log_folder}/{date.today()}.log"
    log = open(log_file, "a")
    log.write(datetime.now().strftime("%H:%M:%S") + " " + text + "\n")
    log.close()
    
def my_download_file(attachment, filepath):
    try:
        with open(filepath, "wb") as fp:
            fp.write(attachment.get('content').read())               
        my_log("Saved file to: " + filepath)
        print(filepath)
        my_print_file(filepath)
    except:
            print(traceback.print_exc())    

def my_print_file(filepath):  
    global sendmail
    if filepath.endswith('pdf'):
        args = f'lp -d {printer} '  
        my_log("Sending file to printer: " + filepath)
        try:
            sendprint = args + filepath
            subprocess.call(sendprint, shell=True)
            my_log("Printer file now: " + filepath)
            sendmail = 1
        except:
            print(traceback.print_exc())
            my_log("Failed to print file: " + filepath)

    if filepath.endswith('xlsx'):
        subprocess.call(["libreoffice", "--headless", "--invisible",
            "--convert-to pdf",
            "--outdir", download_folder,
            os.path.basename(filepath)
        ])
        filepath = str(Path(filepath).with_suffix('.pdf'))

        args = f'lp -d {printer} '  
        my_log("Sending file to printer: " + filepath)
        try:
           sendprint = args + filepath
           subprocess.call(sendprint, shell=True)
           my_log("Printer file now: " + filepath)
           sendmail = 1
        except:
           print(traceback.print_exc())
           my_log("Failed to print file: " + filepath)

if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)

if not os.path.isdir(log_folder):
    os.makedirs(log_folder, exist_ok=True)
    
mail = Imbox(host, port=port, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages(unread=True,subject=f'{Subject}') # Unread messages
sendmail = 0
for (uid, message) in messages:
    mail.mark_seen(uid) # optional, mark message as read   
    for idx, attachment in enumerate(message.attachments):
        att_fn = attachment.get('filename').lower()
        try:
            if att_fn.endswith(tuple(suffixes)):
                download_path = f"{download_folder}/{date.today()}_{att_fn}"
                print(download_path)
                my_download_file(attachment, download_path.replace(" ", "-"))                     
        except:
            print(traceback.print_exc())
    if sendmail == 1:
        my_sendmail(message.sent_from[0]['email'])
mail.logout()
