# Skrivet av Mikael Padwick för Åva Båtsällskap utan någon som helst support eller garanti av funktion.
# Ver 1.2.1

import os
import sys
sys.path.append(os.getcwd()+'\Lib\site-packages')
from imbox import Imbox # pip install imbox
import traceback
import ghostscript
from datetime import date, datetime
import subprocess
import time
from smtplib import SMTP_SSL, SMTP_SSL_PORT

host = "mailcluster.loopia.se"
port="993"
username = 'mail@domain.com'
password = 'password'
printer = 'HP Color LaserJet CP1215'
Subject = '#print'
suffixes = ['pdf']
GSPath = 'C:\Program Files\gs\gs10.01.1'

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
    log_file = f"{log_folder}\{date.today()}.log"
    log = open(log_file, "a")
    log.write(datetime.now().strftime("%H:%M:%S") + " " + text + "\n")
    log.close()
    
def my_print_file(attachment, filepath):
    GS = GSPath + '\\bin\gswin64c.exe'
    if att_fn.endswith(tuple(suffixes)):
        with open(filepath, "wb") as fp:
            fp.write(attachment.get('content').read())
               
        args = f'\"{GS}\" ' \
               '-sDEVICE=mswinpr2 ' \
               '-dBATCH ' \
               '-dNOPAUSE ' \
               f'-sOutputFile#"%printer%{printer}" '
                
        my_log("Sending to printer: " + filepath)
        ghostscript = args + '\"' + filepath.replace('\\', '\\\\') + '\"'
        subprocess.call(ghostscript, shell=True)
        my_log("Printer file now: " + filepath)

download_folder = os.getcwd()+"\Download"
log_folder = os.getcwd()+"\Log"

if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)

if not os.path.isdir(log_folder):
    os.makedirs(log_folder, exist_ok=True)
    
mail = Imbox(host, port=port, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages(unread=True,subject=f'{Subject}') # Unread messages

for (uid, message) in messages:
    mail.mark_seen(uid) # optional, mark message as read   
    for idx, attachment in enumerate(message.attachments):
        print(attachment.get('filename'))
        try:
            att_fn = attachment.get('filename')
            #att_fn.replace(" ", "_")
            print(att_fn)
            download_path = f"{download_folder}\{date.today()}_{att_fn}"
            my_print_file(attachment, download_path)                
        except:
            print(traceback.print_exc())    
    my_sendmail(message.sent_from[0]['email'])
    
mail.logout()