# Skrivet av Mikael Padwick för Åva Båtsällskap utan någon som helst support eller garanti av funktion
# Ver 1.0

import os
import sys
sys.path.append(os.getcwd()+'\Lib\site-packages')
from imbox import Imbox # pip install imbox
import traceback
import ghostscript
from datetime import date, datetime
import subprocess
import time

host = "mailcluster.loopia.se"
port="993"
username = 'mail@domain.com'
password = 'password'
printer = 'HP Color LaserJet CP1215'
Subject = '#print'
suffixes = ['pdf']
GSPath = 'C:\Program Files\gs\gs10.01.1'

#-------------------------------------------------------------------------------------

download_folder = os.getcwd()+"\Download"
log_folder = os.getcwd()+"\Log"

if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)

if not os.path.isdir(log_folder):
    os.makedirs(log_folder, exist_ok=True)
    
mail = Imbox(host, port=port, username=username, password=password, ssl=True, ssl_context=None, starttls=False)
messages = mail.messages(unread=True,subject=f'{Subject}') # Unread messages

GSPath = GSPath + '\\bin\gswin64c.exe'
for (uid, message) in messages:
    mail.mark_seen(uid) # optional, mark message as read
    for idx, attachment in enumerate(message.attachments):
        print(attachment.get('filename'))
        try:
            att_fn = attachment.get('filename')
            #att_fn.replace(" ", "_")
            print(att_fn)
            download_path = f"{download_folder}\{date.today()}_{att_fn}"
            if att_fn.endswith(tuple(suffixes)):
                with open(download_path, "wb") as fp:
                    fp.write(attachment.get('content').read())
               
                args = f'\"{GSPath}\" ' \
                       '-sDEVICE=mswinpr2 ' \
                       '-dBATCH ' \
                       '-dNOPAUSE ' \
                       f'-sOutputFile#"%printer%{printer}" '

                log_file = f"{log_folder}\{date.today()}.log"
                log = open(log_file, "a")
                log.write(datetime.now().strftime("%H:%M:%S") + " Sending to printer: " + download_path + "\n")
                ghostscript = args + '\"' + download_path.replace('\\', '\\\\') + '\"'
                subprocess.call(ghostscript, shell=True)
                log.write(datetime.now().strftime("%H:%M:%S") + " Printer file now: " + download_path + "\n")
                log.close()

                
        except:
            print(traceback.print_exc())
    
mail.logout()