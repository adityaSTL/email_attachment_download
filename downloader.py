import xlrd
import pandas as pd
import numpy as np
from datetime import date,timedelta
import datetime as dt
import openpyxl
import matplotlib.pyplot as plt
from openpyxl_image_loader import SheetImageLoader
import excel2img
import imaplib, email
import os,shutil
from imbox import Imbox
import traceback
from datetime import date,datetime,timedelta
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

global log
log=""
def check(name):
    global log
    pkg_b=['PKG-B','PKGB','PKG B','Package-B','Package B','PackageB','PkgB','Pkg B','Pkg-B','PACKAGEB','PACKAGE B','PACKAGE-B']
    pkg_c=['PKG-C','PKGC','PKG C','Package-C','Package C','PackageC','PkgC','Pkg C','Pkg-C','PACKAGEC','PACKAGE C','PACKAGE-C']
    pop=['POP','pop']
    at_tracker=['AT','AT Tracker','ATC Tracker','ATC','AT tracker','at tracker']
    dpr=['Daily Progress Reports','dpr','DPRs','DPR','Progress Reports']
    #print("Before calling",name)
    string=""
    count=0
    if any(x in name for x in pkg_b):
        string+="Pkg-B "
        count+=1
    elif any(x in name for x in pkg_c):
        string+="Pkg-C "
        count+=1
        
    if any(x in name for x in at_tracker):
        string+="ATC "
        count+=1
    elif any(x in name for x in pop):
        string+="POP " 
        count+=1
    elif any(x in name for x in dpr):
        string+="DPR "  
        count+=1
    
    
    string+=str(date.today()-timedelta(days=1))
    
    string+=name[-5:]
    if count==2:
        log+="/nFound: "+string
        return (string,count)
    else:
        #string="False"
        return ("False",count)
def folder_cleaner():
    download_folder = r"D:\Reports and Trackers"
    for xl in os.listdir(download_folder):
        name,count=check(xl)
        print(xl,name)
        if name=="False":
            os.remove(os.path.join(download_folder,xl))
        else:
            os.rename(os.path.join(download_folder,xl),os.path.join(download_folder,name))
        print("----------------------------------------------------------------")
def get_date(date):
    date_patterns = ["%a, %d %b %Y %H:%M:%S %z","%a, %d %b %Y %H:%M:%S %Z","%a, %d %b %Y %H:%M:%S %z (%Z)"]
    for pattern in date_patterns:
        try:
            return datetime.strptime(date, pattern).date()
        except:
            pass
# def send_update():
#     port,smtp_server,host,username,paswd=get_details()
#     sender_email = "xyz@gmail.com"
#     receiver_email ="abc@gmail.com"
#     password = "adsfaegrh"
#     message = MIMEMultipart("alternative")
#     date=datetime.today()
#     message["Subject"] = "Update on attachment downloader "+str(date)
#     message["From"] = sender_email
#     message["To"] = receiver_email
#     text = """\
#     Here is your update on attachment downloaded and possible errors for today():
#     """
#     html = """\
#     <html>
#     <body>
#         <p>Heyo A,<br>
#         Developed by Adi<br>
#         <a href="https://github.com/adityaSTL">github repo</a> 
#         Thank you!
#         </p>
        
#         """+"Here is your update on attachment downloaded and possible errors for today():"+"\nScript ran @:"""+str(datetime.now())+"\nRun by:"+str(os.getlogin())+"\nLogs: "+log+"""
#         <h2>
#         Test
#         </h2>
#     </body>
#     </html>
#     """
#     part1 = MIMEText(text, "plain")
#     part2 = MIMEText(html, "html")
#     message.attach(part1)
#     message.attach(part2)
#     context = ssl.create_default_context()
#     with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
#         server.login(sender_email, password)
#         server.sendmail(
#             sender_email, receiver_email, message.as_string()
#         )
def get_details():
    port=465
    smtp_server = "smtp.gmail.com"
    host = "imap.gmail.com"
    username = 'sender@email.address'
    paswd = 'adsfadsf'
    return (port,smtp_server,host,username,paswd)
def get_attachment():
    port,smtp_server,host,username,paswd=get_details()
    download_folder = r"D:\Reports and Trackers"
    if os.path.isdir(download_folder):
        shutil.rmtree(download_folder)
    os.makedirs(download_folder, exist_ok=True)
    mail = Imbox(host, username=username, password=paswd, ssl=True, ssl_context=None, starttls=False)
    messages = mail.messages() # defaults to inbox
    count=0
    for (uid, message) in reversed(messages):
        count+=1
        if count<15:  
            print(message.subject,message.date)
            datetime_object=get_date(message.date)

            if (datetime_object==date.today()):
                for idx, attachment in enumerate(message.attachments):
                    try:
                        att_fn = attachment.get('filename')    
                        if(att_fn[-4:]=='xlsx' or att_fn[-4:]=='xlsb'):
                            download_path = f"{download_folder}/{att_fn}"
                            print(download_path)
                            with open(download_path, "wb") as fp:
                                fp.write(attachment.get('content').read())
                    except:
                        print(traceback.print_exc())
        else:
            break

get_attachment()
folder_cleaner()
#generate_images()

#Can ignore send_update function its just additional thing
#send_update()        
