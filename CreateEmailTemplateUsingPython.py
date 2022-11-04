from pathlib import Path
import win32com.client as win32
import pandas as pd
from datetime import datetime
import shutil

# Make date in appropriate format
now = datetime.now()
month = now.strftime("%m")
day = now.strftime("%d")
year = now.strftime("%Y")
date = str(month) + "-" + str(day) + "-" + str(year)

# Create zip files
shutil.make_archive('OutputFileHere', 'zip', 'FilesLiveHere')

# Load email distribution list
covidEmailList = pd.read_excel('PathOfEmailList')

# Create Email
outlook = win32.Dispatch("outlook.application")
for index, row in covidEmailList.iterrows():
    mail = outlook.CreateItem(0)
    mail.To = row["Email"]
    mail.CC = row["CC"]
    mail.Subject = f"TextHere"
    mail.HTMLBody = f"""
                    Good afternoon team,<br><br>
                    Text.<br><br>
                    Please let me know if you have any questions or concerns.<br><br>
                    Thank you,<br><br>"""
    attachmentPath1 = 'PathToAttachment'
    attachmentPath2 = 'PathToAttachment'
    mail.Attachments.Add(Source=attachmentPath1)
    mail.Attachments.Add(Source=attachmentPath2)

    mail.Display()
    #mail.Send()