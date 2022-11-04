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

# Create zip files for COVID Reports
shutil.make_archive('C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Beyond_Lapse_Rules_Report', 'zip', 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/Market_Beyond_Lapse_Rules_Reports')
shutil.make_archive('C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Cancellation_Report', 'zip', 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/Market_Cancellation_Reports')
shutil.make_archive('C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Lapse_Report', 'zip', 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/Market_Lapse_Reports')

# Load email distribution list
covidEmailList = pd.read_excel('C:/Users/e20789/Desktop/Python Scripts/Files/Email_List_COVID_Should_Be_Lapsed.xlsx')

# Create Email
outlook = win32.Dispatch("outlook.application")
for index, row in covidEmailList.iterrows():
    mail = outlook.CreateItem(0)
    mail.To = row["Email"]
    mail.CC = row["CC"]
    mail.Subject = f"Account Cancellation & At-Risk Visibility Reporting | " + date
    mail.HTMLBody = f"""
                    Good afternoon team,<br><br>
                    Attached are the three ZIP folders that contain the “Account Cancellations”, “Beyond Lapse Rules”, and “Account Lapses” raw data.<br><br>
                    A brief description of how an account appears on each report is below:<br>
                    •&nbsp;&nbsp;&nbsp;&nbsp;Cancellations = The account requested to cancel their coverage since the start of 2022.<br>
                    •&nbsp;&nbsp;&nbsp;&nbsp;Beyond Lapse Rules = The account’s current payment history has placed them beyond standard lapse rules (CPR and Non-CPR logic).<br>
                    •&nbsp;&nbsp;&nbsp;&nbsp;Lapses = The account lapsed since the start of 2022 and remains lapsed as of today’s date<br><br>
                    Please let me know if you have any questions or concerns.<br><br>
                    Thank you,<br><br>"""
    attachmentPath1 = 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Beyond_Lapse_Rules_Report.zip'
    attachmentPath2 = 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Cancellation_Report.zip'
    attachmentPath3 = 'C:/Users/e20789/Desktop/COVID19_Should_Be_Lapsed/'+date + '_Lapse_Report.zip'
    mail.Attachments.Add(Source=attachmentPath1)
    mail.Attachments.Add(Source=attachmentPath2)
    mail.Attachments.Add(Source=attachmentPath3)

    mail.Display()
    #mail.Send()

