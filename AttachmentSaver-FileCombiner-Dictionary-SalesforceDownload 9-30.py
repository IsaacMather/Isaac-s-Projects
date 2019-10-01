#this library is to set the directory
import os

#these libraries are to manipulate the excel files
import pandas as pd
import openpyxl

#this library is to interact with MS Outlook
import win32com.client as win32

#this library is to check todays date
import datetime

#this library is to manipulate salesforce
from salesforce_reporting import Connection, ReportParser


##The below code helps by downloading the attachments from outlook emails that are 'Unread' (and changes the mail to Read.)
##or from 'Today's' date.without altering the file name. Just pass the 'Subject' argument.
##https://stackoverflow.com/questions/39656433/how-to-download-outlook-attachment-from-python-script



#download the necessary report from salesforcefrom salesforce_reporting import Connection

sf = Connection(username='isaama2@vsp.com', password='c@lmBunny13', security_token='szDUtiFW0gVMmrU72vhGhyyj')
report = sf.get_report('Eloqua Opportunity ID Reference', includeDetails=True)
parser = salesforce_reporting.ReportParser(report)

parser.records()




##def saveattachments(subject):
##    path = os.path.expanduser(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments')
##    today = datetime.date.today()
##    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
##    inbox = outlook.GetDefaultFolder(6) 
##    messages = inbox.Items
##    #set the title of the email you want to look for here. can we use regex to make flexible email title searches?
##    subject = 'Working with Global Analytics and Insights Team Checklist (002)'
##    for message in messages:
##        if message.Subject == subject: #and message.Unread or message.Senton.date() == today:
##            #body_content = message.body
##            attachments = message.Attachments
##            attachment = attachments.Item(1)
##            for attachment in message.Attachments:
##                attachment.SaveAsFile(os.path.join(path, str(attachment)))
##                #if message.Subject == subject and message.Unread:
##                #    message.Unread = False
##                print('Attachment Saved!')
##                break
            
##saveattachments(subject)


##dataframe_directory = 'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments'
##eloqua_results = 'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test'
##POA_lists = 'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team\Old Files'
##combined_directory = 'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists'


#create a dictionary using the reference sheet. ferences the dictionary to to know which file sent from the POA team to the eloqua team to use as a resource to retrieve TaxID.
###Then saves the combined sheet
##os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments')
###get the read excel to just open the dataframe name
##df = pd.read_excel(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\POA_Eloqua_email_and_lists.xlsx')
##df = df.set_index('Eloqua File Name')['POA File Name'].to_dict()
###os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team')
###print(df)

##for key, val in df.items():
##print(key, "=>", val)


##clean up the directories by making them dynamic
##get the new file out of the attachment combiner function
##get the file into the file mailer function
##need to retrieve the opportunity ID from salesforce
##
##function to find a filename and its corresponding key, then combine the file and its key
##def attachment_combiner(df):
##    files = os.listdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test')
##    for Eloqua_file in files:
##        POA_file = df[Eloqua_file]
##        print(Eloqua_file)
##        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test')
##        sheets = pd.ExcelFile(Eloqua_file)
##        sheets = sheets.sheet_names
##        for Eloqua_sheet in sheets:
##            #print(Eloqua_sheets)    
##            main = pd.read_excel(Eloqua_file, sheet_name = Eloqua_sheet, index_col = None)
##            os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team\Old Files')
##            secondary = pd.read_excel(POA_file, sheet_name = Eloqua_sheet, index_col = None)
##            combined = pd.merge(main, secondary, sort=False, left_on=['Email Address'], right_on=['Contact: Primary Contact Email'], how = 'left')
##            os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists')
##            file_name = Eloqua_file + ' + ' + Eloqua_sheet + ' done by Python.xlsx'
##            combined.to_excel(file_name, index=False)
##            print('combination complete! file saved to folder!')
##
##attachment_combiner(df)
##
##def mail_new_file():
##    #send the file in an Outlook email
##    import win32com.client
##    outlook = win32.Dispatch('outlook.application')
##    mail = outlook.CreateItem(0)
##    mail.To = 'isaama2@vsp.com'
##    mail.Subject = 'Perfect Pair Rebate List'
##    mail.Body = 'Hi Alex, \r\n\nHere is the perfect pair rebate list!\r\n\nCheers, \r\nIsaac'
##    attachment = r'\\ntsca126\PRmisc\Provider Operations and Analysis\Reporting & Analytics\Marketing\Sales\POA-1364 Perfect Pair Rebate\POA-1164 Development\python perfect pair.xlsx'
##    #attachment = directory + \ + file_name
##    print('working correctly')
##    mail.Attachments.Add(attachment)
##    mail.Send()
##    #clarify it worked
##    print('Operation successful!')
##
##mail_new_file()
