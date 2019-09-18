#this library is to set the directory
import os

#these libraries are to manipulate the excel files
import pandas as pd
import openpyxl

#this library is to interact with MS Outlook
import win32com.client as win32

#this library is to check todays date
import datetime

#The below code helps by downloading the attachments from outlook emails that are 'Unread' (and changes the mail to Read.)
#or from 'Today's' date.without altering the file name. Just pass the 'Subject' argument.
#https://stackoverflow.com/questions/39656433/how-to-download-outlook-attachment-from-python-script

def saveattachments(subject):
    path = os.path.expanduser(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments')
    today = datetime.date.today()
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) 
    messages = inbox.Items
    #set the title of the email you want to look for here. can we use regex to make flexible email title searches?
    subject = 'Working with Global Analytics and Insights Team Checklist (002)'
    for message in messages:
        if message.Subject == subject: #and message.Unread or message.Senton.date() == today:
            #body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                #if message.Subject == subject and message.Unread:
                #    message.Unread = False
                print('Attachment Saved!')
                break
            
saveattachments(subject)


#create a dictionary using the reference sheet
os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments')
df = pd.read_excel(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\POA_Eloqua_email_and_lists.xlsx')
dictionary = df.to_dict()
#os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team')

#function to find a filename and its corresponding key, then combine the file and its key
######need to get the merge key set up for all spreadsheets
######probably run a test run on one file in a new file, will need to reset the directories
def attachment_combiner(dictionary):
    files = os.listdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test')
    for Eloqua_file in files():
        POA_file = dictionary[Eloqua_file]
        main = pd.read_excel(Eloqua_file,index_col = None)
        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team')
        secondary = pd.read_excel(POA_file,index_col = None)
        combined = pd.merge(main, secondary, sort=False, on='Email Address', how = 'left')
        ##need to set a new directory for the program to save the new files to
        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists')
        #combined = combined.drop_duplicates(subset=['Tax ID','Opportunity ID'],keep='first', inplace=False)
        file_name = 'test results'
        combined.to_excel(file_name, index=False)

attachment_combiner(dictionary)

def mail_new_file():
    #send the file in an Outlook email
    import win32com.client
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'isaama2@vsp.com'
    mail.Subject = 'Perfect Pair Rebate List'
    mail.Body = 'Hi Alex, \r\n\nHere is the perfect pair rebate list!\r\n\nCheers, \r\nIsaac'
    attachment = r'\\ntsca126\PRmisc\Provider Operations and Analysis\Reporting & Analytics\Marketing\Sales\POA-1364 Perfect Pair Rebate\POA-1164 Development\python perfect pair.xlsx'
    #attachment = directory + \ + file_name
    print('working correctly')
    mail.Attachments.Add(attachment)
    mail.Send()
    #clarify it worked
    print('Operation successful!')

mail_new_file()
