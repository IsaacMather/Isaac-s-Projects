
#this library is to set the directory
import os

#these libraries are to manipulate the excel files
import pandas as pd
import openpyxl

#this library is to interact with MS Outlook
import win32com.client as win32

#this library is to check todays date
import datetime


#create a dictionary using the reference sheet
os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments')
df = pd.read_excel(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\POA_Eloqua_email_and_lists.xlsx')
df.to_dict()
#os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team')
print(df)

print('in df:')
print('WelcomeJS_01-31-19_Email.xlsx' in df)


#function to find a filename and its corresponding key, then combine the file and its key
######need to get the merge key set up for all spreadsheets
######probably run a test run on one file in a new file, will need to reset the directories
def attachment_combiner(df):
    files = os.listdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test')
    print(files)
    for Eloqua_file in files:
        POA_file = df[Eloqua_file]
        main = pd.read_excel(Eloqua_file,index_col = None)
        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team')
        secondary = pd.read_excel(POA_file,index_col = None)
        combined = pd.merge(main, secondary, sort=False, on='Email Address', how = 'left')
        ##need to set a new directory for the program to save the new files to
        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists')
        #combined = combined.drop_duplicates(subset=['Tax ID','Opportunity ID'],keep='first', inplace=False)
        file_name = 'test results'
        combined.to_excel(file_name, index=False)

attachment_combiner(df)

