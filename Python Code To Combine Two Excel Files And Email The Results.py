#this library is to set the directory
import os

#these libraries are to manipulate the excel files
import pandas as pd
import openpyxl

#this library is to send email
import win32com.client as win32

#set the directory you want
os.chdir(r'Directory path')

#get current working directory
print(os.getcwd())

#pull attachments from MS Outlook
#https://stackoverflow.com/questions/39656433/how-to-download-outlook-attachment-from-python-script

content here




#opens the files and selects only the columns we want
main = pd.read_excel(r'File Name.xlsx',index_col = None, usecols='Selected Columns')
secondary = pd.read_excel(r'Secondary File Name.xlsx',index_col = None, usecols = 'Selected Columns')

#combine the files
combined = pd.merge(main, secondary, sort=False, on='Chosen Key', how = 'left')

#drop any duplicates created in the process
#combined = combined.drop_duplicates(subset=['Column1','Column2'],keep='first', inplace=False)

#export to the current working directory
file_name = 'New File Name.xlsx'
combined.to_excel(file_name, index=False)

#find out which directory we are in
#directory = os.getcwd()

#send the file in an Outlook email
import win32com.client
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Email Address'
mail.Subject = 'Email Name'
mail.Body = 'Email Body'
attachment = r'File Path'
#attachment = directory + \ + file_name #me trying to create a responsive file path
mail.Attachments.Add(attachment)
mail.Send()

#clarify it worked
print('Operation successful!')
