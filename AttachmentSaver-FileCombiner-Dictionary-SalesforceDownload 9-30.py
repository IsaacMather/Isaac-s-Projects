##              To Do List
##              1. clean up the directories by making them dynamic
##              2. get the new file out of the attachment combiner function
##              3. get the file into the file mailer function
##              4. need to retrieve the opportunity ID from salesforce\
##              5. need to combine the opportunity id with our combined results lists before they are mailed to the OPS team


#this library is to set the directory
import os

#these libraries are to manipulate the excel files
import pandas as pd
import openpyxl
import csv


#this library is to interact with MS Outlook
import win32com.client as win32

#this library is to check todays date
import datetime


##The below code helps by downloading the attachments from outlook emails that are 'Unread' (and changes the mail to Read.) or from 'Today's' date.without altering the file name. Just pass the 'Subject' argument.
##https://stackoverflow.com/questions/39656433/how-to-download-outlook-attachment-from-python-script
eloqua_results_file_locations = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments'
##email_subject = 'Working with Global Analytics and Insights Team Checklist (002)'
##
##def saveattachments(email_subject, eloqua_results_file_locations):
##    eloqua_results_file_locations = os.path.expanduser(eloqua_results_directory)
##    print(path)
##    today = datetime.date.today()
##    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
##    inbox = outlook.GetDefaultFolder(6) 
##    messages = inbox.Items
##    #set the title of the email you want to look for here. can we use regex to make flexible email title searches?
##    for message in messages:
##        if message.Subject == email_subject: #and message.Unread or message.Senton.date() == today:
##            #body_content = message.body
##            attachments = message.Attachments
##            attachment = attachments.Item(1)
##            for attachment in message.Attachments:
##                attachment.SaveAsFile(os.path.join(eloqua_results_file_locations, str(attachment)))
##                #if message.Subject == subject and message.Unread:
##                #    message.Unread = False
##                print('Attachment Saved!')
##                break
##            
##saveattachments(email_subject, eloqua_results_file_locations)
##
##
##







POA_Eloqua_Team_Dataframe_Location = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments'

##POA_lists = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team\Old Files'
##combined_directory = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists'

#~~~create a dictionary using the reference sheet. references the dictionary to to know which file sent from the POA team to the eloqua team to use as a resource to retrieve TaxID. Then saves the combined sheet
def dictionary_creater(POA_Eloqua_Team_Dataframe_Location):
    os.chdir(POA_Eloqua_Team_Dataframe_Location)
    #get the read excel to just open the dataframe name
    df = pd.read_excel('POA_Eloqua_email_and_lists.xlsx')
    eloqua_results_dictionary = df.set_index('Eloqua File Name')['POA File Name'].to_dict()
    return(eloqua_results_dictionary)

eloqua_results_dictionary = dictionary_creater(eloqua_results_directory)


###~~~this code checks if we successfully created the dictionary to correlate eloqua results files with POA lists
##for key, val in eloqua_results_dictionary.items():
##    print(key, "=>", val)





#function to find a filename and its corresponding key, then combine the file and its key
def attachment_combiner(eloqua_results_dictionary, eloqua_results_file_locations):
    files = os.listdir(eloqua_results)
    for Eloqua_file in files:
        POA_file = eloqua_results_dictionary[Eloqua_file]
        print(Eloqua_file)
        os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\Test')
        sheets = pd.ExcelFile(Eloqua_file)
        sheets = sheets.sheet_names
        for Eloqua_sheet in sheets:
            #print(Eloqua_sheets)    
            main = pd.read_excel(Eloqua_file, sheet_name = Eloqua_sheet, index_col = None)
            os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team\Old Files')
            secondary = pd.read_excel(POA_file, sheet_name = Eloqua_sheet, index_col = None)
            combined = pd.merge(main, secondary, sort=False, left_on=['Email Address'], right_on=['Contact: Primary Contact Email'], how = 'left')
            os.chdir(r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Combined Lists')
            file_name = Eloqua_file + ' + ' + Eloqua_sheet + ' done by Python.xlsx'
            combined.to_excel(file_name, index=False)
            print('combination complete! file saved to folder!')

attachment_combiner(eloqua_results_dictionary, eloqua_results_file_locations)











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
