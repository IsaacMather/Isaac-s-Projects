############to make this work
#1. Set the email subject on line 37
#2. Set the sheet names in the POA lists to their equivelant sheet in the Eloqua results
#3. Update the POA_Eloqua_email_and_lists.xlsx dataframe with the correlating POA lists and Eloqua Results names
#4. Set the eloqua_results_file_locations for the correct month
#5. Set the POA_lists_file_location for the correct month

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

email_subject = 'FW: September'
POA_Eloqua_Team_Dataframe_Location = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments'
eloqua_results_file_locations = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email List Results Sent to POA Team\September'
combined_results_file_location = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Test File Combiner Folder'
POA_lists_file_location = r'C:\Users\isaama2\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Python 3.7\Test Programs\Test Attachments\Email Lists Sent to Eloqua Team\September Files'
opportunity_ID_file_location = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Opportunity ID Reference Folder\Opportunity ID Reference File.xlsx' 
new_combined_file_with_opportunity_ID_directory = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Test File Combiner Folder\Test Files With Opportunity ID'

#This function goes into Outlook, and pulls out files from specific email titles we name. This is how we get the data to begin with. Credit goes to https://stackoverflow.com/questions/39656433/how-to-download-outlook-attachment-from-python-script
def saveattachments(email_subject, eloqua_results_file_locations):
    path = eloqua_results_file_locations
    #print(path)
    today = datetime.date.today()
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) 
    messages = inbox.Items
    #set the title of the email you want to look for here. can we use regex to make flexible email title searches?
    for message in messages:
        if message.Subject == email_subject: #and message.Unread or message.Senton.date() == today:
            #body_content = message.body
            attachments = message.Attachments
            attachment = attachments.Item(1)
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(path, str(attachment)))
                print('Attachment Saved!')

##saveattachments(email_subject, eloqua_results_file_locations)


##This function creates a 'dictionary' that we use to combine the email lists POA sends, with the results we get back from the Eloqua team. We can ask the dictionary, "what's the corresponding file for this particular name" and it will tell us. 
def dictionary_creater(POA_Eloqua_Team_Dataframe_Location):
    os.chdir(POA_Eloqua_Team_Dataframe_Location)
    #get the read excel to just open the dataframe name
    df = pd.read_excel('POA_Eloqua_email_and_lists.xlsx')
    eloqua_results_dictionary = df.set_index('POA File Name')['Eloqua File Name'].to_dict()
    print('Dictionary made')
    return(eloqua_results_dictionary)
    
eloqua_results_dictionary = dictionary_creater(POA_Eloqua_Team_Dataframe_Location)


###function to find a filename and its corresponding key, then combine the file and its key
def attachment_combiner(eloqua_results_dictionary, eloqua_results_file_locations, POA_lists_file_location, combined_results_file_location):
    files = os.listdir(POA_lists_file_location)
    for POA_file in files:
        print(POA_file)
        os.chdir(POA_lists_file_location)
        sheets = pd.ExcelFile(POA_file)
        sheets = sheets.sheet_names
        print(sheets)
        for POA_sheet in sheets:
            Eloqua_file = eloqua_results_dictionary[POA_file]
            os.chdir(POA_lists_file_location)
            #print(Eloqua_file)
            secondary = pd.read_excel(POA_file, sheet_name = POA_sheet, index_col = None)
            os.chdir(eloqua_results_file_locations)
            main = pd.read_excel(Eloqua_file, sheet_name = POA_sheet, index_col = None)
            combined = pd.merge(main, secondary[['Contact: Primary Contact Email','Tax ID']], sort=False, left_on=['Email Address'], right_on=['Contact: Primary Contact Email'], how = 'left').drop(columns=['First Name','Last Name','SFDC Contact ID','Email Send Date','Contact: Primary Contact Email','Total Possible Forwards','Total Hard Bouncebacks'])
            os.chdir(combined_results_file_location)
            file_name = POA_file + ' + ' + POA_sheet + ' done by Python.xlsx'
            combined.to_excel(file_name, index=False)
            print('Eloqua combination complete! file saved to folder!')

attachment_combiner(eloqua_results_dictionary, eloqua_results_file_locations, POA_lists_file_location, combined_results_file_location)


#This function takes the product of the previous function, and combined it with a reference file for opportunity ID's. The product of this function is to have email results data combined with TaxID and Opportunity ID. Now the Eloqua results data is ready for uploading. 
def opportunity_ID_combiner(opportunity_ID_file_location, combined_results_file_location, new_combined_file_with_opportunity_ID_directory):
    files = os.listdir(combined_results_file_location)
    for combined_file in files:
                os.chdir(combined_results_file_location)
                combined_results_file = pd.read_excel(combined_file, index_col = None)
                opportunity_ID_file = pd.read_excel(opportunity_ID_file_location, index_col = None)  #it is set up for .xlsx
                combined = pd.merge(combined_results_file, opportunity_ID_file[['Opportunity ID','Tax ID']], on = 'Tax ID', how = 'left') 
                file_name = combined_file
                os.chdir(new_combined_file_with_opportunity_ID_directory)
                combined.to_excel(file_name, index = False)
                print('Opportunity ID Combination for ' + combined_file + 'complete!')

opportunity_ID_combiner(opportunity_ID_file_location, combined_results_file_location, new_combined_file_with_opportunity_ID_directory)


#This function sends each complete file to a person we choose, with the email subject and content we choose
def mail_the_files_to_ops(new_combined_file_with_opportunity_ID_directory):
        outlook = win32.Dispatch('outlook.application')
        files = os.listdir(new_combined_file_with_opportunity_ID_directory)
        for file in files:
            mail = outlook.CreateItem(0)    
            mail.To = 'isaama2@vsp.com'
            mail.Subject = 'Eloqua Results' + file
            mail.Body = 'Hi, \r\n\nHere is the' + file + '\\r\n\nCheers, \r\nIsaac'
            attachment = file
            mail.Attachments.Add(attachment)
            mail.Send()
            #clarify it worked
        print('Operation successful!')

#mail_the_files_to_ops(new_combined_file_with_opportunity_ID_directory)
