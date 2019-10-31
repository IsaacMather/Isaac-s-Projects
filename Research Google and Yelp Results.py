##What's the challenge:
##	Ensuring the provider database is up to date and accurate
##	
##Can we use the Google search or yelo api, to search an address, and get back information about the location? 
##Do it by step by step, and ask for user input, and see if this is what they are looking for. 
##Keep the search result link, the yelp and the google result, for each line item, to act as a reference 
##
##Take a practice I know exists, use it as a test, have the program rip through the spreadsheet of addresses and return search results. 
##
##Create a flow chart, and reconnect with Justin.


#1. iterate through the sheets, combining name and address into a text field that we search for. Return the top 5 search results for this thing. 




##https://python-googlesearch.readthedocs.io/en/latest/
##https://pypi.org/project/google/
##https://cmdlinetips.com/2018/12/how-to-loop-through-pandas-rows-or-how-to-iterate-over-pandas-rows/\

from googlesearch import search
import os
import pandas as pd
import numpy as np

file_location_of_list_of_practices = '' #need to add this practice file
def search_for_web_results(file_location_of_list_of_practices):
    practices = pd.ExcelFile(file_location_of_list_of_practices)
    for index, row in practices.iterrows(): #don't use this, use https://stackoverflow.com/questions/16476924/how-to-iterate-over-rows-in-a-dataframe-in-pandas/55557758#55557758
        for url in search('"Breaking Code" WordPress blog', stop=20):
        print(url)



#search_for_web_results(file_location_of_list_of_practices)

