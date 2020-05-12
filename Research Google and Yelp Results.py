#TODO
#3. get the results from the Yelp into the practice info spreadsheet
#4. get the loop to only save the info to the file once, not each time a new practice is called. roughly line 194
#5. check what is going into the query api and main functions, need to clean up which files are being used where




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
##search google maps for the search result, and see if it is open or closed

#1. iterate through the sheets, combining name and address into a text field that we search for. Return the top 5 search results for this thing.

##https://python-googlesearch.readthedocs.io/en/latest/
##https://pypi.org/project/google/
##https://cmdlinetips.com/2018/12/how-to-loop-through-pandas-rows-or-how-to-iterate-over-pandas-rows/

#flowchart https://code2flow.com/3sskBv

#what are static types in python? Does adding them speed up the code? also, cython, for adding c performance to your python, eases the heavy lifying for some computationally intensive sections of your code

#google api key: 
from __future__ import print_function
import argparse
import json
import pprint
import requests
import sys
import urllib

#for querying google
from googlesearch import search
import os
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from googleplaces import GooglePlaces, types, lang


file_location_of_list_of_practices = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Investigating Possibly Closed Locations\Possibly Closed Locations List.xlsx' #need to add this practice file
YOUR_API_KEY = 
directory_where_you_want_to_save_the_new_file = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Investigating Possibly Closed Locations'
new_file_name = 'google_places_results.xlsx'
def search_for_web_results(file_location_of_list_of_practices, YOUR_API_KEY, directory_where_you_want_to_save_the_new_file):
    practices = pd.read_excel(file_location_of_list_of_practices)
    #print(practices)
    for index, row in practices.iterrows(): 
        practice_name = getattr(row, "Common Account Name")
        practice_address = getattr(row, "Physical Street")
        practice_city = getattr(row, "Physical City")
        practice_state = getattr(row, "Physical State/Province")
        practice_zip = str(getattr(row, "Physical Zip/Postal Code"))
        google_places = GooglePlaces(YOUR_API_KEY)
        query_result = google_places.text_search(query = practice_name + ' ' + practice_address + ' ' + practice_city + ' ' + practice_state + ' ' + practice_zip)
        print(len(query_result.places))
        for place in query_result.places:
            try:
                place.get_details()
                practices.iloc[index, 11] = place.name
                practices.iloc[index, 12] = place.formatted_address
                practices.iloc[index, 13] = place.international_phone_number
                practices.iloc[index, 14] = place.place_id
                practices.iloc[index, 15] = place.url
                practices.iloc[index, 16] = place.website
                practices.iloc[index, 17] = place.permanently_closed
            except:
                continue
    os.chdir(directory_where_you_want_to_save_the_new_file)
    practices.to_excel(new_file_name, index = False) 
        
##      google maps api for python: https://developers.google.com/places/web-service/intro, https://developers.google.com/places/web-service/get-api-key, https://console.cloud.google.com/projectselector2/google/maps-apis/overview?pli=1&supportedpurview=project
##      https://stackoverflow.com/questions/50504897/google-places-api-in-python
##        for url in search(practice_name + practice_address, pause = 2.0, stop = 1):


#https://github.com/Yelp/yelp-fusion/blob/master/fusion/python/sample.py

search_for_web_results(file_location_of_list_of_practices, YOUR_API_KEY, directory_where_you_want_to_save_the_new_file)



"""
Yelp Fusion API code sample.
This program demonstrates the capability of the Yelp Fusion API
by using the Search API to query for businesses by a search term and location,
and the Business API to query additional information about the top result
from the search query.
Please refer to http://www.yelp.com/developers/v3/documentation for the API
documentation.
This program requires the Python requests library, which you can install via:
`pip install -r requirements.txt`.
Sample usage of the program:
`python sample.py --term="bars" --location="San Francisco, CA"`
"""




# This client code can run on Python 2.x or 3.x.  Your imports can be
# simpler if you only need one of those.
try:
    # For Python 3.0 and later
    from urllib.error import HTTPError
    from urllib.parse import quote
    from urllib.parse import urlencode
except ImportError:
    # Fall back to Python 2's urllib2 and urllib
    from urllib2 import HTTPError
    from urllib import quote
    from urllib import urlencode


# Yelp Fusion no longer uses OAuth as of December 7, 2017.
# You no longer need to provide Client ID to fetch Data
# It now uses private keys to authenticate requests (API Key)
# You can find it on
# https://www.yelp.com/developers/v3/manage_app
#Client_ID = l-DWj94VPSmyoi7C5-W8cg
API_KEY= 'pmOEz334MEOnzqavqD7HRoUiK9yu9bSpIwEnJs7RQ0rL9Z6CnAi-jLMXXF2cRUPITLSoFmyblpIJbRg1eNEC4qGLdfaP7jDutLPkrto4FLjRcyivKFiD2INNguDFXXYx'


# API constants, you shouldn't have to change these.
API_HOST = 'https://api.yelp.com'
SEARCH_PATH = '/v3/businesses/search'
BUSINESS_PATH = '/v3/businesses/'  # Business ID will come after slash.
SEARCH_LIMIT = 3

google_places_results = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Investigating Possibly Closed Locations\google_places_results.xlsx'
yelp_and_google_results = 'yelp_and_google_results_file.xlsx'


def request(host, path, api_key, url_params=None):
    """Given your API_KEY, send a GET request to the API.
    Args:
        host (str): The domain host of the API.
        path (str): The path of the API after the domain.
        API_KEY (str): Your API Key.
        url_params (dict): An optional set of query parameters in the request.
    Returns:
        dict: The JSON response from the request.
    Raises:
        HTTPError: An error occurs from the HTTP request.
    """
    url_params = url_params or {}
    url = '{0}{1}'.format(host, quote(path.encode('utf8')))
    headers = {'Authorization': 'Bearer %s' % api_key,}
##    print(u'Querying {0} ...'.format(url))
    response = requests.request('GET', url, headers=headers, params=url_params)
    return response.json()

def get_business(api_key, business_id):
    """Query the Business API by a business ID.
    Args: business_id (str): The ID of the business to query.
    Returns: dict: The JSON response from the request."""
    business_path = BUSINESS_PATH + business_id
    return request(API_HOST, business_path, api_key)

def search(api_key, term, location):
    """Query the Search API by a search term and location.
    Args:
        term (str): The search term passed to the API.
        location (str): The search location passed to the API.
    Returns:
        dict: The JSON response from the request."""
    url_params = {'term': term.replace(' ', '+'), 'location': location.replace(' ', '+'), 'limit': SEARCH_LIMIT}
    return request(API_HOST, SEARCH_PATH, api_key, url_params=url_params)


def query_api(term, location, index):
    """Queries the API by the input values from the user.
    Args: term (str): The search term to query.
    location (str): The location of the business to query."""
    response = search(API_KEY, term, location)
    businesses = response.get('businesses')
    if not businesses:
##        print(u'No businesses for {0} in {1} found.'.format(term, location))
        return
    business_id = businesses[0]['id']
    
##    print(u'{0} businesses found, querying business info for the top result "{1}" ...'.format(len(businesses), business_id))
    response = get_business(API_KEY, business_id)
    print(u'Result for business "{0}" found:'.format(business_id))
##    pprint.pprint(response, indent=2)
##    print('name' in response)
    practices = pd.read_excel(google_places_results)
    practices.iloc[index, 18] = response['name']
    practices.iloc[index, 19] = response['display_phone']
    practices.iloc[index, 20] = response['is_closed']
##    practices.iloc[index, 21] = response['location']
    practices.iloc[index, 22] = response['display_phone']        
    os.chdir(directory_where_you_want_to_save_the_new_file)
    practices.to_excel(yelp_and_google_results, index = False)    

##    print('name' in response)
##    return response

##    print(type(response))


def main(google_places_results, yelp_and_google_results):
    practices = pd.read_excel(google_places_results)
    index = 0
    for index, row in practices.iterrows(): #get practices spreadsheet in here
        practice_name = getattr(row, "Common Account Name")
        practice_address = getattr(row, "Physical Street")
        practice_city = getattr(row, "Physical City")
        practice_state = getattr(row, "Physical State/Province")
        practice_zip = str(getattr(row, "Physical Zip/Postal Code"))
        parser = argparse.ArgumentParser()
        parser.add_argument('-q', '--term', dest='term', default = practice_name, type=str, help='Search term (default: %(default)s)')
        parser.add_argument('-l', '--location', dest='location', default = practice_address + ' ' + practice_city + ' ' + practice_state + ' ' + practice_zip, type=str, help='Search location (default: %(default)s)')
        input_values = parser.parse_args()
        
        try:
            index = index + 1
            query_api(input_values.term, input_values.location, index)
            
        except HTTPError as error:
            sys.exit(
            'Encountered HTTP error {0} on {1}:\n {2}\nAbort program.'.format(error.code, error.url, error.read(),))
    

##if __name__ == '__main__':
    main(google_places_results, yelp_and_google_results)


