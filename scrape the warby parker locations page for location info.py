#scrape the warby parker web page for locationbs, format it so it is excel friendly,

#street, city state, zip, etc,

#have a constantly updated list for the warby parker locations,

#also any of the other providers

#find a way to isolate the locations that are coming soon, and notate that
#potentially "opening"

#https://realpython.com/python-web-scraping-practical-introduction/

# TODO
#1. filter my URL results so only specific URL's are returned
#2. put the results in a spreadsheet

from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup, NavigableString, Tag
import os
import re
import urllib
import pandas as pd
import time

def simple_get(url):
    """
    Attempts to get the content at `url` by making an HTTP GET request.
    If the content-type of response is some kind of HTML/XML, return the
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    """
    Returns True if the response seems to be HTML, False otherwise.
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200 
            and content_type is not None 
            and content_type.find('html') > -1)


def log_error(e):
    """
    It is always a good idea to log errors. 
    This function just prints them, but you can
    make it do anything.
    """
    print(e)

#answers https://stackoverflow.com/questions/37009287/using-pandas-append-within-for-loop
#https://stackoverflow.com/questions/29110820/how-to-scrape-between-span-tags-using-beautifulsoup


#eyemart
#standon optical https://www.stantonoptical.com/locations/
#cohens fashion https://www.cohensfashionoptical.com/all-locations/
##add "coming soon" option
excel_sheet = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Warby Parker Locations\Warby Parker.xlsx'
directory_where_you_want_to_save_the_new_file = r'C:\Users\isaama2\Desktop\Eloqua Data Combiner Files\Warby Parker Locations'
new_file_name = "warby_parker_locations.xlsx"
def pull_warby_parker_locations():
    print(new_file_name)
    address_list = []
    city_state_zip_list = []
    raw_html = simple_get('https://www.warbyparker.com/retail')
    html = BeautifulSoup(raw_html, 'html.parser')
    for elem in html.find_all('a', href=re.compile('retail')):
        address_url = 'https://www.warbyparker.com' + elem['href']
        raw_address_html = simple_get(address_url)
        cleaned_raw_address_html = BeautifulSoup(raw_address_html, 'html.parser')
        for hyperlink in cleaned_raw_address_html.find_all('a', href=re.compile('goo')):
##        for elem in cleaned_raw_address_html.find_all('span'):
            for i, span in enumerate(hyperlink.find_all('span')):
                if i == 0:
                    address_list.append(span.text)
                    print(span.text)
                elif i == 1:
                    city_state_zip_list.append(span.text)
                    print(span.text)
    os.chdir(directory_where_you_want_to_save_the_new_file)
    dictionary = {'Warby Parker Address': address_list,'Warby Parker City/State/Zip': city_state_zip_list}
    df = pd.DataFrame(dictionary)
    df.to_excel(new_file_name, index = False)

pull_warby_parker_locations()

#extracting text between an element https://stackoverflow.com/questions/16835449/python-beautifulsoup-extract-text-between-element
stanton_file_name = "stanton_optical_locations.xlsx"
def stanton_optical_locations():
    raw_html = simple_get('https://www.stantonoptical.com/locations/')
    html = BeautifulSoup(raw_html,'html.parser')
    state_list = []
    address_list = []
    city_state_zip_list = []
    for ptag in html.find_all('p'):
      for i, content in enumerate(ptag.contents):
          if content.string is not None:
              if content.string is not None: #filter out the spurious results
                  if i == 0:
                      state_list.append(content.string)
                  elif i == 1:
                      address_list.append(content.string)
                      print(content)
                  elif i == 3:
                      city_state_zip_list.append(content.string)
                      print(content)
                  elif i > 3:
                      continue
                
    os.chdir(directory_where_you_want_to_save_the_new_file)
    address = pd.series(address_list, name = 'Addresses')
    city = pd.series(city_state_zip_list, name = 'City/State/Zip')
    df = pd.concat([address,city],axis = 1)
##    dictionary = {'Stanton Optical Address':address_list,'Stanton Optical City/State/Zip': city_state_zip_list}
##    df = pd.DataFrame.from_dict(dictionary)
    df.to_excel(stanton_file_name, index = False)                    
##          print(ptag.contents)
##        next_s = element.nextSibling
##        if not (next_s and isinstance(next_s,NavigableString)):
##            continue
##        next2_s = next_s.nextSibling
##        if next2_s and isinstance(next2_s,Tag) and next2_s.name == 'br':
##            text = str(next_s).strip()
##            if text:
##                print("Found:", next_s)

##stanton_optical_locations()

##
##for i, url in enumerate(html.select('a')):
##    print(url)
    #more_raw_html = simple_get(
##for h1, p in zip(html.select('h1'), html.select('p')):
##    print(h1.text, p.text)
