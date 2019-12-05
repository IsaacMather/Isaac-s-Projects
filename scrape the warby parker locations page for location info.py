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
import numpy

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
    suite_list = []
    city_list = []
    state_list = []
    zip_list = []
    zip_list = numpy.array(zip_list)
    raw_html = simple_get('https://www.warbyparker.com/retail')
    html = BeautifulSoup(raw_html, 'html.parser')
    for elem in html.find_all('a', href=re.compile('retail')):
        print('Address List: ', len(address_list))
        print('City List: ', len(city_list))
        print('State List: ', len(state_list))
        print('Zip List: ', len(zip_list))
        print('Suite List: ', len(suite_list))
        print('City/State/Zip List :', len(city_state_zip_list))
##        print(address_list)
##        print(city_list)
        print(state_list)
        print(zip_list)
        print(suite_list)
        address_url = 'https://www.warbyparker.com' + elem['href']
        raw_address_html = simple_get(address_url)
        try:
            cleaned_raw_address_html = BeautifulSoup(raw_address_html, 'html.parser')
        except:
            pass
        for a, hyperlink in enumerate(cleaned_raw_address_html.find_all('a', href=re.compile('goo'))): #need to add a integer specific spacing for the event that they are opening and there is no zip
            if a > 1:
                zip_list[a] = [' ']
##        for elem in cleaned_raw_address_html.find_all('span'):
            for i, span in enumerate(hyperlink.find_all('span')):
                if i == 0:
                    address_list.append(span.text)
                    print(span.text)
                elif i == 1:
                    if 'Suite' not in span.text and 'Space' not in span.text and 'loor' not in span.text and 'ueen' not in span.text and 'ufferin' not in span.text:
                        suite_list.append(' ')
                        city_state_zip_list.append(span.text)
                        for i, row in enumerate(span.text.split(', ')):
                            print(i)
                            print(row)
                            if i == 0:
                                city_list.append(row)
                            elif i == 1:
                                print(row)
                                for i, row in enumerate(row.split(' ')):
                                    if i == 0:
                                        state_list.append(row)
                                    if i == 1:
                                        zip_list[a] = row
                    elif 'Suite' or 'Space' or 'loor' in span.text:
                        suite_list.append(span.text)    
####                    city_state_zip_list.append(span.text)
                        print(span.text)
                elif i == 2:
##                    city_state_zip_list.append(span.text)
                    for i, row in enumerate(span.text.split(', ')):
##                        print(i)
##                        print(row)
                        if i == 0:
                            city_list.append(row)
                        elif i == 1:
                            for i, row in enumerate(row.split(' ')):
                                if i == 0:
                                    state_list.append(row)
                                if i == 1:
                                    zip_list.append(row)
##                                else:
##                                    zip_list.append(row)
    os.chdir(directory_where_you_want_to_save_the_new_file)
    dictionary = {'Warby Parker Address': address_list,'Suite': suite_list, 'Location City/State/Zip': city_state_zip_list} #,'Location State': state_list,'Location Zip Code': zip_list}
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
