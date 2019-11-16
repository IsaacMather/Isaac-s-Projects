#scrape the warby parker web page for locationbs, format it so it is excel friendly,

#street, city state, zip, etc,

#have a constantly updated list for the warby parker locations,

#also any of the other providers

#find a way to isolate the locations that are coming soon, and notate that
#potentially "opening"

#https://realpython.com/python-web-scraping-practical-introduction/

from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import re
import urllib

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

raw_html = simple_get('https://www.warbyparker.com/retail')
html = BeautifulSoup(raw_html, 'html.parser')

for i, url in enumerate(html.select('a')):
    print(url)
    #more_raw_html = simple_get(
##for h1, p in zip(html.select('h1'), html.select('p')):
##    print(h1.text, p.text)
