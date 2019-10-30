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

##https://python-googlesearch.readthedocs.io/en/latest/
##https://pypi.org/project/google/

from googlesearch import search
for url in search('"Breaking Code" WordPress blog', stop=20):
    print(url)
