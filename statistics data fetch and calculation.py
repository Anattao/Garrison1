# Date: Mar 18th, 2018
# Yiming Alice Chai

# this program is to fetch basic company ratios from yahoo finance-Statistics
# the company list will come as a feed, which includes companies in a certain sector and a specific industry that are the S&P 500
# values to be looked at are Trailing P/E, PEG ratio, Dead Equity Ratio, Market Cap, Enterprise Value/EBIDTA, Beta, VIX, Capital Expenditure
# weighted average values will be calculated for each ratio with repect to their market capital

import requests
from bs4 import BeautifulSoup

# company list input 
# here is just a example list. it will be replaced by excel input
complist=["WFC", "CCF", "AFG"]
symbol=complist[0]

# Generate URL and get html
URL = "https://finance.yahoo.com/quote/%s/key-statistics?p=%s"%(symbol,symbol)
print ("getting data from "+URL)
web = requests.get(URL)
htmlfile = BeautifulSoup(web.text,'html.parser')

# interpret html file and extract out the useful part.
text=htmlfile.get_text()
start=text.find("Currency in USDValuation Measures")
end=text.find("Last Split Date")
usefultext=text[start+15:end+29]    # the string variable that contains all the statistics values
print (usefultext)

# next proceed to get the value of interested



    


# clean html to useful text
#for script in htmlfile(["script", "style"]):
#for item in htmlfile:
#    print(item)
#print (htmlfile.stripped_string)
