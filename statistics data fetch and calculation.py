# Date: Mar 18th, 2018
# Yiming Alice Chai

# this program is to fetch basic company ratios from yahoo finance-Statistics
# the company list will come as a feed, which includes companies in a certain sector and a specific industry that are the S&P 500
# values to be looked at are Trailing P/E, PEG ratio, Dead Equity Ratio, Market Cap, Enterprise Value/EBIDTA, Beta, VIX, Capital Expenditure
# weighted average values will be calculated for each ratio with repect to their market capital

import requests
from bs4 import BeautifulSoup

#=======================================================================================
# functions to get the values of interest
# Market Capital
def market_cap(file):
    location=file.find("Market Cap (intraday) 5")
    i=0
    length=len("Market Cap (intraday) 5")
    while not file[location+length+i].isalpha():
        i+=1
    return file[location+length:location+length+i+1]

# Trailing P/E
def trailing_PE(file):
    location=file.find("Trailing P/E")
    length=len("Trailing P/E ")
    i=0
    while not file[location+length+i].isalpha():
        i+=1
    return file[location+length:location+length+i]
#PEG ratio
def PEG_ratio(file):
    location=file.find("PEG Ratio (5 yr expected) 1")
    length=len("PEG Ratio (5 yr expected) 1")
    i=0
    while not file[location+length+i].isalpha():
        i+=1
    return file[location+length:location+length+i]    
# Debts Equity ratio
def Debts_Equity_ratio(file):
    location=file.find("Total Debt/Equity (mrq)")
    length=len("Total Debt/Equity (mrq)")
    i=0
    while not file[location+length+i].isalpha():
        i+=1
    return file[location+length:location+length+max(i,3)]  
# Enterprise Value / EBITDA
def EBITDA(file):
    location=file.find("EBITDA")
    length=len("EBITDA ")
    location=file.find("EBITDA",location+length)
    i=0
    if file[location+length+i].isalpha():
        return file[location+length:location+length+3] 
    else:
        while not file[location+length+i].isalpha():
            i+=1
        return file[location+length:location+length+i+1] 
# Beta
def beta(file):
    location=file.find("Beta")
    length=len("Beta ")
    return file[location+length:location+length+4] 
# operating cash flow
def cash_flow(file):
    location=file.find("StatementOperating Cash Flow (ttm)")
    i=0
    length=len("StatementOperating Cash Flow (ttm)")
    while not file[location+length+i].isalpha():
        i+=1
    return file[location+length:location+length+i+1]
#===============================================================

# company list input 
# here is just a example list. it will be replaced by excel input
complist=["WFC", "CCF", "AFG"]
ratios={}
for COMP in complist:
# Generate URL and get html
    URL = "https://finance.yahoo.com/quote/%s/key-statistics?p=%s"%(COMP,COMP)
    print ("getting data from "+URL)
    web = requests.get(URL)
    htmlfile = BeautifulSoup(web.text,'html.parser')

    # interpret html file and extract out the useful part.
    text=htmlfile.get_text()
    start=text.find("Currency in USDValuation Measures")
    end=text.find("Last Split Date")
    usefultext=text[start+15:end+29]    # the string variable that contains all the statistics values
    
    # start to extract the value of interest
    ratios[COMP]={}
    ratios[COMP]['Market_Cap']=market_cap(usefultext)
    ratios[COMP]['Trailing PE']=trailing_PE(usefultext)
    ratios[COMP]['PEG_ratio']=PEG_ratio(usefultext)
    ratios[COMP]['Debts_Equity_ratio']=Debts_Equity_ratio(usefultext)
    ratios[COMP]['EBITDA']=EBITDA(usefultext)
    ratios[COMP]['beta']=beta(usefultext)
    ratios[COMP]['cash_flow']=cash_flow(usefultext)






    


