import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from DataFetchFunctions import *
#----------------------------------------------------------------------------
#get user input
uForm = pd.read_excel('Garrison Project.xlsm','User Form')
iForm = pd.read_excel('Garrison Project.xlsm','Companies by Industry')
userData=uForm['Response'].values
symbol = userData[1]
sector = userData[2]
industry = userData[3]
industryComp = iForm[industry].dropna(axis=0,how='all')
industryComp = industryComp.values
#----------------------------------------------------------------------------
# Generate URL and get html for user input company
URL = "https://finance.yahoo.com/quote/%s/history?p=%s" % (symbol, symbol)
print("getting data from " + URL)
web = requests.get(URL)
htmlfile = BeautifulSoup(web.text, 'html.parser')

# clean html to useful text
for script in htmlfile(["script", "style"]):
    script.extract()

# gather stock data
alldata = []
for string in htmlfile.stripped_strings:
    alldata.append(string)
start = alldata.index("Volume") + 1
end = alldata.index("*Close price adjusted for splits.") - 1
usefuldata = alldata[start:end]

# organize data to be grouped by per day
Datalist = []
i = 1
for datapoint in usefuldata:
    if not datapoint[0].isdigit() and i % 7 == 1:
        Dataperday = []
        Datalist.append(Dataperday)
        i = 1
    if not datapoint[0].isdigit() and i % 7 == 3:
        i = 0
    datapoint = datapoint.replace(",", "")
    Dataperday.append(datapoint)
    i += 1
print(Datalist)
#----------------------------------------------------------------------------
# get ratio data for user input company
URL = "https://finance.yahoo.com/quote/%s/key-statistics?p=%s" % (symbol, symbol)
print("getting data from " + URL)
web = requests.get(URL)
htmlfile = BeautifulSoup(web.text, 'html.parser')
text = htmlfile.get_text()
start = text.find("Currency in USDValuation Measures")
end = text.find("Last Split Date")
usefultext = text[start + 15:end + 29]  # the string variable that contains all the statistics values
ucompany = [symbol,0,0,0,0,0,0,0]
ucompany[1]=market_cap(usefultext)
ucompany[2]=trailing_PE(usefultext)
ucompany[3]=PEG_ratio(usefultext)
ucompany[4]=Debts_Equity_ratio(usefultext)
ucompany[5]=EBITDA(usefultext)
ucompany[6]=beta(usefultext)
ucompany[7]=cash_flow(usefultext)
print(ucompany)
#----------------------------------------------------------------------------
# Writing datapoints into excel file
#outfile = symbol + '.csv'
#data = 'Date, Open, High, Low, Close, Adj Close, Volume\n'

#for dailydata in Datalist:
#    addindata = ",".join(item for item in dailydata)
#    data = data + addindata + "\n"

#with open(outfile, 'w') as f:
#    f.write(data)

#print("Finished")

#----------------------------------------------------------------------------
# get weighted industry average for selected industry
complist = industryComp
ratios = {}
ratiosum=[industry,0,0,0,0,0,0,0]
for COMP in complist:
    # Generate URL and get html
    URL = "https://finance.yahoo.com/quote/%s/key-statistics?p=%s" % (COMP, COMP)
    print("getting data from " + URL)
    web = requests.get(URL)
    htmlfile = BeautifulSoup(web.text, 'html.parser')

    # interpret html file and extract out the useful part.
    text = htmlfile.get_text()
    start = text.find("Currency in USDValuation Measures")
    end = text.find("Last Split Date")
    usefultext = text[start + 15:end + 29]  # the string variable that contains all the statistics values
    # start to extract the value of interest
    ratios[COMP] = {}
    ratios[COMP]['Market_Cap'] = market_cap(usefultext)
    ratios[COMP]['Trailing PE'] = trailing_PE(usefultext)
    ratios[COMP]['PEG_ratio'] = PEG_ratio(usefultext)
    ratios[COMP]['Debts_Equity_ratio'] = Debts_Equity_ratio(usefultext)
    ratios[COMP]['EBITDA'] = EBITDA(usefultext)
    ratios[COMP]['beta'] = beta(usefultext)
    ratios[COMP]['cash_flow'] = cash_flow(usefultext)
    # generate sum of ratio data for each ratio
    ratiosum[1]+=market_cap(usefultext)
    ratiosum[2]+=trailing_PE(usefultext)
    ratiosum[3]+=PEG_ratio(usefultext)
    ratiosum[4]+=Debts_Equity_ratio(usefultext)
    ratiosum[5]+=EBITDA(usefultext)
    ratiosum[6]+=beta(usefultext)
    ratiosum[7]+=cash_flow(usefultext)
print(ratiosum)
# gets weighted averages for all ratios and puts it in a list
weightedvalues=[industry,0,0,0,0,0,0,0]
for keys,values in ratios.items():
    for keys1,values1 in values.items():
        if keys1 == 'Market_Cap':
            weightedvalues[1]=ratiosum[1]/len(complist)
            marketcapWeight=(values1/ratiosum[1])
        if keys1 == 'Trailing PE':
            weightedvalues[2]+=marketcapWeight*values1
        if keys1 == 'PEG_ratio':
            weightedvalues[3]+=marketcapWeight*values1
        if keys1 == 'Debts_Equity_ratio':
            weightedvalues[4]+=marketcapWeight*values1
        if keys1 == 'EBITDA':
            weightedvalues[5]+=marketcapWeight*values1
        if keys1 == 'beta':
            weightedvalues[6]+=marketcapWeight*values1
        if keys1 == 'cash_flow':
            weightedvalues[7]+=marketcapWeight*values1
print(weightedvalues)
#----------------------------------------------------------------------------
labelsStockData=['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']
labelsRatios=['Ratio_Type','Market_Cap','Trailing PE','PEG_ratio','Debts_Equity_ratio','EBITDA','beta','cash_flow']
finalStockData= [labelsStockData]+Datalist
finalratios = [labelsRatios,ucompany,weightedvalues]
print(finalStockData)
df1 = pd.DataFrame(finalStockData[1:],columns=finalStockData[0])
df2 = pd.DataFrame(finalratios[1:], columns= finalratios[0])
print(df1)
print(df2)
writer = pd.ExcelWriter(symbol+'.xlsx', engine='xlsxwriter')   # Creating Excel Writer Object from Pandas
workbook=writer.book
df1.to_excel(writer,sheet_name=symbol,startrow=0 , startcol=0)
df2.to_excel(writer,sheet_name=symbol,startrow=0, startcol=10)










