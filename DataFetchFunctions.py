import iexfinance
import pandas as pd
import numpy as np
import datetime

uForm = pd.read_excel('Garrison Project.xlsm','User Form')
iForm = pd.read_excel('Garrison Project.xlsm','Companies by Industry')
userData=uForm['Response'].values
symbol = userData[1]
sector = userData[2]
industry = userData[3]
industryComp = iForm[industry].dropna(axis=0,how='all')
industryComp = industryComp.values

uData = iexfinance.Stock(symbol)
allUserData=uData.get_key_stats()
allUserData2=uData.get_financials()[3]
#print(allUserData2)
userCompanyData={}
userCompanyData[symbol]={}
userCompanyData[symbol]['Market Cap']=allUserData['marketcap']
userCompanyData[symbol]['Beta'] = allUserData['beta']
userCompanyData[symbol]['Debt Equity Ratio'] = allUserData2['totalLiabilities']/allUserData2['shareholderEquity']
userCompanyData[symbol]['PE Ratio'] = (allUserData['peRatioHigh']+allUserData['peRatioLow'])/2
userCompanyData[symbol]['Price to Sales'] = allUserData['priceToSales']
userCompanyData[symbol]['Price to Book'] = allUserData['priceToBook']
userCompanyData[symbol]['Short Ratio'] = allUserData['shortRatio']
userCompanyData[symbol]['50 Day Moving Average'] = allUserData['day50MovingAvg']
userCompanyData[symbol]['200 Day Moving Average'] = allUserData['day200MovingAvg']

#print(userCompanyData)

complist = industryComp
ratios = {}
allData = {}
allData2 = {}
for COMP in complist:
    cData=iexfinance.Stock(COMP)
    ratios[COMP] = {}
    allData[COMP] = cData.get_key_stats()
    allData2[COMP] = cData.get_financials()[3]
    ratios[COMP]['Market Cap']=allData[COMP]['marketcap']
    ratios[COMP]['Beta'] = allData[COMP]['beta']
    ratios[COMP]['Debt Equity Ratio'] = allData2[COMP]['totalLiabilities']/allData2[COMP]['shareholderEquity']
    ratios[COMP]['PE Ratio'] = (allData[COMP]['peRatioHigh']+allData[COMP]['peRatioLow'])/2
    ratios[COMP]['Price to Sales'] = allData[COMP]['priceToSales']
    ratios[COMP]['Price to Book'] = allData[COMP]['priceToBook']
    ratios[COMP]['Short Ratio'] = allData[COMP]['shortRatio']
    ratios[COMP]['50 Day Moving Average'] = allData[COMP]['day50MovingAvg']
    ratios[COMP]['200 Day Moving Average'] = allData[COMP]['day200MovingAvg']
#print(ratios)

# calculated weighted average
sumMC=0  #sum of Market Capital of representative companies
for item in ratios.keys():
    sumMC+=ratios[item]['Market Cap']

# function to calculate weighted average
def Weighted_Average(term):
    weighted_avg=0
    for comp in ratios.keys():
        weighted_avg+=ratios[comp][term]*ratios[comp]['Market Cap']/sumMC
    return weighted_avg
#compile data tile and value to two lists
ratio_output={}
ratio_output['weighted_avg_ratios']={}
for key in ratios[complist[0]].keys():
    ratio_output['weighted_avg_ratios'][key]=Weighted_Average(key)
    
# historical stock price    
start = datetime.date.today()-datetime.timedelta(100)
end = datetime.date.today()

stockData = iexfinance.get_historical_data(symbol,start,end)
#print(stockData)




#--------------------------------------
# write stock data into file named after the company of interest
df1 = pd.DataFrame(stockData[symbol]).T          # stock price data
df2 = pd.DataFrame(ratio_output)          #weighted average of ratios
outfile=symbol+'.xlsx'

writer = pd.ExcelWriter(outfile, engine='xlsxwriter')   # Creating Excel Writer Object from Pandas
#workbook=writer.book
df1.to_excel(writer,sheet_name=symbol,startrow=0 , startcol=0)
df2.to_excel(writer,sheet_name=symbol,startrow=0, startcol=8)
writer.save()


