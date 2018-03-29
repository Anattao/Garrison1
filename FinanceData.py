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
allUserData2=uData.get_financials()[-1]
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

print(userCompanyData)

complist = industryComp
ratios = {}
allData = {}
allData2 = {}
totalMarketCap=0
for COMP in complist:
    cData=iexfinance.Stock(COMP)
    ratios[COMP] = {}
    allData[COMP] = cData.get_key_stats()
    allData2[COMP] = cData.get_financials()[-1]
    ratios[COMP]['Market Cap']=allData[COMP]['marketcap']
    ratios[COMP]['Beta'] = allData[COMP]['beta']
    ratios[COMP]['Debt Equity Ratio'] = allData2[COMP]['totalLiabilities']/allData2[COMP]['shareholderEquity']
    ratios[COMP]['PE Ratio'] = (allData[COMP]['peRatioHigh']+allData[COMP]['peRatioLow'])/2
    ratios[COMP]['Price to Sales'] = allData[COMP]['priceToSales']
    ratios[COMP]['Price to Book'] = allData[COMP]['priceToBook']
    ratios[COMP]['Short Ratio'] = allData[COMP]['shortRatio']
    ratios[COMP]['50 Day Moving Average'] = allData[COMP]['day50MovingAvg']
    ratios[COMP]['200 Day Moving Average'] = allData[COMP]['day200MovingAvg']
    totalMarketCap+=allData[COMP]['marketcap']
print(ratios)

start = datetime.date.today()-datetime.timedelta(100)
end = datetime.date.today()

stockData = iexfinance.get_historical_data(symbol,start,end)
print(stockData)

weightIndAvg={}
weightIndAvg[industry]={'Market Cap': 0, 'Beta': 0, 'Debt Equity Ratio': 0, 'PE Ratio': 0, 'Price to Sales': 0, 'Price to Book': 0, 'Short Ratio': 0, '50 Day Moving Average': 0, '200 Day Moving Average': 0}
for keys,values in ratios.items():
    for keys1,values1 in values.items():
        if keys1 == 'Market Cap':
            weight = values1/totalMarketCap
            weightIndAvg[industry]['Market Cap']+=weight*values1
        if keys1 == 'Beta':
            weightIndAvg[industry]['Beta']+=weight*values1
        if keys1 == 'Debt Equity Ratio':
            weightIndAvg[industry]['Debt Equity Ratio']+=weight*values1
        if keys1 == 'PE Ratio':
            weightIndAvg[industry]['PE Ratio']+=weight*values1
        if keys1 == 'Price to Sales':
            weightIndAvg[industry]['Price to Sales']+=weight*values1
        if keys1 == 'Price to Book':
            weightIndAvg[industry]['Price to Book']+= weight*values1
        if keys1 == 'Short Ratio':
            weightIndAvg[industry]['Short Ratio']+=weight*values1
        if keys1 == '50 Day Moving Average':
            weightIndAvg[industry]['50 Day Moving Average']+=weight*values1
        if keys1 == '200 Day Moving Average':
            weightIndAvg[industry]['200 Day Moving Average']+=weight*values1

print(weightIndAvg)




