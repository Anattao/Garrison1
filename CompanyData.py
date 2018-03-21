import requests
from bs4 import BeautifulSoup
import xlrd

workbook = xlrd.open_workbook()
# Ask for company name input
print("Please type in the company name code==>")
symbol = str(input())

# Generate URL and get html
URL = "https://finance.yahoo.com/quote/%s/history?p=%s" % (symbol, symbol)
print("getting data from " + URL)
web = requests.get(URL)
htmlfile = BeautifulSoup(web.text, 'html.parser')

# clean html to useful text
for script in htmlfile(["script", "style"]):
    script.extract()

# gather data into a list and then extract the useful portion
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
# Writing datapoints into excel file
outfile = symbol + '.csv'
data = 'Date, Open, High, Low, Close, Adj Close, Volume\n'

for dailydata in Datalist:
    addindata = ",".join(item for item in dailydata)
    data = data + addindata + "\n"

with open(outfile, 'w') as f:
    f.write(data)

print("Finished")











