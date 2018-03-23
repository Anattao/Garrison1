import requests
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
# =======================================================================================
# functions to get the values of interest
# Market Capital
def market_cap(file):
    location = file.find("Market Cap (intraday) 5")
    i = 0
    length = len("Market Cap (intraday) 5")
    while not file[location + length + i].isalpha():
        i += 1
    data = file[location + length:location + length + i + 1]
    if data == 'N/A':
        data = 0
    else:
        data = data.replace(",","")[:-1]
    return float(data)


# Trailing P/E
def trailing_PE(file):
    location = file.find("Trailing P/E")
    length = len("Trailing P/E ")
    i = 0
    while not file[location + length + i].isalpha():
        i += 1
    data = file[location + length:location + length + i]
    if data == 'N/A':
        data = 0
    else:
        data = data.replace(",","")
    return float(data)


# PEG ratio
def PEG_ratio(file):
    location = file.find("PEG Ratio (5 yr expected) 1")
    length = len("PEG Ratio (5 yr expected) 1")
    i = 0
    while not file[location + length + i].isalpha():
        i += 1
    data =file[location + length:location + length + i]
    if data=='N/A':
        data = 0
    else:
        data = data.replace(",","")
    return float(data)


# Debts Equity ratio
def Debts_Equity_ratio(file):
    location = file.find("Total Debt/Equity (mrq)")
    length = len("Total Debt/Equity (mrq)")
    i = 0
    while not file[location + length + i].isalpha():
        i += 1
    data = file[location + length:location + length + max(i, 3)]
    if data == 'N/A':
        data = 0
    else:
        data = data.replace(",","")
    return float(data)


# Enterprise Value / EBITDA
def EBITDA(file):
    location = file.find("EBITDA")
    length = len("EBITDA ")
    location = file.find("EBITDA", location + length)
    i = 0
    if file[location + length + i].isalpha():
        data = file[location + length:location + length + 3]
        if data=='N/A':
            data = 0
        else:
            data = data.replace(",","")[:-1]
        return float(data)
    else:
        while not file[location + length + i].isalpha():
            i += 1
        data=file[location + length:location + length + i + 1]
        if data == 'N/A':
            data = 0
        else:
            data = data.replace(",","")[:-1]
        return float(data)
    # Beta


def beta(file):
    location = file.find("Beta")
    length = len("Beta ")
    data=file[location + length:location + length + 4]
    if data=='N/A':
        data = 0
    else:
        data = data.replace(",","")
    return float(data)


# operating cash flow
def cash_flow(file):
    location = file.find("StatementOperating Cash Flow (ttm)")
    i = 0
    length = len("StatementOperating Cash Flow (ttm)")
    while not file[location + length + i].isalpha():
        i += 1
    data = file[location + length:location + length + i + 1]
    if data=='N/A':
        data = 0
    else:
        data = data.replace(",","")[:-1]
    return float(data)


# ===============================================================