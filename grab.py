from random import vonmisesvariate
from h11 import Data
import pandas as pd
import requests
import numpy as np
import os
import xlsxwriter
import math
from bs4 import BeautifulSoup
import time
import random
from numpy import loadtxt
# Aktien auswählen:

fileObj = open("Array.txt", "r")  # opens the file in read mode
words = fileObj.read().splitlines()  # puts the file into an array
tickerArray = words
fileObj.close()


def execute(ticker):
    def getdata(url):
        r = requests.get(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
        data = pd.read_html(r.text)
        return data

    # Daten von Yahoo Finance Statistics in Variable statistics_data ziehen
    statistics_url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics?p={ticker}'
    statistics_data = getdata(statistics_url)

    # Daten von Yahoo Finance Holders in Variable statistics_data ziehen und NaN entfernen
    holders_url = f'https://finance.yahoo.com/quote/{ticker}/holders?p={ticker}'
    holders_data = getdata(holders_url)
    try:
        holders_data[0]
    except:
        pass
    else:
        holders_data = pd.DataFrame(holders_data[0])
        holders_data = holders_data.fillna("N/A")

    try:
        holders_data_check = getdata(holders_url)
        holders_data_check[1]
    except:
        pass
    else:
        holders_url = f'https://finance.yahoo.com/quote/{ticker}/holders?p={ticker}'
        holders_data1 = getdata(holders_url)
        holders_data1 = pd.DataFrame(holders_data1[1])
        holders_data1 = holders_data1.fillna("N/A")

    try:
        holders_data_check = getdata(holders_url)
        holders_data_check[2]
    except:
        pass
    else:
        holders_url = f'https://finance.yahoo.com/quote/{ticker}/holders?p={ticker}'
        holders_data2 = getdata(holders_url)
        holders_data2 = pd.DataFrame(holders_data2[2])
        holders_data2 = holders_data2.fillna("N/A")

    # Rohdaten auswählen und NaN entfernen

    data0 = statistics_data[0]
    df0 = data0
    df0 = df0.fillna('N/A')

    data1 = statistics_data[1]
    df1 = data1
    df1 = df1.fillna('N/A')

    data2 = statistics_data[2]
    df2 = data2
    df2 = df2.fillna('N/A')

    data3 = statistics_data[3]
    df3 = data3
    df3 = df3.fillna('N/A')

    data4 = statistics_data[4]
    df4 = data4
    df4 = df4.fillna('N/A')

    data5 = statistics_data[5]
    df5 = data5
    df5 = df5.fillna('N/A')

    data6 = statistics_data[6]
    df6 = data6
    df6 = df6.fillna('N/A')

    data7 = statistics_data[7]
    df7 = data7
    df7 = df7.fillna('N/A')

    data8 = statistics_data[8]
    df8 = data8
    df8 = df8.fillna('N/A')

    data9 = statistics_data[9]
    df9 = data9
    df9 = df9.fillna('N/A')

    # Excel Datei erstellen, wenn sie noch nicht existiert
    fileName = (ticker + '.xlsx')

    if not os.path.exists(fileName):
        workbook = xlsxwriter.Workbook(fileName)
        workbook.close()

    def getdata1(url):
        r = requests.get(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
        return r.text

    # Daten vom Income Statement von der Website ziehen
    financial_url = f'https://finance.yahoo.com/quote/{ticker}/financials?p={ticker}'
    financial_data = getdata1(financial_url)

    soup = BeautifulSoup(financial_data, 'lxml')

    close_price = [entry.text for entry in soup.find_all(
        'span', {'class': 'Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'})]

    features = soup.find_all('div', class_='D(tbr)')

    headers = []
    temp_list = []
    label_list = []
    final = []
    index = 0
    # create headers
    for item in features[0].find_all('div', class_='D(ib)'):
        headers.append(item.text)
    # statement contents
    while index <= len(features)-1:
        # filter for each line of the statement
        temp = features[index].find_all('div', class_='D(tbc)')
        for line in temp:
            # each item adding to a temporary list
            temp_list.append(line.text)
        # temp_list added to final list
        final.append(temp_list)
        # clear temp_list
        temp_list = []
        index += 1
    dfn = pd.DataFrame(final[1:])
    dfn.columns = headers
    dfn = dfn.fillna("N/A")

    # Daten aus dem Balance Sheet ziehen
    balance_url = f'https://finance.yahoo.com/quote/{ticker}/balance-sheet?p={ticker}'
    balance_data = getdata1(balance_url)

    soup = BeautifulSoup(balance_data, 'lxml')

    close_price = [entry.text for entry in soup.find_all(
        'span', {'class': 'Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'})]

    features = soup.find_all('div', class_='D(tbr)')

    headers = []
    temp_list = []
    label_list = []
    final = []
    index = 0
    # create headers
    for item in features[0].find_all('div', class_='D(ib)'):
        headers.append(item.text)
    # statement contents
    while index <= len(features)-1:
        # filter for each line of the statement
        temp = features[index].find_all('div', class_='D(tbc)')
        for line in temp:
            # each item adding to a temporary list
            temp_list.append(line.text)
        # temp_list added to final list
        final.append(temp_list)
        # clear temp_list
        temp_list = []
        index += 1
    dfe = pd.DataFrame(final[1:])
    dfe.columns = headers
    dfe = dfe.fillna("N/A")

    # Daten aus Cash Flow ziehen
    Cashflow_url = f'https://finance.yahoo.com/quote/{ticker}/cash-flow?p={ticker}'
    Cashflow_data = getdata1(Cashflow_url)

    soup = BeautifulSoup(Cashflow_data, 'lxml')

    close_price = [entry.text for entry in soup.find_all(
        'span', {'class': 'Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)'})]

    features = soup.find_all('div', class_='D(tbr)')

    headers = []
    temp_list = []
    label_list = []
    final = []
    index = 0
    # create headers
    for item in features[0].find_all('div', class_='D(ib)'):
        headers.append(item.text)
    # statement contents
    while index <= len(features)-1:
        # filter for each line of the statement
        temp = features[index].find_all('div', class_='D(tbc)')
        for line in temp:
            # each item adding to a temporary list
            temp_list.append(line.text)
        # temp_list added to final list
        final.append(temp_list)
        # clear temp_list
        temp_list = []
        index += 1
    dfi = pd.DataFrame(final[1:])
    dfi.columns = headers
    dfi = dfi.fillna("N/A")

    # Exceldatei ansprechen
    tickerWorkbook = xlsxwriter.Workbook((ticker+".xlsx"))

    # alle Daten aus Valuation Measures in Excel übertragen
    Sheet4 = tickerWorkbook.add_worksheet('Balance Sheet')
    Sheet3 = tickerWorkbook.add_worksheet('Income Statement')
    Sheet5 = tickerWorkbook.add_worksheet('Cash Flow')
    Sheet1 = tickerWorkbook.add_worksheet('Statistics')
    Sheet2 = tickerWorkbook.add_worksheet('Holders')

    # alle Daten aus Key Statistics in Excel übertragen

    if 0 in df0:

        for i in range(0, len(df0[0])):
            Sheet1.write("A" + str(i + 1), df0[0][i])
            Sheet1.write("B" + str(i + 1), df0[1][i])

    if 0 in df1:
        l1 = len(df0[0])
        for i in range(0, len(df1[0])):
            Sheet1.write("A" + str(i + 1 + l1), df1[0][i])
            Sheet1.write("B" + str(i + 1 + l1), df1[1][i])

    if 0 in df2:
        l2 = l1 + len(df1[0])
        for i in range(0, len(df2[0])):
            Sheet1.write("A" + str(i + 1 + l2), df2[0][i])
            Sheet1.write("B" + str(i + 1 + l2), df2[1][i])

    if 0 in df3:
        l3 = l2 + len(df2[0])
        for i in range(0, len(df3[0])):
            Sheet1.write("A" + str(i + 1 + l3), df3[0][i])
            Sheet1.write("B" + str(i + 1 + l3), df3[1][i])

    if 0 in df4:
        l4 = l3 + len(df3[0])
        for i in range(0, len(df4[0])):
            Sheet1.write("A" + str(i + 1 + l4), df4[0][i])
            Sheet1.write("B" + str(i + 1 + l4), df4[1][i])

    if 0 in df5:
        l5 = l4 + len(df4[0])
        for i in range(0, len(df5[0])):
            Sheet1.write("A" + str(i + 1 + l5), df5[0][i])
            Sheet1.write("B" + str(i + 1 + l5), df5[1][i])

    if 0 in df6:
        l6 = l5 + len(df5[0])
        for i in range(0, len(df6[0])):
            Sheet1.write("A" + str(i + 1 + l6), df6[0][i])
            Sheet1.write("B" + str(i + 1 + l6), df6[1][i])

    if 0 in df7:
        l7 = l6 + len(df6[0])
        for i in range(0, len(df7[0])):
            Sheet1.write("A" + str(i + 1 + l7), df7[0][i])
            Sheet1.write("B" + str(i + 1 + l7), df7[1][i])

    if 0 in df8:
        l8 = l7 + len(df7[0])
        for i in range(0, len(df8[0])):
            Sheet1.write("A" + str(i + 1 + l8), df8[0][i])
            Sheet1.write("B" + str(i + 1 + l8), df8[1][i])

    if 0 in df9:
        l9 = l8 + len(df8[0])
        for i in range(0, len(df9[0])):
            Sheet1.write("A" + str(i + 1 + l9), df9[0][i])
            Sheet1.write("B" + str(i + 1 + l9), df9[1][i])

    # alle Daten aus Holders in Excel übertragen
    if np.logical_and(0 in holders_data, 0 in holders_data[0]):
        for i in range(0, len(holders_data[0])):
            Sheet2.write(("A"+str(i + 1)), holders_data[1][i])
            Sheet2.write(("B"+str(i + 1)), holders_data[0][i])

    # Überschriften für Top Institutional Holders in Excel einfügen
    Sheet2.write("A5", "Holder")
    Sheet2.write("B5", "Shares")
    Sheet2.write("C5", "Date Reported")
    Sheet2.write("D5", "% Out")
    Sheet2.write("E5", "Value")

    # Institutional Holders in Excel einfügen
    try:
        holders_data_check[1]
    except:
        print("keine Institutional Holders vorhanden")
    else:
        # Überschriften für Top Institutional Holders in Excel einfügen
        Sheet2.write("A5", "Holder")
        Sheet2.write("B5", "Shares")
        Sheet2.write("C5", "Date Reported")
        Sheet2.write("D5", "% Out")
        Sheet2.write("E5", "Value")

        for i in range(0, len(holders_data1["Holder"])):
            Sheet2.write(
                ("A" + str(i + 2 + len(holders_data[0]))), holders_data1["Holder"][i])
            Sheet2.write(
                ("B" + str(i + 2 + len(holders_data[0]))), holders_data1["Shares"][i])
            Sheet2.write(
                ("C" + str(i + 2 + len(holders_data[0]))), holders_data1["Date Reported"][i])
            Sheet2.write(
                ("D" + str(i + 2 + len(holders_data[0]))), holders_data1["% Out"][i])
            Sheet2.write(
                ("E" + str(i + 2 + len(holders_data[0]))), holders_data1["Value"][i])

    # Top Mutual Fund Holders einfügen

    try:
        holders_data_check[2]
    except:
        print("keine Mutual Fund Holders vorhanden")
    else:
        Sheet2.write(
            "A" + str(2 + len(holders_data1["Holder"]) + len(holders_data[0][0])), "Top Institutional Holders: oben Top Mutual Fund Holders: unten")

        for i in range(0, len(holders_data2["Holder"])):
            Sheet2.write(
                ("A" + str(i + 3 + len(holders_data1["Holder"]) + len(holders_data[0][0]))), holders_data2["Holder"][i])
            Sheet2.write(
                ("B" + str(i + 3 + len(holders_data1["Holder"]) + len(holders_data[0][0]))), holders_data2["Shares"][i])
            Sheet2.write(
                ("C" + str(i + 3 + len(holders_data1["Holder"]) + len(holders_data[0][0]))), holders_data2["Date Reported"][i])
            Sheet2.write(
                ("D" + str(i + 3 + len(holders_data1["Holder"]) + len(holders_data[0][0]))), holders_data2["% Out"][i])
            Sheet2.write(
                ("E" + str(i + 3 + len(holders_data1["Holder"]) + len(holders_data[0][0]))), holders_data2["Value"][i])

    if "Breakdown" in dfn:
        for i in range(0, len(dfn[dfn.columns[0]])):
            try:
                dfn[dfn.columns[0]][i]
            except:
                pass
            else:
                Sheet3.write("A" + str(i + 2), dfn[dfn.columns[0]][i])
                Sheet3.write("A1", dfn.columns[0])

            try:
                dfn[dfn.columns[1]][i]
            except:
                pass
            else:
                Sheet3.write("B" + str(i + 2), dfn[dfn.columns[1]][i])
                Sheet3.write("B1", dfn.columns[1])

            try:
                dfn[dfn.columns[2]][i]
            except:
                pass
            else:
                Sheet3.write("C" + str(i + 2), dfn[dfn.columns[2]][i])
                Sheet3.write("C1", dfn.columns[2])

            try:
                dfn[dfn.columns[3]][i]
            except:
                pass
            else:
                Sheet3.write("D" + str(i + 2), dfn[dfn.columns[3]][i])
                Sheet3.write("D1", dfn.columns[3])

            try:
                dfn[dfn.columns[4]][i]
            except:
                pass
            else:
                Sheet3.write("E" + str(i + 2), dfn[dfn.columns[4]][i])
                Sheet3.write("E1", dfn.columns[4])

    if "Breakdown" in dfe:
        for i in range(0, len(dfe[dfe.columns[0]])):
            try:
                dfe[dfe.columns[0]][i]
            except:
                pass
            else:
                Sheet4.write("A" + str(i + 1), dfe[dfe.columns[0]][i])
                Sheet4.write("A1", dfe.columns[0])

            try:
                dfe[dfe.columns[1]][i]
            except:
                pass
            else:
                Sheet4.write("B" + str(i + 1), dfe[dfe.columns[1]][i])
                Sheet4.write("B1", dfe.columns[1])

            try:
                dfe[dfe.columns[2]][i]
            except:
                pass
            else:
                Sheet4.write("C" + str(i + 1), dfe[dfe.columns[2]][i])
                Sheet4.write("C1", dfe.columns[2])

            try:
                dfe[dfe.columns[3]][i]
            except:
                pass
            else:
                Sheet4.write("D" + str(i + 1), dfe[dfe.columns[3]][i])
                Sheet4.write("D1", dfe.columns[3])

            try:
                dfe[dfe.columns[4]][i]
            except:
                pass
            else:
                Sheet4.write("E" + str(i + 1), dfe[dfe.columns[4]][i])
                Sheet4.write("E1", dfe.columns[4])

    if "Breakdown" in dfi:
        for i in range(0, len(dfi[dfi.columns[0]])):
            try:
                dfi[dfi.columns[0]][i]
            except:
                pass
            else:
                Sheet5.write("A" + str(i + 1), dfi[dfi.columns[0]][i])
                Sheet5.write("A1", dfi.columns[0])

            try:
                dfi[dfi.columns[1]][i]
            except:
                pass
            else:
                Sheet5.write("B" + str(i + 1), dfi[dfi.columns[1]][i])
                Sheet5.write("B1", dfi.columns[1])

            try:
                dfi[dfi.columns[2]][i]
            except:
                pass
            else:
                Sheet5.write("C" + str(i + 1), dfi[dfi.columns[2]][i])
                Sheet5.write("C1", dfi.columns[2])

            try:
                dfi[dfi.columns[3]][i]
            except:
                pass
            else:
                Sheet5.write("D" + str(i + 1), dfi[dfi.columns[3]][i])
                Sheet5.write("D1", dfi.columns[3])

            try:
                dfi[dfi.columns[4]][i]
            except:
                pass
            else:
                Sheet5.write("E" + str(i + 1), dfi[dfi.columns[4]][i])
                Sheet5.write("E1", dfi.columns[4])

    tickerWorkbook.close()


for i in range(0, len(tickerArray)):
    execute(tickerArray[i])
    print("Number " + str(i + 1) + " done")
    if i != len(tickerArray)-1:
        time.sleep(90 + random.randrange(0, 10))
