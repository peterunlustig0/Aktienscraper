from random import vonmisesvariate
from h11 import Data
import pandas as pd
import requests
import numpy as np
import os
import xlsxwriter
import math
from bs4 import BeautifulSoup
import random
# Aktien auswählen:

ticker = "AIR.PA"

print(random.randrange(1, 10) + 1)


def getdata(url):
    r = requests.get(url, headers={
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'})
    data = pd.read_html(r.text)
    return data


holders_url = f'https://finance.yahoo.com/quote/{ticker}/holders?p={ticker}'
holders_data = getdata(holders_url)

try:
    holders_data[1]
except:
    pass
else:
    print("yeehaw")
