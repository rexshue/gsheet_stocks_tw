# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import sys
import io
import tempfile
import time

def get_stock_dividend(no):
    url ='https://tw.stock.yahoo.com/d/s/dividend_' + str(no) +'.html'
    req = requests.get(url, timeout=5, allow_redirects=True)
    tmpName = tempfile.gettempdir() + '\\t2.html'
    open(tmpName, 'wb').write(req.content)
    results=[]
    with io.open(tmpName,'r') as f:
        content = f.read()
        soup = BeautifulSoup(content, 'lxml')
        trs = soup.find_all('tr', {'bgcolor':'#FFFFFF'})
        for tr in trs:
            tds = tr.find_all('td')
            result=[]
            for td in tds:
                result.append(float(td.text))
            results.append(result)
    return results

def get_stock_price(no):
    url ='https://tw.stock.yahoo.com/q/ts?s='+str(no)
    req = requests.get(url, timeout=5, allow_redirects=True)
    tmpName = tempfile.gettempdir() + '\\t1.html'
    open(tmpName,'wb').write(req.content)
    result=0.0
    with io.open(tmpName,'r') as f:
        content = f.read()
        soup = BeautifulSoup(content, 'lxml')
        trs = soup.find_all('tr', {'bgcolor':'#ffffff'})
        tds = trs[0].find_all('td')
        if len(tds) > 3 :
            result= float(tds[3].text)
    return result;
#print(get_stock_dividend(2892))
#time.sleep(1)
#print(get_stock_dividend(2834))
#print(get_stock_price(2892))
