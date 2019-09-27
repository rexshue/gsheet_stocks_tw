# -*- coding: utf-8 -*-
from bs4 import BeautifulSoup
import requests
import sys
import io
import tempfile
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import sys
import os

#取得股利資訊
def get_stock_dividend(no):
    #remove old file
    tmpName = tempfile.gettempdir() + '\\t2.html'
    if os.path.exists(tmpName):
        os.remove(tmpName)
    #
    url ='https://tw.stock.yahoo.com/d/s/dividend_' + str(no) +'.html'
    req = requests.get(url, timeout=5, allow_redirects=True)
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

#取得股價
def get_stock_price(no):
    #remove old file
    tmpName = tempfile.gettempdir() + '\\t'+ no + '.html'
    if os.path.exists(tmpName):
        os.remove(tmpName)
    #
    url ='https://tw.stock.yahoo.com/q/ts?s='+str(no)
    req = requests.get(url, timeout=5, allow_redirects=True)
    open(tmpName,'wb').write(req.content)
    result=0.0
    with io.open(tmpName,'r') as f:
        content = f.read()
        soup = BeautifulSoup(content, 'lxml')
        trs = soup.find_all('tr', {'bgcolor':'#ffffff'})
        if len(trs) > 0:
            tds = trs[0].find_all('td')
            if len(tds) > 3 :
                result= float(tds[3].text)
            else:
                sys.stdout.write('v')
        else:
            sys.stdout.write('z')
    return result;

def auth_gss_client(path, scopes):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scopes)
    return gspread.authorize(credentials)

def get_sheet_value(sheet, row, col):
    time.sleep(2.1)
    return sheet.cell(row,col).value

def set_sheet_value(sheet, row, col, value):
    time.sleep(2.1)
    sheet.update_cell(row,col,value)

def set_sheet_values(sheet, row, col, values):
    alpha='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    preA=alpha[col-1]+str(row)
    m=col+len(values)-2
    j=m/len(alpha)
    k=m%len(alpha)
    if j>0:
        preB=alpha[j-1]+alpha[k]+str(row)
    else:
        preB=alpha[col+len(values)-2]+str(row)
    label=preA+':'+preB
    cell_list=sheet.range(label)
    for c,v in zip(cell_list, values):
        c.value = v
    time.sleep(2.1)
    sheet.update_cells(cell_list)

def update_divident(sheet,no,rows):
    results = get_stock_dividend(no)
    colIndex = 14
    cell_value = []
    for result in results:
        cell_value.append(result[1])
        cell_value.append(result[2])
    tmpVal = get_sheet_value(sheet, rows ,colIndex)
    if((len(cell_value) > 0) and (tmpVal != cell_value[0])):
        set_sheet_values(sheet,rows,colIndex, cell_value)
    else:
        sys.stdout.write('x')

def main():
    auth_json_path = 'auth.json'
    gss_scopes = ['https://spreadsheets.google.com/feeds']
    gss_client = auth_gss_client(auth_json_path, gss_scopes)
    s_key=open('spreadsheet_key','r').read()
    wks = gss_client.open_by_key(s_key)
    stockNo = 1
    lineIndex = 2
    run = True
    print('Get Price\n')
    while run:
        stockNo = get_sheet_value(wks.sheet1, lineIndex ,1)
        if stockNo != '':
            sys.stdout.write(str(stockNo))
            #set_sheet_value(wks.sheet1, lineIndex, 3, get_stock_price(stockNo))
            #如果是空值，則更新數值
            #divValue = get_sheet_value(wks.sheet1, lineIndex, 14)
            #if divValue == '':
            update_divident(wks.sheet1,stockNo, lineIndex)
            sys.stdout.write(',')
            #連發
            time.sleep(3)
            lineIndex += 1
        else:
            print('Bye!\n')
            run = False

if __name__ == '__main__':
    main()
