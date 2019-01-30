# -*- coding: utf-8 -*-

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from myStock import get_stock_dividend
from myStock import get_stock_price
import time

def auth_gss_client(path, scopes):
    credentials = ServiceAccountCredentials.from_json_keyfile_name(path, scopes)
    return gspread.authorize(credentials)

def get_sheet_value(sheet, row, col):
    time.sleep(1.1)
    return sheet.cell(row,col).value

def set_sheet_value(sheet, row, col, value):
    time.sleep(1.1)
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
    time.sleep(1.1)
    sheet.update_cells(cell_list)

def update_divident(no,rows):
    results = get_stock_dividend(no)
    colIndex = 14
    cell_value = []
    for result in results:
        cell_value.append(result[1])
        cell_value.append(result[2])
        #set_sheet_value(wks.sheet1,lineIndex, colIndex, result[1])
        #colIndex += 1
        #set_sheet_value(wks.sheet1,lineIndex, colIndex, result[2])
        #colIndex += 1
    set_sheet_values(wks.sheet1,rows,colIndex, cell_value)

auth_json_path = 'auth.json'
gss_scopes = ['https://spreadsheets.google.com/feeds']

gss_client = auth_gss_client(auth_json_path, gss_scopes)
s_key=open('spreadsheet_key','r').read()

wks = gss_client.open_by_key(s_key)
#print(get_sheet_value(wks.sheet1, 2, 1))
stockNo = 1
lineIndex = 2
run = True
while run:
    stockNo = get_sheet_value(wks.sheet1, lineIndex ,1)
    if stockNo != '':
        print('Get Stock '+str(stockNo)+'\n')
        set_sheet_value(wks.sheet1, lineIndex, 3, get_stock_price(stockNo))
        divValue = get_sheet_value(wks.sheet1, lineIndex, 14)
        if divValue == '':
            update_divident(stockNo, lineIndex)
        time.sleep(1)
        lineIndex += 1
    else:
        print('Bye!\n')
        run = False
