#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Mar 23 13:36:29 2022

@author: weibyapps
"""
import csv
import pygsheets
import pandas as pd
import time
from itertools import islice
from collections import OrderedDict

with open('/Users/weibyapps/Downloads/店家列表.csv', newline='') as csvfile:
        rows = csv.reader(csvfile)
        email = []
        
        for row in islice(rows,1,None):
            if row[6] == '微碧智慧店面 個人版':
                email.append(row[3])
        #print(email)


def sheet_connect(url, count): #連接google sheet  count為第幾個試算表
    gc = pygsheets.authorize(service_file='rosy-etching-343806-d6e90a79ae23.json')
    #登入試算表
    sht = gc.open_by_url(url)
    wks = sht[count] #第四個試算表
    return wks


sheet_url5 = 'https://docs.google.com/spreadsheets/d/1WzftHeQ1QPx0UmmUHgNxi1_CpsT1B2CjJd4HtqMif_E/edit?pli=1#gid=0&fvid=273619360'#鳳新

gs5 = sheet_connect(sheet_url5, 1)

e = gs5.get_all_records()


weiby_email = []
for item in e:
        for k,v in item.items():
            weiby_email.append(item['weiby_uber_email'])
            weiby_email.append(item['weiby_fp_email'])



weiby_email = list(OrderedDict.fromkeys(weiby_email))
#print(weiby_email)



intersection = [x for x in email for y in weiby_email if x == y]

a = list(set(intersection))
print("個人版店家與串接表比對總數為：",len(a)-1) #-1是因為有一個" "

gs4 = sheet_connect(sheet_url5, 0)

gs4.cell('C7').value = len(a)-1
print('寫入雲端完成')
    
