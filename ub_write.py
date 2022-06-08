#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 31 11:25:08 2022

@author: weibyapps
"""

import csv
import pygsheets
import time
from itertools import islice
from collections import OrderedDict

start = time.time()
uber_uuid = []
uber_name = []
count = 0

def sheet_connect(url, count): #連接google sheet  count為第幾個試算表
    gc = pygsheets.authorize(service_file='rosy-etching-343806-d6e90a79ae23.json') #金鑰
    #登入試算表
    sht = gc.open_by_url(url)
    wks = sht[count] #第四個試算表
    return wks



#連接到uber_onboaring
uber_url = 'https://docs.google.com/spreadsheets/d/1Zs3VTv7kTyBc1sNESSng8pyV0ihXjQ44hrqTgrU6Spo/edit?pli=1#gid=0'
uber_connect = sheet_connect(uber_url,0)
uber = uber_connect.get_all_records()

#連接到ub申請表
sheet_url = 'https://docs.google.com/spreadsheets/d/1hpZVaLqE_FeLxjcNSiSeKRhznQx9mzZh1b-6n6Tu1gU/edit#gid=1636056984' #uber申請表
uber_apply = sheet_connect(sheet_url,0)
weiby_apply = uber_apply.get_all_records()


###--------------------------------------------------------------------------------------------------------###
#查詢表單內未新增的uuid，格式只能為「https://merchants.ubereats.com/manager/home/uuid」

for uuid in weiby_apply:
    if uuid['status'] == '':
        if uuid['uuid'] != '':
            if 'https://merchants.ubereats.com/manager/home/' in uuid['uuid']:
                if len(uuid['uuid']) == 80:
                    a = uuid['uuid']
                    b = a[-36:]
                    uber_uuid.append(b)
                    uber_name.append(uuid['store_name'])
                else:
                    a = len(uuid['uuid'])-80 #不是80個字元，代表uuid後面還有東西
                    b = uuid['uuid']
                    c = b[-(36+a):] # 先把uuid前面清空
                    d = c[:-a] # 再把uuid後面清空
                    uber_uuid.append(d)
                    uber_name.append(uuid['store_name'])
            else:
                if len(uuid['uuid']) == 36:
                    uber_uuid.append(uuid['uuid'])
                    uber_name.append(uuid['store_name'])

                else:
                    print('申請表內店家：',uuid['store_name'],'的網址->',uuid['uuid'],'格式不符合，請手動新增')
                    
                
uber_uuid = list(OrderedDict.fromkeys(uber_uuid))
#uber_name = list(OrderedDict.fromkeys(uber_name))


#查看是否重複
for uuid in uber: 
    for k,v in zip(uber_uuid,uber_name):
        if k == uuid['Store UUID']:
            uber_uuid.remove(k)
            uber_name.remove(v)
            count -= 1

count = len(uber_uuid)
uber_name = uber_name[:count]


###--------------------------------------------------------------------------------------------------------###
#把新增的資料上傳到Onboarding

total_count = uber_connect.get_values("B","B") #目前總數
new = []

if uber_uuid != []:#把uuid填入onboarding
    for i in range(count):
        uber_connect.cell((len(total_count)+i+1,2)).value = uber_uuid[i] #把lis丟到google sheet
        uber_connect.cell((len(total_count)+i+1,3)).value = uber_name[i]
        uber_connect.cell((len(total_count)+i+1,6)).value = '51ad4c8d-4e2a-4c29-98e4-7ba21dffab5f'
        new.append(uber_name[i])
    print('上傳成功:新增'+str(count)+'筆資料')
    print('Onboarding寫入完成，申請表資料更新中...')
else:
    print('沒有新的資料')
    

###--------------------------------------------------------------------------------------------------------###  
#把ub申請表的資料填完整

ub_tilte = uber_apply.get_values("A","A")
ub_uuid = uber_apply.get_values("N","N")
weiby_count = len(ub_uuid) - len(ub_tilte)

for i in range(weiby_count):
    uber_apply.cell('A'+str(len(ub_uuid)+i-weiby_count+1)).value = len(ub_uuid)- weiby_count +i-2
    uber_apply.cell('B'+str(len(ub_uuid)+i-weiby_count+1)).value = '=if(Q'+str(len(ub_uuid)+i-weiby_count+1)+'="YES", if(T'+str(len(ub_uuid)+i-weiby_count+1)+'>1/1,"已開通","處理中"),"UE處理中")'
    uber_apply.cell('N'+str(len(ub_uuid)+i-weiby_count+1)).value = uber_apply.get_value('N'+str(len(ub_uuid)+i-weiby_count+1))[-36:]
    uber_apply.cell('R'+str(len(ub_uuid)+i-weiby_count+1)).value = "=VLOOKUP(N"+str(len(ub_uuid)+i-weiby_count+1)+",'工作表2'!B:D,3,false)"

###--------------------------------------------------------------------------------------------------------###
#show出更新內容

end = time.time()

if new != []:
    print('新增的店家為：', new)


print('花費:%f秒' %(end-start))
