#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Apr  4 10:50:12 2022

@author: weibyapps
"""

import csv
import pygsheets
import time
from itertools import islice
from collections import OrderedDict


def sheet_connect(url, count): #連接google sheet  count為第幾個試算表
    gc = pygsheets.authorize(service_file='rosy-etching-343806-d6e90a79ae23.json') #金鑰
    #登入試算表
    sht = gc.open_by_url(url)
    wks = sht[count] #第四個試算表
    return wks

#連接到拉亞申請表
laya_url = 'https://docs.google.com/spreadsheets/d/1vsq1950gjGVTmHm2pmg_hpaxBEJkh5-tTL54odzNSVU/edit#gid=865636093'
laya_connect = sheet_connect(laya_url,0)
laya = laya_connect.get_all_records()
laya_ub = laya_connect.get_values("N5","N")



#連到Uber Onboarding表
uber_url = 'https://docs.google.com/spreadsheets/d/1Zs3VTv7kTyBc1sNESSng8pyV0ihXjQ44hrqTgrU6Spo/edit?pli=1#gid=0'
uber_connect = sheet_connect(uber_url,0)
uber = uber_connect.get_all_records()

#連接到foodpanda_onboaring
foodpanda_url = 'https://docs.google.com/spreadsheets/d/1YRco6c29_LKyprVY4THAtLTbOSh1bd_SF8XY60kiVr0/edit#gid=1616046752'
fp_connect = sheet_connect(foodpanda_url,1)
fp = fp_connect.get_all_records()

ub_uuid = []
ub_name = []
fp_uuid = []
fp_name = []

###--------------------------------------------------------------------------------------------------------###
#讀取拉亞申請表的uuid資料

print('正在讀取拉亞資料中，請稍等...')

for i in laya:
    if i['聯合信用卡中心'] == 'Uber Eats外送串接服務（注意：需已完成外送平台帳號申請）':
        if i['Uber Eats'] != "":
            if 'https://merchants.ubereats.com/manager/home/' in i['Uber Eats']:
                if len(i['Uber Eats']) == 80:
                    uuid = i['Uber Eats']
                    ub_uuid.append(uuid[-36:])
                    ub_name.append(i['分店名稱'])
                else:
                    a = len(i['Uber Eats'])-80 #不是80個字元，代表uuid後面還有東西
                    b = i['Uber Eats']
                    c = b[-(36+a):] # 先把uuid前面清空
                    d = c[:-a] # 再把uuid後面清空
                    ub_uuid.append(d)
                    ub_name.append(i['分店名稱'])
            elif len(i['Uber Eats']) == 36:
                    continue
            elif 'https://merchant.ubereats.com/manager/home/' in i['Uber Eats']:
                if len(i['Uber Eats']) == 79:
                    uuid = i['Uber Eats']
                    ub_uuid.append(uuid[-36:])
                    ub_name.append(i['分店名稱'])
                else:
                    a = len(i['Uber Eats'])-79 #不是80個字元，代表uuid後面還有東西
                    b = i['Uber Eats']
                    c = b[-(36+a):] # 先把uuid前面清空
                    d = c[:-a] # 再把uuid後面清空
                    ub_uuid.append(d)
                    ub_name.append(i['分店名稱'])
            else:
                print('*****申請表內店家：',i['分店名稱'],' 的網址格式不符合，請手動/檢查新增*****\n')
            
        if i['foodpanda'] != "":
            fp_uuid.append(i['foodpanda'])
            fp_name.append(i['分店名稱'])
    elif i['聯合信用卡中心'] == 'foodpanda外送串接服務（注意：需已完成外送平台帳號申請）':
        if i['Uber Eats'] != "":
            if 'https://merchants.ubereats.com/manager/home/' in i['Uber Eats']:
                if len(i['Uber Eats']) == 80:
                    uuid = i['Uber Eats']
                    ub_uuid.append(uuid[-36:])
                    ub_name.append(i['分店名稱'])
                else:
                    a = len(i['Uber Eats'])-80 #不是80個字元，代表uuid後面還有東西
                    b = i['Uber Eats']
                    c = b[-(36+a):] # 先把uuid前面清空
                    d = c[:-a] # 再把uuid後面清空
                    ub_uuid.append(d)
                    ub_name.append(i['分店名稱'])
            elif len(i['Uber Eats']) == 36:
                    continue
            else:
                print('*****申請表內店家：',i['分店名稱'],' 的網址格式不符合，請手動/檢查新增*****\n')
        if i['foodpanda'] != "":
            fp_uuid.append(i['foodpanda'])
            fp_name.append(i['分店名稱'])

###--------------------------------------------------------------------------------------------------------###
#寫入uber
print('正在查詢與Uber Onboarding資料是否有不同...')

#查看ub是否重複
for uuid in uber: #跑uber_Onboarding
    for k,v in zip(ub_uuid,ub_name): #跑laya_ub
        if k == uuid['Store UUID']:
            ub_uuid.remove(k)
            ub_name.remove(v)

#檢查列表內是否有重複
ub_uuid = list(OrderedDict.fromkeys(ub_uuid))
ub_name = list(OrderedDict.fromkeys(ub_name))


ub_count = len(ub_uuid)
ub_new = []
ub_total_count = uber_connect.get_values("B2","B") #Onboarding 目前總數

if ub_uuid != []:#把uuid填入onboarding
    for i in range(ub_count):
        uber_connect.cell((len(ub_total_count)+i+1,2)).value = ub_uuid[i] #把lis丟到google sheet
        uber_connect.cell((len(ub_total_count)+i+1,3)).value = ub_name[i]
        uber_connect.cell((len(ub_total_count)+i+1,6)).value = '51ad4c8d-4e2a-4c29-98e4-7ba21dffab5f'
        ub_new.append(ub_name[i])
    print('Uber新增成功:新增'+str(ub_count)+'筆資料')


if ub_new != []:
    print('Uber寫入完成，查詢熊貓中...\n')
else:
    print('Uber沒有新資料，查詢熊貓中...\n')
    
    
###--------------------------------------------------------------------------------------------------------###
print('正在查詢與Foodpanda Onboarding資料是否有不同...')
#寫入foodpanda

#查看ub是否重複
for uuid in fp: #跑uber_Onboarding
    for k,v in zip(fp_uuid,fp_name): #跑laya_ub
        if k == uuid['python']:
            fp_uuid.remove(k)
            fp_name.remove(v)
            
#檢查列表內是否有重複
fp_uuid = list(OrderedDict.fromkeys(fp_uuid))
fp_name = list(OrderedDict.fromkeys(fp_name))


fp_count = len(fp_uuid)
fp_new = []
fp_total_count = fp_connect.get_values("B2","B") #Onboarding 目前總數


if fp_uuid != []:
    for i in range(fp_count):
        fp_connect.cell((len(fp_total_count)+i+1,1)).value = fp_uuid[i] #把lis丟到google sheet
        fp_connect.cell((len(fp_total_count)+i+1,2)).value = fp_uuid[i]
        fp_connect.cell((len(fp_total_count)+i+1,3)).value = fp_name[i]
        fp_connect.cell((len(fp_total_count)+i+1,7)).value = 'True'
        fp_new.append(fp_name[i])
    print('Foopdanda新增成功:新增'+str(fp_count)+'筆資料')

    
if fp_new != []:
    print('Foodpanda寫入成功，寫入步驟完成\n')
else:
    print('Foodpanda沒有新資料，寫入步驟完成\n')

###--------------------------------------------------------------------------------------------------------###
#寫入分析表
if fp_new or ub_new != []:
    print('正在寫入分析表中...')
    laya_Analysis = sheet_connect(laya_url,2)
    total_count = len(laya_Analysis.get_values("A","A"))
    Analysis_list = fp_uuid+ub_uuid
    Analysis_name = fp_name+ub_name
    
    new_count = len(Analysis_list)
    
    
    for i in range(new_count):
        laya_Analysis.cell('A'+str(total_count+i+1)).value = Analysis_name[i]
        if len(Analysis_list[i]) > 4:
            laya_Analysis.cell('B'+str(total_count+i+1)).value = Analysis_list[i]
        else:
            laya_Analysis.cell('C'+str(total_count+i+1)).value = Analysis_list[i]
        laya_Analysis.cell('J'+str(total_count+i+1)).value = '=ifs(B'+str(total_count+i+1)+'>1/1,C'+str(total_count+i+1)+'>1/1,true,false)'
    
    print('分析表寫入完成')


###--------------------------------------------------------------------------------------------------------###
#show出更新內容

if ub_new != []:
    print('Uber新增的店家為：', ub_new)
else:
    print('Uber沒有新增的店家')

if fp_new != []:
    print('Foodpanda新增的店家為：', fp_new)
else:
    print('Foodpanda沒有新增的店家')



