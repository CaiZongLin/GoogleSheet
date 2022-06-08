#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Mar 14 13:18:06 2022

@author: weibyapps
"""

import csv
import pygsheets
import time
from itertools import islice



def sheet_connect(url, count): #連接google sheet  count為第幾個試算表
    gc = pygsheets.authorize(service_file='rosy-etching-343806-d6e90a79ae23.json') #金鑰
    #登入試算表
    sht = gc.open_by_url(url)
    wks = sht[count] #第四個試算表
    return wks
"""
def get_key (dict, value):
    return [k for k, v in dict.items() if v == value]
"""


start = time.time()
count = 0  #計算有幾筆資料
weiby_number = 0 #weiby有幾筆

foodpanda_uuid = []
foodpanda_name = []

#辰峰
sheet_url1 = 'https://docs.google.com/spreadsheets/d/1CziJW5TRqSbskmQirEuxmetrSCwIW9BO1U-pcuTa7Ks/edit?resourcekey#gid=793693489' 
WIS_Changfon = sheet_connect(sheet_url1, 0)
Changfon = WIS_Changfon.get_all_records()

#中華麟
sheet_url2 = 'https://docs.google.com/spreadsheets/d/1oURWodRMekGSG-IcViHWbWfxHQD_3IXDcDFGd1C3IQ8/edit#gid=793693489' 
WIS_Hlmcoltd = sheet_connect(sheet_url2, 0) 
Hlmcoltd = WIS_Hlmcoltd.get_all_records()

#文泓
sheet_url3 = 'https://docs.google.com/spreadsheets/d/1qw-TxfC-XRjz4z6OK9BG7fxaiZgOZaEg-VKWpyN7RQo/edit#gid=793693489'
WIS_WenHung = sheet_connect(sheet_url3, 0) 
WenHung = WIS_WenHung.get_all_records()

#葆光
sheet_url4 = 'https://docs.google.com/spreadsheets/d/1T8jutltgOnrJcufovnkmmP2Vp6pbSIb67MQrn2H8hn0/edit#gid=793693489'
WIS_FitSoft = sheet_connect(sheet_url4, 0) 
FitSoft = WIS_FitSoft.get_all_records()

#鳳新
sheet_url5 = 'https://docs.google.com/spreadsheets/d/1qGyok7EPc8eezGDut0KHG26AV29TW8E7rEf5sbU8HhM/edit#gid=793693489'
WIS_POS365 = sheet_connect(sheet_url5, 0) 
POS365= WIS_POS365.get_all_records()


#fp申請表
sheet_url6 = 'https://docs.google.com/spreadsheets/d/1ekbUbg_ggXi1Max7XZz0cxmKmO_mmP_p02ozSw9tMPo/edit#gid=2130903407'
fp_apply = sheet_connect(sheet_url6, 1) 
weiby_fp = fp_apply.get_all_records()


#連接到foodpanda_onboaring
foodpanda_url = 'https://docs.google.com/spreadsheets/d/1YRco6c29_LKyprVY4THAtLTbOSh1bd_SF8XY60kiVr0/edit#gid=1616046752'
fp_connect = sheet_connect(foodpanda_url,1)
fp = fp_connect.get_all_records()



def catch(WIS): #取得WIS資料
    global count
    global foodpanda_uuid 
    global foodpanda_name
    for item in WIS:  #取得試算表資料，for迴圈跑資料
        for k,v in item.items():
            if item[k] == '#N/A':
                foodpanda_uuid.append(item['2. foodpanda 店家代碼'])
                foodpanda_name.append(item['1. foodpanda店名'])
                count += 1
    return foodpanda_uuid,foodpanda_name
                
def catch_weiby(weiby): #weiby的foodpanda uuid
    global count
    global weiby_number
    global foodpanda_uuid 
    global foodpanda_name
    for item in weiby:
        if item['自動生成'] != item['手動生成']:
            foodpanda_uuid.append(item['自動生成'])
            foodpanda_name.append(item['店家名稱'])
            count += 1
            weiby_number += 1
    return foodpanda_uuid,foodpanda_name



catch(Changfon)
catch(Hlmcoltd)  
catch(WenHung)
catch(FitSoft)
catch(POS365)

catch_weiby(weiby_fp)

###--------------------------------------------------------------------------------------------------------###
# 欲新增的uuid 與foodpanda_onboarding 比對是否重複

for uuid in fp: 
    for k,v in zip(foodpanda_uuid,foodpanda_name):
        if k == uuid['python']:
            foodpanda_uuid.remove(k)
            foodpanda_name.remove(v)
            count -= 1

###--------------------------------------------------------------------------------------------------------###
#把新增的資料上傳到Onboarding

total_count = fp_connect.get_values("B","B") #目前總數
new = []

if foodpanda_uuid != []:
    for i in range(count):
        fp_connect.cell((len(total_count)+i+1,1)).value = foodpanda_uuid[i] #把lis丟到google sheet
        fp_connect.cell((len(total_count)+i+1,2)).value = foodpanda_uuid[i]
        fp_connect.cell((len(total_count)+i+1,3)).value = foodpanda_name[i]
        fp_connect.cell((len(total_count)+i+1,7)).value = 'True'
        new.append(foodpanda_name[i])
    print('上傳成功:新增'+str(count)+'筆資料')
    print('Onboarding寫入完成，申請表資料更新中...')
else:
    print('沒有新的資料')

###--------------------------------------------------------------------------------------------------------###
#申請表內資料更新

fp_uuid_Automatic = fp_apply.get_values("K","K")
fp_uuid_key = fp_apply.get_values("A","A")
if len(fp_uuid_Automatic) != len(fp_uuid_key):
    weiby_count = len(fp_uuid_Automatic) - len(fp_uuid_key)
    for i in range(weiby_count):
        fp_apply.cell('A'+str(len(fp_uuid_Automatic)+i-weiby_count+1)).value = len(fp_uuid_Automatic)- weiby_count +i-2
        fp_apply.cell('B'+str(len(fp_uuid_Automatic)+i-weiby_count+1)).value = '=if(T'+str(len(fp_uuid_Automatic)+i-weiby_count+1)+'="YES", if(V'+str(len(fp_uuid_Automatic)+i-weiby_count+1)+'>1/1,"已完成","處理中"),"FP處理中")'
        fp_apply.cell('L'+str(len(fp_uuid_Automatic)+i-weiby_count+1)).value = fp_apply.get_value('K'+str(len(fp_uuid_Automatic)+i-weiby_count+1))
        fp_apply.cell('T'+str(len(fp_uuid_Automatic)+i-weiby_count+1)).value = "=vlookup(L"+str(len(fp_uuid_Automatic)+i-weiby_count+1)+",'工作表4'!B:C,2,false)"
    print('申請表更新完畢')
    
###--------------------------------------------------------------------------------------------------------###
#show出更新內容

if new != []:
    print('新增的店家為：', new)
    if weiby_number-count != 0:
        print('WIS有',abs(weiby_number-count),'間')
    if weiby_number != 0:
        print('Weiby有',abs(weiby_number),'間')
    
end = time.time()

print('花費:%f秒' %(end-start))

