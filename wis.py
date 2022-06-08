"""自動讀檔寫入google sheet"""
import csv
import pygsheets
import pandas as pd
import time
from itertools import islice

time = time.localtime(time.time())


def sheet_connect(): #連接google sheet
    gc = pygsheets.authorize(service_file='rosy-etching-343806-d6e90a79ae23.json')
    #登入試算表
    sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1WzftHeQ1QPx0UmmUHgNxi1_CpsT1B2CjJd4HtqMif_E/edit?pli=1#gid=712968489')
    wks = sht[3] #第四個試算表
    return wks

gs = sheet_connect() 

def Hlmcoltd(): #讀取檔案、寫入google sheet
    # 開啟 CSV 檔案
    if time.tm_mon <10:
        csv_file = 'Hlmcoltd 中華麟股份有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    else:
        csv_file = 'Hlmcoltd 中華麟股份有限公司{}年{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    
    
    with open('/Users/weibyapps/Downloads/'+csv_file, newline='') as csvfile: #檔案路徑有誤請修改這邊 把檔案丟到右邊terminal就會有位置，路徑把檔名去掉
        rows = csv.reader(csvfile)
        UberEats = []
        foodpanda = []
        
        for row in islice(rows,1,None): #把csv檔案寫入list
            if row[2] == 'UberEats':
                UberEats.append(row[1])
            elif row[2] == 'foodpanda':
                foodpanda.append(row[1])
    csvfile.close()
    
    if (UberEats,foodpanda) != []:
        for i in range(len(UberEats)): 
            gs.cell((i+4,1)).value = UberEats[i] #把lis丟到google sheet
        for i in range(len(foodpanda)):
            gs.cell((i+4,2)).value = foodpanda[i]
        print('中華麟寫入成功')
    else:
        print('匯入失敗，請檢查檔案名稱是否正確')

def Changfon():
    # 開啟 CSV 檔案
    if time.tm_mon <10:
        csv_file = 'Changfon 辰峰科技有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    else:
        csv_file = 'Changfon 辰峰科技有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    
    
    with open('/Users/weibyapps/Downloads/'+csv_file, newline='') as csvfile:
        rows = csv.reader(csvfile)
        UberEats = []
        foodpanda = []
        
        for row in islice(rows,1,None):
            if row[2] == 'UberEats':
                UberEats.append(row[1])
            elif row[2] == 'foodpanda':
                foodpanda.append(row[1])
    csvfile.close()
    
    if (UberEats,foodpanda) != []:
        for i in range(len(UberEats)):
            gs.cell((i+4,4)).value = UberEats[i]
        for i in range(len(foodpanda)):
            gs.cell((i+4,5)).value = foodpanda[i]
        print('辰峰寫入成功')
    else:
        print('匯入失敗，請檢查檔案名稱是否正確')

def WinHong():
    # 開啟 CSV 檔案
    if time.tm_mon <10:
        csv_file = '文泓資訊有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    else:
        csv_file = '文泓資訊有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    
    
    with open('/Users/weibyapps/Downloads/'+csv_file, newline='') as csvfile:
        rows = csv.reader(csvfile)
        UberEats = []
        foodpanda = []
        
        for row in islice(rows,1,None):
            if row[2] == 'UberEats':
                UberEats.append(row[1])
            elif row[2] == 'foodpanda':
                foodpanda.append(row[1])
    csvfile.close()
    
    if (UberEats,foodpanda) != []:
        for i in range(len(UberEats)):
            gs.cell((i+4,10)).value = UberEats[i]
        for i in range(len(foodpanda)):
            gs.cell((i+4,11)).value = foodpanda[i]
        print('文泓寫入成功')
    else:
        print('匯入失敗，請檢查檔案名稱是否正確')

def FitSoft():
    # 開啟 CSV 檔案
    if time.tm_mon <10:
        csv_file = 'FitSoft 葆光系統股份有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    else:
        csv_file = 'FitSoft 葆光系統股份有限公司{}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    
    
    with open('/Users/weibyapps/Downloads/'+csv_file, newline='') as csvfile:
        rows = csv.reader(csvfile)
        UberEats = []
        foodpanda = []
        
        for row in islice(rows,1,None):
            if row[2] == 'UberEats':
                UberEats.append(row[1])
            elif row[2] == 'foodpanda':
                foodpanda.append(row[1])
    csvfile.close()
    
    if (UberEats,foodpanda) != []:
        for i in range(len(UberEats)):
            gs.cell((i+4,13)).value = UberEats[i]
        for i in range(len(foodpanda)):
            gs.cell((i+4,14)).value = foodpanda[i]
        print('葆光寫入成功')
    else:
        print('匯入失敗，請檢查檔案名稱是否正確')

def POS365():
    # 開啟 CSV 檔案
    if time.tm_mon <10:
        csv_file = 'POS365 鳳新電腦有限公司 {}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    else:
        csv_file = 'POS365 鳳新電腦有限公司 {}年0{}月份對帳單.csv'.format(time.tm_year,time.tm_mon)
    
    
    with open('/Users/weibyapps/Downloads/'+csv_file, newline='') as csvfile:
        rows = csv.reader(csvfile)
        UberEats = []
        foodpanda = []
        
        for row in islice(rows,1,None):
            if row[2] == 'UberEats':
                UberEats.append(row[1])
            elif row[2] == 'foodpanda':
                foodpanda.append(row[1])
    csvfile.close()
    
    if (UberEats,foodpanda) != []:
        for i in range(len(UberEats)):
            gs.cell((i+4,7)).value = UberEats[i]
        for i in range(len(foodpanda)):
            gs.cell((i+4,8)).value = foodpanda[i]
        print('鳳新寫入成功')
    else:
        print('匯入失敗，請檢查檔案名稱是否正確')

Hlmcoltd()
Changfon()
POS365()
WinHong()
FitSoft()


