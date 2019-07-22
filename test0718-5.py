import requests, json
from bs4 import BeautifulSoup
import re
import os,sys
from urllib import request
import pandas as pd
import urllib.request
import csv
import xlrd
import time

process_record="Start:"+time.strftime('%X %x %Z')
#print("Now:",time.strftime('%X %x %Z'))

##Initial Data
co_id_1=""
SDATE=""
EDATE=""
YEAR1=""
YEAR2=""
MONTH1=""
MONTH2=""
SDAY=""
EDAY=""
sort=""
rpt=""

##Read Query Data from Config Excel
ConFilename ='C:/Users/User/Desktop/查詢設定.xls'   ##'/路徑/檔名.xlsx'
book = xlrd.open_workbook(ConFilename)
table = book.sheet_by_name("config") #.sheet_by_index(0)
nrows = table.nrows
ncols = table.ncols

##Get Data from Excel         
for i in range(0,nrows):
    for j in range(1,2):
        if table.cell(i,j).value:
            if table.cell(i,j).value=="sort":            
                sort=int(table.cell(i,j+1).value)
            ##elif table.cell(i,j).value=="co_id_1":
              ##  co_id_1=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="SDATE":
                SDATE=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="EDATE":
                EDATE=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="YEAR1":
                YEAR1=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="YEAR2":
                YEAR2=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="MONTH1":
                MONTH1=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="MONTH2":
                MONTH2=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="SDAY":
                SDAY=int(table.cell(i,j+1).value)
            elif table.cell(i,j).value=="EDAY":
                EDAY=int(table.cell(i,j+1).value)

tablel = book.sheet_by_name("AnnList")
nrows = tablel.nrows
ncols = tablel.ncols
AnnList=[]
AnnListName=[]
for k in range (1,nrows):
 for l in range(2,3):
  if tablel.cell(k,l).value :
    #print("row:",k,",col:",l,",value:",tablel.cell(k,l).value)       
    AnnList.append(tablel.cell(k,l).value)
    AnnListName.append(tablel.cell(k,l-1).value)

tablec = book.sheet_by_name("company")
nrows = tablec.nrows
ncols = tablec.ncols
CompanyList=[]
for i in range(1,nrows):
    for j in range(0,1):
        if tablec.cell(i,j).value :           
           CompanyList.append(int(tablec.cell(i,j).value))

#print(len(AnnList))
idx=0
with open('E:/output.csv', 'w', newline='', encoding='utf-8') as csvfile:
 parm1=""
 writer = csv.writer(csvfile, delimiter=' ')
 for x in range(0,len(AnnList)):
    print("公告種類:",AnnListName[x],",rpt:",AnnList[x])
    rpt=AnnList[x]
    for y in range(0,len(CompanyList)):
      print("company:",CompanyList[y])
      idx+=1
      r = requests.Session()
      payload ={
               "encodeURIComponent":"1",
               "step":"2",
               "firstin":"1",
               "TYPEK":"all",
               "co_id_1":CompanyList[y],#公司代號
               "sort":sort,
               "scope":"1",    
               "SDATE": SDATE,
               "EDATE": EDATE,
               "YEAR1": YEAR1,
               "YEAR2": YEAR2,
               "MONTH1": MONTH1,
               "MONTH2": MONTH2,
               "SDAY": SDAY,
               "EDAY": EDAY,
               "rpt": rpt #公告種類
               }
      r2 = r.post("https://mops.twse.com.tw/mops/web/ajax_t146sb10",payload)
      soup = BeautifulSoup(r2.text,"html.parser") #將網頁資料以html.parser
      #print(soup)
      ##-------------------------------------------------------------           
      #dc=str(soup).find("查無所需資料")
      #print("查無所需資料 result:",dc)
      #if dc==-1:
      #   print('got data')           
      t1=soup.find_all("table")
      for i in t1:   
                 #print("i-class:",i.get('class'))  
                 if str(i['class'])=="['noBorder']":
                     #print('is noBorder')
                     ##取得公告名稱
                     h0=i.select("b")
                     for s in h0:
                      print('公告名稱:',s.text)
                     
                     parm1="'公告名稱','"+s.text+"'"
                     writer.writerow(parm1)
                     parm1=""
                 else:
                     #print('not noBorder')      
                     d=i.find_all("tr")
                     for d1 in d:
                        #print("2:",d1.get('class'))
                        if str(d1.get('class'))=="['tblHead']":
                            #print("is tblHead")
                            d2=d1.find_all("th")
                            for d3 in d2:
                             #print("標題:",d3.text);
                             if parm1=="": parm1=d3.text
                             else:parm1=parm1+"','"+d3.text
                            parm1="'"+parm1+"'"
                            print(parm1)
                            parm1=parm1.replace('\xa0', '')
                            writer.writerow([parm1])
                            parm1=""
                        else:
                             #print("not tblHead")                             
                             d2=d1.find_all("td")
                             for d3 in d2:
                               #print("內容:",d3)
                               elem = d3.find('input')
                               #print("find input elem:",elem)                               
                               #t = re.search('<td.*?>(.*?)</td>',str(d3))                               
                               #print("t:",t)                               
                               #switch = re.search('<input onclick=\'(.*?)";openWindow',str(t.group(1)))
                               #switch = re.search('<input onclick=\'(.*?)";openWindow',str(elem))
                               #print("switch:",switch)
                               if (elem) :
                                 #print("Got Input:", str(elem)) 
                                 data2 = re.search(';action="(.*?)";openWindow',str(elem))
                                 action=""
                                 ##有些公告的詳細資料連結沒有action attribute
                                 if data2:
                                     #print("action:",data2.group(1))
                                     action=data2.group(1)
                                 #data3 = re.findall('document.fm_t59sb08.(.*?).value="(.*?)";',str(switch.group(1)))
                                 ##有些公告用正則表示式無法正確切出資料                                     
                                 #data3 = re.findall('document.fm_t59sb08.(.*?).value="(.*?)";',str(elem))
                                 rpt1=str(rpt).replace('bool_','')    
                                 match_str="document.fm_"+rpt1+".(.*?).value=\"(.*?)\";"
                                 #print("match:",match_str)
                                 data3 = re.findall(match_str,str(elem))
                                 i=0
                                 parm="?firstin=true"
                                 for x in data3:
                                   #print("data3:",str(data3[i][0]),"||",str(data3[i][1]))
                                   parm+='&'+str(data3[i][0])+'='+str(data3[i][1])
                                   i=i+1;
                                 result='https://mops.twse.com.tw'+str(action)+str(parm)
                                 print("詳細資料url:",result)
                                 parm1="'"+parm1+"','"+result+"'"
                                 print(parm1)
                                 parm1=parm1.replace('\xa0', '')
                                 parm1=parm1.replace('<br/>','')
                                 writer.writerow([parm1])                 
                                 parm1=""
                               else:
                                 #print(t.group(1))
                                 #print("Not Input data:",d3)
                                 t = re.search('<td.*?>(.*?)</td>',str(d3)) 
                                 if parm1=="":
                                   parm1=str(t.group(1))
                                 else:
                                   parm1=parm1+"','"+str(t.group(1))
      #else:
        #print('no data')
        
      ##-------------------------------------------------------------
      r.close()
      time.sleep(3) 

process_record=process_record+",End:"+time.strftime('%X %x %Z')
print(process_record)                
