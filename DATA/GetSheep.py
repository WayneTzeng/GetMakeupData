import requests 
import json 
import io 
import os 
import time 
import datetime 
import pyodbc 
import concurrent.futures 
import openpyxl 
import pandas
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

class product :
    pass
ProductInfo = product()


def DoNo(Dono):
    if Dono > 11 :
        return 11
    else : 
        return Dono

Key = openpyxl.load_workbook('ShopDataTT.xlsx')
Sheet = Key["userdataKey"]
rows = Sheet.rows

for row in list(rows):  # 遍歷每行資料
    caseOri = []   # 用於存放一行資料
    for c in row:  # 把每行的每個單元格的值取出來，存放到case裡
        caseOri.append(c.value)
    print(caseOri[0])

    if caseOri[0] != "NO":
        keyword = str(caseOri[0])

        header = {
            "Accept":	"*/*",
            "Accept-Encoding":	"gzip, deflate, br",
            "Accept-Language":	"zh-TW,zh;q=0.8,en-US;q=0.5,en;q=0.3",
            "Connection":	"keep-alive",
            #"Cookie":'SPC_IA=-1; SPC_EC=BHVxYmwLtA91/UeNs0CbG6nnuSFci9Y+Eu1rRSrQRORfk0iCWvocaqiwKvm5Gn6jGSzrHpWtklA9Nri+bQ5Jo70OVde1oxMtvQa4W+mdVUA1aoJhaJ2r8j0DWrjS/Xp/iNSs2DuJ+WUKdwBYRbf7tdsDYApKKlxnLDmodZ7neYY=; SPC_F=VZuFrwdYDEgnc2r5B93IWKHyn6psXNbR; REC_T_ID=d69554fa-919f-11e9-925d-f8f21e1a8170; SPC_T_ID="inP/JmymYNR1fUIEbiGPKvubpVvzcxCXn6c+nadGJGYoEQK5LHf6G5dfz6J/B4zHFWo3FKCYOQcndLATOFlxUKy378jNEsmlC7P6I2DDXe0="; SPC_U=180232515; SPC_T_IV="bS6TMawNwqxCzHGES2FUag=="; _ga=GA1.2.694026113.1560845138; __BWfp=c1560845152774x3b44…g=="; _ga_8XD7H14FP3=GS1.1.1616864903.26.0.1616864903.0; REC_T_ID=dc39c2cc-2ffe-11eb-b0c6-d0946657b48c; SC_DFP=xAavbumzzuGSC1GM4jpBki6FbW9geLv6; _ga_RPSBE3TQZZ=GS1.1.1616923286.17.1.1616924079.56; _gcl_au=1.1.2138296484.1614775875; _med=refer; csrftoken=l139H6fHMotQ0grTC0Wa5QveGTOEPysl; SPC_SI=bfftoctw1.h0sWHdR806I8LWgReGgsOpkUWxcFJaYX; welcomePkgShown=true; AMP_TOKEN=%24NOT_FOUND; _gid=GA1.2.251293147.1616923287; G_ENABLED_IDPS=google; SPC_CLIENTID=Vlp1RnJ3ZFlERWdudcwkfnnstsgujacm; _dc_gtm_UA-61915057-6=1',
            "Host":	"shopee.tw",
            "If-None-Match-":	"55b03-ee368bcf86f47c70f7592412a74b7ec0",
            "Referer":	"https://shopee.tw/search?keyword="+keyword+"&locations=-1&noCorrection=true&page=0",
            "TE":	"Trailers",
            "User-Agent":	"Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:87.0) Gecko/20100101 Firefox/87.0",
            "X-API-SOURCE":	"pc",
            "X-Requested-With":	"XMLHttpRequest",
            "X-Shopee-Language":	"zh-Hant",
        }


        URL = "https://shopee.tw/api/v4/search/search_items?by=relevancy&keyword=" + str(keyword) +   "&limit=50&locations=-1&newest=0&order=desc&page_type=search&skip_autocorrect=1&version=2"

        data = requests.get(URL).text
        JData = json.loads(data)
        Dono = int(len(JData["items"]))
        print(Dono)
        #判斷次數
        print(DoNo(Dono))
        if Dono != 0 :
            for i in range (0,DoNo(Dono)):
                ProductInfo.itemid = JData["items"][i]['item_basic']["itemid"]
                ProductInfo.shopid = JData["items"][i]['item_basic']["shopid"]
                ProductInfo.name = JData["items"][i]['item_basic']["name"]
                ImgesUrl = JData["items"][i]['item_basic']['images']
                imgno = len(ImgesUrl)
                img=[]
                for o in range(0,imgno):
                    img += ["https://cf.shopee.tw/file/"+ ImgesUrl[o]]
                    ProductInfo.images = img
                #print(ProductInfo.images)
                
                ProductInfo.price = int(JData["items"][i]['item_basic']['price']/100000)
                ProductInfo.adskeyword = JData["items"][i]['ads_keyword']
                #print(ProductInfo.itemid)
                #print(ProductInfo.shopid)
                #print(ProductInfo.name)
                #print(ProductInfo.images)
                #print(ProductInfo.price)
                #print(ProductInfo.adskeyword)
                Pdata = [keyword,ProductInfo.itemid,ProductInfo.shopid,ProductInfo.name,ProductInfo.price]
                wb = openpyxl.load_workbook('ShopDataTT.xlsx')
                sheet = wb["Data"]
                sheet.append(Pdata+ProductInfo.images)
                wb.save("ShopDataTT.xlsx")
        else:
            pass
    else:
        pass




