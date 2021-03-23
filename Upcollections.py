#!/usr/local/bin/python
# -*- coding: utf-8 -*-
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

DOADD = {
    "1":"時尚彩妝",
    "2":"口紅",
    "3":"眼影",
    "4":"乳液",
    "5":"氣墊/粉凝霜",
    "6":"BB/CC/粉底液",
    "7":"睡眠面膜",
    "8":"腮紅/修容",
    "9":"眉筆",
    "10":"眼線/睫毛",
    "11":"美妝工具",
    "12":"香水",
    "13":"妝前/隔離/美白霜",
    "14":"眉筆",
    "15":"幹粉/濕粉/定妝粉",
    "16":"身體美白霜",    
    "16":"修容/打亮",
    "16":"睫毛",
}






options = Options()
chrome_path = "C:\\PYTHON\\driver\\chromedriver" 
tmap = int(time.time())
localtime = time.strftime("%m%d",time.localtime())
#print(localtime)
driver = webdriver.Chrome(chrome_path,options=options) 
driver.get("https://admin.easystore.co/products/collections")
driver.find_element_by_css_selector('#app > div > main > div > section > div > form > div:nth-child(5) > div > div > div > input[type=email]').click()
driver.find_element_by_css_selector('#app > div > main > div > section > div > form > div:nth-child(5) > div > div > div > input[type=email]').send_keys("plauvvnn@gmail.com")
driver.find_element_by_css_selector('#app > div > main > div > section > div > form > div:nth-child(6) > div > div > div > input[type=password]').click()
driver.find_element_by_css_selector('#app > div > main > div > section > div > form > div:nth-child(6) > div > div > div > input[type=password]').send_keys("Aa5212230619")
driver.find_element_by_css_selector('#app > div > main > div > section > div > form > div:nth-child(7) > button').click()
sleep(3)
driver.find_element_by_css_selector('#app > div > main > div > section > div > div > div:nth-child(1) > div > div.card-wrapper > div.card > div').click()
sleep(3)
driver.get("https://admin.easystore.co/products/collections")

for i , ii in DOADD.items():
    sleep(2)
    driver.find_element_by_css_selector('#app > div > main > div > section:nth-child(1) > div > div.navigation > div > a').click()
    sleep(2)
    driver.find_element_by_css_selector('#title').send_keys(ii)
    sleep(2)
    driver.find_element_by_css_selector('#app > div > main > div > form > section:nth-child(1) > div.top-bar.focus > button').click()
    sleep(2)




