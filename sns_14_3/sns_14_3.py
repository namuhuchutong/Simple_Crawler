import pandas as pd
import sys
from bs4 import BeautifulSoup as bs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.expected_conditions import presence_of_element_located
import time
import xlwt
import openpyxl
import numpy as np
import re
import requests
import copy
import os
import urllib
import xlsxwriter
import io

url = 'http://corners.gmarket.co.kr/Bestsellers'
columns = ['판매순위', '제품소개', '원래가격', '판매가격', '할인률']

file_path = input('input file path : ')

with webdriver.Chrome('/Users/hoyeonjang/Downloads/chromedriver') as driver:
    wait = WebDriverWait(driver, 10)

    driver.get(url)

    html = driver.page_source
    soup = bs(html, 'html.parser')

div = soup.find('div', class_='best-list').find_next_sibling('div')
lists = div.find_all('li')

data = []

for i, li in enumerate(lists):
    info = []

    try:
        p_name = li.find('a', class_='itemname').text
        s_price = li.find('div', class_='s-price').find('span').text
        if li.find('div', class_='o-price').find('span') == None or li.find('div', class_='s-price').find('em') == None:
            o_price = s_price
            sale = 'NA'
        else:
            o_price = li.find('div', class_='o-price').find('span').text
            sale = li.find('div', class_='s-price').find('em').text
    except Exception as err:
        print(err)
        continue

    info.append(i + 1)
    info.append(p_name)
    info.append(o_price)
    info.append(s_price)
    info.append(sale)

    data.append(info)

df = pd.DataFrame(data, columns=columns)

try:
    if not os.path.exists(file_path):
        os.makedirs(file_path)
except OSError:
    print('Error: Creating directory. ' + file_path)


try:
    df.to_csv(file_path+"/a.csv")
    df.to_excel(file_path+"/b.xls", sheet_name="new")
    np.savetxt(file_path+"/c.txt", df.values, delimiter='\n', fmt='%s')
except FileNotFoundError:
    print("Check your path")

"""

    엑셀에 상품 이미지 추가 불가.

     - https://xlsxwriter.readthedocs.io/bugs.html -

    *Images not displayed correctly in Excel 2001 for Mac and non-Excel applications*

    Images inserted into worksheets via insert_image() may not display correctly in Excel 2011 for Mac and non-Excel applications such as OpenOffice and LibreOffice.
    Specifically the images may looked stretched or squashed.
    This is not specifically an XlsxWriter issue. It also occurs with files created in Excel 2007 and Excel 2010.


    아래 코드는 정상 작동하나 맥에서는 결과를 볼 수 없음.

workbook = xlsxwriter.Workbook(file_path+'/b.xlsx')
worksheet = workbook.add_worksheet()

for i, img in enumerate(img_urls):
    img_data = io.BytesIO(urllib.request.urlopen(img).read())
    worksheet.insert_image('C'+str(i+2), img, {'image_data':img_data})

"""