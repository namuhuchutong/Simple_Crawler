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

url = 'https://www.amazon.com/'
sign_in = '/html/body/div[1]/header/div/div[1]/div[3]/div/a[2]'

amazon_id = os.environ['AMAZON_ID']
amazon_pwd = os.environ['AMAZON_PWD']

best_category = []

with open('sns_14_2/category_list.txt', 'r') as f:
    best_category = f.readlines()

print('-'*50)

for i, x in enumerate(best_category):
    print(str(i+1) + ". " + x, end="")
print('-'*50)

category_idx = int(input("input the number : "))
count = int(input('how many? : '))
file_path = input('file path : ')

data = []
img_urls = []
cnt = 0

with webdriver.Chrome('/Users/hoyeonjang/Downloads/chromedriver') as driver:
    wait = WebDriverWait(driver, 10)

    driver.get(url)
    wait.until(presence_of_element_located((By.XPATH, sign_in))).click()

    e = driver.find_element(By.XPATH, '//*[@id="ap_email"]')
    e.send_keys(amazon_id)
    e.send_keys(Keys.RETURN)
    time.sleep(3)

    e = driver.find_element(By.XPATH, '//*[@id="ap_password"]')
    e.send_keys(amazon_pwd)
    e.send_keys(Keys.RETURN)

    wait.until(presence_of_element_located((By.XPATH, '//*[@id="nav-xshop"]/a[1]'))).click()
    time.sleep(3)

    html = driver.page_source
    soup = bs(html, 'html.parser')

    hrefs = soup.find('ul', {'id': 'zg_browseRoot'}).find_all('a', href=True)

    driver.get(hrefs[category_idx - 1]['href'])
    time.sleep(3)

    best1 = driver.page_source
    soup1 = bs(best1, 'html.parser')

    driver.find_element(By.XPATH, '/html/body/div[1]/div[3]/div/div/div[1]/div/div[2]/div/ul/li[4]/a').click()
    time.sleep(3)

    best2 = driver.page_source
    soup2 = bs(best2, 'html.parser')

    elements = soup1.find_all('li', {'class': 'zg-item-immersion', 'role': 'gridcell'})

    for e in elements:

        if cnt >= count:
            break

        info = []

        try:
            rank = cnt + 1
            p_info = e.find('div', class_='p13n-sc-truncated').text
            price = e.find('span', class_='p13n-sc-price').text
            comment_num = e.find('span', class_='a-icon-alt').text
            stars = e.find('a', class_='a-size-small a-link-normal').text
        except Exception as ex:
            print(ex)
            continue

        info.append(rank)
        info.append(p_info)
        info.append(price)
        info.append(comment_num)
        info.append(stars)

        img_urls.append(e.find('img')['src'])
        data.append(info)
        cnt += 1

    elements = soup2.find_all('li', {'class': 'zg-item-immersion', 'role': 'gridcell'})

    for e in elements:

        if cnt >= count:
            break

        info = []

        try:
            rank = cnt + 1
            p_info = e.find('div', class_='p13n-sc-truncated').text
            price = e.find('span', class_='p13n-sc-price').text
            comment_num = e.find('span', class_='a-icon-alt').text
            stars = e.find('a', class_='a-size-small a-link-normal').text
        except Exception as ex:
            print(ex)
            info.append('')

        info.append(rank)
        info.append(p_info)
        info.append(price)
        info.append(comment_num)
        info.append(stars)

        img_urls.append(e.find('img')['src'])
        data.append(info)
        cnt += 1

df = pd.DataFrame(data, columns=['판매순위', '제품소개', '가격', '상품평수', '평점'])

try:
    if not os.path.exists(file_path):
        os.makedirs(file_path)
except OSError:
    print('Error: Creating directory. ' + file_path)

try:
    df.to_csv(file_path+"/a.csv")
    df.to_excel(file_path+"/b.xlsx", sheet_name="new")
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