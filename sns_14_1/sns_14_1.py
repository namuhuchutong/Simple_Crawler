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

url = 'https://movie.naver.com/'
movie = input("Input Movie name : ")


def next_page(pageAnchor):
    wait.until(presence_of_element_located((By.CLASS_NAME, pageAnchor))).click()
    time.sleep(1)


data = []

with webdriver.Chrome('/Users/hoyeonjang/Downloads/chromedriver') as driver:
    wait = WebDriverWait(driver, 10)
    driver.get(url)
    driver.find_element_by_xpath('/html/body/div/div[2]/div/div/fieldset/div/span/input').send_keys(movie + Keys.RETURN)
    wait.until(presence_of_element_located(
        (By.XPATH, '/html/body/div/div[4]/div/div/div/div/div[1]/ul[2]/li[1]/dl/dt/a'))).click()
    newURL = driver.current_url
    driver.get(newURL.replace('basic', 'point'))

    driver.switch_to.frame('pointAfterListIframe')

    pa = 'pg_next'

    count = 0

    while count < 100:

        page = driver.page_source
        soup = bs(page, 'html.parser')

        elements = soup.find('div', class_='score_result').findAll('li')
        for i, li in enumerate(elements):

            if (count == 100):
                break

            info = []
            info.append(li.findChild('div', class_='star_score').find('em').text)

            id_name = '_filtered_ment_' + str(i)

            info.append(li.findChild('span', {'id': id_name}).text.strip())
            print(info[0])
            info.append(li.findChild('a', {'target': '_top'}).find('span').text)
            print(info[1])
            info.append(li.findChild('dt').find('em').find_next('em').text)
            print(info[2])
            info.append(li.findChild('strong').text)
            print(info[3])
            info.append(li.findChild('strong').find_next('strong').text)
            print(info[4])

            data.append(info)
            count += 1

        next_page(pa)

df = pd.DataFrame(data, columns=['별점', '리뷰내용', '작성자', '작성일자', '공감 횟수', '비공감 횟수'])

file_path = 'data3/'

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