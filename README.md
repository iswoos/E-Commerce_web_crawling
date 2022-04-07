import requests
from bs4 import BeautifulSoup
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

options = webdriver.ChromeOptions()

options.add_argument("disable-gpu")
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36')
browser = webdriver.Chrome(options=options) #'./chromdriver.exe'

url = 'http://www.tmon.co.kr/best/52030000?bestType=best_100&gender=0&age=0&price=0&tab=2&parent=52000000'

browser.get(url)
browser.implicitly_wait(10)

items = browser.find_elements(By.CSS_SELECTOR,'ul#_dealList > li')
for item in items:
    productname = item.find_element(By.CSS_SELECTOR,'p.title_name').text
    productprice = item.find_element(By.CSS_SELECTOR,'span.price>i').text
    try:
        productreview = item.find_element(By.CSS_SELECTOR,'span.grade_average_count > span.num').text
    except:
        productreview = '리뷰없음'
    productbuy = item.find_element(By.CSS_SELECTOR,'span.buy_count').text[:-2]
    productlink = item.find_element(By.CSS_SELECTOR,'a').get_attribute('href')
    print(productreview)
