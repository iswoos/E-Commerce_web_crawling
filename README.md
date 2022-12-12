import time
import re
from datetime import datetime
from typing import ItemsView
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import requests
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import openpyxl


search_product =['고구마']


for i in search_product:
    excel_file = openpyxl.Workbook()

    options = webdriver.ChromeOptions()
    #options.add_argument('headless')
    options.add_argument("--disable-blink-features")
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36')

    browser = webdriver.Chrome(options=options) #'./chromdriver.exe'

    excel_sheet3 = excel_file.active
    excel_sheet3.append(['순위','상품명', '가격','리뷰수', '구매수','링크'])
    excel_sheet3.title = '지마켓'+i
    excel_sheet3.freeze_panes = 'A2'
    excel_sheet3.column_dimensions['A'].width = 5
    excel_sheet3.column_dimensions['B'].width = 80
    excel_sheet3.column_dimensions['C'].width = 10
    excel_sheet3.column_dimensions['D'].width = 10
    excel_sheet3.column_dimensions['E'].width = 10
    excel_sheet3.column_dimensions['E'].width = 10


    browser.get('https://www.gmarket.co.kr/')
    time.sleep(2)
    browser.find_element(By.CSS_SELECTOR,'#skip-navigation-search > span > input').send_keys(i)
    browser.find_element(By.CSS_SELECTOR,'#skip-navigation-search > span > input').send_keys(Keys.ENTER)
    browser.find_element(By.CSS_SELECTOR,'#region__content-status-information > div > div > div.box__control-area > div.box__sort-control > div.box__sort-control-selected > button').click()
    browser.find_element(By.CSS_SELECTOR,'#region__content-status-information > div > div > div.box__control-area > div.box__sort-control.box__sort-control--active > div.box__sort-control-list > ul > li:nth-child(2) > a').click()
    time.sleep(3)
    G_source = browser.page_source
    G_soup = BeautifulSoup(G_source,'html.parser')


    items = G_soup.select('#section__inner-content-body-container > div:nth-child(2) > div')
    for index, item in enumerate(items,start=1):

        productname = item.select_one('div.box__item-title span.text__item').text
        productprice = item.select_one('div.box__item-price strong.text.text__value').text
        try:
            productreview = item.select_one('div.box__information-score li.list-item.list-item__feedback-count span.text').text[1:-1]
        except:
            pass
        try:
            productbuy = item.select_one('div.box__information-score li.list-item.list-item__pay-count span.text').text[2:]
        except:
            pass
        try:
            productlink = item.select_one('div.box__item-title > span > a')['href']
        except:
            pass
        print(productname,productprice,productreview,productbuy,productlink)
        excel_sheet3.append([index,productname,productprice,productreview,productbuy,productlink])
        excel_sheet3.cell(row=index+1, column=6).hyperlink = productlink



    excel_sheet4 =excel_file.create_sheet('티몬'+i)
    excel_sheet4.append(['순위','상품명', '가격','배송비','리뷰수', '구매수','링크'])
    excel_sheet4.freeze_panes = 'A2'
    excel_sheet4.column_dimensions['A'].width = 5
    excel_sheet4.column_dimensions['B'].width = 80
    excel_sheet4.column_dimensions['C'].width = 10
    excel_sheet4.column_dimensions['D'].width = 15
    excel_sheet4.column_dimensions['E'].width = 10
    excel_sheet4.column_dimensions['F'].width = 10
    excel_sheet4.column_dimensions['G'].width = 10


    browser.get('https://www.tmon.co.kr/')
    browser.implicitly_wait(10)
    browser.find_element_by_name('keyword').send_keys(i)
    browser.find_element_by_class_name('btn_search').click()
    time.sleep(3)
    T_source = browser.page_source
    T_soup = BeautifulSoup(T_source,'html.parser')

    items = T_soup.select('div.deallist_wrap > ul > li')
    for index, item in enumerate (items,start=1):
        ad_badge = item.select_one('div.deal_info > div.label_area > span.lyr_info > button')
        if ad_badge:
            continue
        else:
            pass
        productname = item.select_one('div.deal_info > p').text
        productprice = item.select_one('div.deal_info span.price i.num').text


        #search_app > div.ct_wrap > section.search_deallist > div.deallist_wrap > ul > li:nth-child(1) > a > div.deal_info > div.price_area > div.label_free_shipping > span.text
        #search_app > div.ct_wrap > section.search_deallist > div.deallist_wrap > ul > li:nth-child(3) > a > div.deal_info > div.price_area > div.label_free_shipping > span.text

        productbbprice = item.select_one('div.deal_info > div.price_area > div.label_free_shipping > span.text')
        if productbbprice:
            productbprice = ""
        else:
            productbprice = "배송비 별도"
        try:
            productreview = item.select_one('div.deal_info span.grade_average_count > span.num').text
        except:
            productreview = '리뷰없음'
        try:
            productbuy = item.select_one('div.deal_info span.buy_count').text[:-4]
        except:
            productbuy = '구매없음'
        productlink = item.select_one('li.item > a')['href']

        print(productname,productprice,productbprice,productreview,productbuy,productlink)
        excel_sheet4.append([index,productname,productprice,productbprice,productreview,productbuy,productlink])
        excel_sheet4.cell(row=index+1, column=7).hyperlink = productlink


    excel_sheet5 =excel_file.create_sheet('11번가'+i)
    excel_sheet5.append(['순위','상품명', '가격','배송비','리뷰수', '공급사','링크'])
    excel_sheet5.freeze_panes = 'A2'
    excel_sheet5.column_dimensions['A'].width = 5
    excel_sheet5.column_dimensions['B'].width = 80
    excel_sheet5.column_dimensions['C'].width = 10
    excel_sheet5.column_dimensions['D'].width = 15
    excel_sheet5.column_dimensions['E'].width = 10
    excel_sheet5.column_dimensions['F'].width = 10
    excel_sheet5.column_dimensions['G'].width = 10

    browser.get('https://www.11st.co.kr/main')
    time.sleep(2)
    browser.find_element(By.CSS_SELECTOR,'#tSearch > form > fieldset > input').send_keys(i)
    browser.find_element(By.CSS_SELECTOR,'#tSearch > form > fieldset > input').send_keys(Keys.ENTER)
    time.sleep(2)
    browser.find_element(By.CSS_SELECTOR,'#layBodyWrap > div > div > div.l_search_content > div > div.result_filter_wrap > div > div.filter_cont > div > div > button').click()
    browser.find_element(By.CSS_SELECTOR,'#layBodyWrap > div > div > div.l_search_content > div > div.result_filter_wrap > div > div.filter_cont > div > ul > li:nth-child(1) > button').click()
    time.sleep(3)
    st11_source = browser.page_source
    st11_soup = BeautifulSoup(st11_source,'html.parser')

    items = st11_soup.select('#layBodyWrap > div > div > div.l_search_content > div > section:nth-child(3) > ul > li')
    for index, item in enumerate(items,start=1):

        #layBodyWrap > div > div > div.l_search_content > div > section:nth-child(3) > ul > li:nth-child(1) > div > div:nth-child(2) > div.c_card_info_top > div.c_prd_name.c_prd_name_row_1 > a

        productname = item.select_one('div > div.c_card_info_top > div.c_prd_name.c_prd_name_row_1 > a > strong').text             
        productprice = item.select_one('div.c_card_info_top > div.c_prd_price > dl > dd > span.value').text
        try:
            productbbprice = item.select_one('div.c_card_info_top > div.c_prd_delivery > span').text
        except:
            pass

        if  productbbprice == "무료배송":
            productbprice = ""
        else:
            productbprice = item.select_one('div.c_card_info_top > div.c_prd_delivery > span').text
        try:
            productreview = item.select_one('div.c_card_info_top > div.c_prd_meta > a > em').text
        except:
            pass
        try:
            productsupplier = item.select_one('div.c_prd_seller > a > span').text
        except:
            pass
        try:
            productlink = item.select_one('div > div:nth-child(2) > div.c_card_info_top > div.c_prd_name.c_prd_name_row_1 > a')['href']
        except:
            pass
        print(productname,productprice,productbprice,productreview,productsupplier,productlink)
        excel_sheet5.append([index,productname,productprice,productbprice,productreview,productsupplier,productlink])
        excel_sheet5.cell(row=index+1, column=7).hyperlink = productlink

        cell_A1 = excel_sheet3['A1'] # 셀 선택하기
        cell_A1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_A1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_B1 = excel_sheet3['B1'] # 셀 선택하기
        cell_B1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_B1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_C1 = excel_sheet3['C1'] # 셀 선택하기
        cell_C1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_C1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_D1 = excel_sheet3['D1'] # 셀 선택하기
        cell_D1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_D1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_E1 = excel_sheet3['E1'] # 셀 선택하기
        cell_E1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_E1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_F1 = excel_sheet3['F1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors





        cell_A1 = excel_sheet4['A1'] # 셀 선택하기
        cell_A1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_A1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_B1 = excel_sheet4['B1'] # 셀 선택하기
        cell_B1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_B1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_C1 = excel_sheet4['C1'] # 셀 선택하기
        cell_C1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_C1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_D1 = excel_sheet4['D1'] # 셀 선택하기
        cell_D1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_D1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_E1 = excel_sheet4['E1'] # 셀 선택하기
        cell_E1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_E1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_F1 = excel_sheet4['F1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_F1 = excel_sheet4['G1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors        cell_F1 = excel_sheet4['F1'] # 셀 선택하기

        cell_F1 = excel_sheet4['H1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors





        cell_A1 = excel_sheet5['A1'] # 셀 선택하기
        cell_A1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_A1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_B1 = excel_sheet5['B1'] # 셀 선택하기
        cell_B1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_B1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_C1 = excel_sheet5['C1'] # 셀 선택하기
        cell_C1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_C1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_D1 = excel_sheet5['D1'] # 셀 선택하기
        cell_D1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_D1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_E1 = excel_sheet5['E1'] # 셀 선택하기
        cell_E1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_E1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_F1 = excel_sheet5['F1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

        cell_F1 = excel_sheet5['G1'] # 셀 선택하기
        cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
        cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
        # 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors        cell_F1 = excel_sheet4['F1'] # 셀 선택하기
 
    excel_file.save(i+' 지마켓,티몬,11번가 베스트'+'.xlsx')
    excel_file.close()
    
