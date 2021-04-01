from logging import NOTSET
from urllib.parse import scheme_chars
from openpyxl.worksheet.dimensions import SheetFormatProperties
import requests as rq
from selenium import webdriver
from openpyxl import styles as pxstyle
import openpyxl as px 
from bs4 import BeautifulSoup as bs 
import time
import re

from selenium.webdriver.chrome.webdriver import WebDriver


#sheetの読み込み
file_name = "./【高林様】抽出依頼リスト - コピー.xlsx"
book = px.load_workbook(file_name)
sheet = book.worksheets[0]

"""
#サイトの存在確認
fill_red = pxstyle.PatternFill(patternType='solid', fgColor='FF0000', bgColor='FF0000')#応答がないサイトには赤色で塗る
fill_whi = pxstyle.PatternFill(patternType='solid', fgColor='ffffff', bgColor='ffffff')
for i in range(3, sheet.max_row):
    index = "C"+str(i)
    url = sheet[index].value
    print(index + " . " + url + " : ", end="")
    try:
        respons = rq.get("http://www." + url)
        if respons.status_code != 200:#閲覧できないとき
            print(respons.status_code)
            sheet[index].fill = fill_red
            sheet["D" + str(i)].value = "取得不可"
        else:#通常通り
            page_url = respons.url
            print(respons.status_code)
            print(page_url)
            sheet[index].fill = fill_whi
            #SSLかどうか（httpsで始まっているかどうか）
            if page_url.startswith("https://"):
                print("SSL 有")
                sheet["H" + str(i)].value = "有"
            else:
                sheet["H" + str(i)].value = "無"
            
    except:#そもそもサイトが存在しない？？
        print(respons.status_code)
        sheet[index].fill = fill_red
        sheet["D"+str(i)].value = "不明なエラー"
        pass
book.save(file_name)
"""

#会社情報のスクレイピング
def com_info(com_name, com_domein, index):
    driver = webdriver.Chrome(executable_path="./chromedriver_win32/chromedriver.exe")
    driver.get("https://www.google.co.jp/maps/@35.3646982,139.5381833,15z?hl=ja")
    time.sleep(3)
    serch = driver.find_element_by_xpath('//*[@id="searchboxinput"]')
    serch.send_keys(com_name)
    driver.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    driver.click()
    time.sleep(3)
    #ドメインが合っているか確認
    domein = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[10]/div[2]/button/div[1]/div[2]/div[1]').text
    if domein != com_domein:
        print("ドメイン名が違います")
        return False #間違っていれば関数を終了
    else:
        pass
    #電話番号の抽出
    tel = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[10]/div[3]/button/div[1]/div[2]/div[1]').text
    #住所の抽出
    driver.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    driver.click()
    time.sleep(3)
    address_data = driver.find_element_by_xpath('//*[@id="pane"]/div/div[1]/div/div/div[10]/div[1]/button/div[1]/div[2]/div[1]').text
    all_address = address_data[9:]
    address_com = re.split('[都道府県]', all_address)#県名とそれ以降を分離
    add1 = address_com[0]#県
    add2 = address_com[1]#それ以降
    
    #指定のセルへ書き込み
    sheet["E" + str(index)].value = tel
    sheet["F" + str(index)].value = add1
    sheet["G" + str(index)].value = add2

for i in range(3, 10):
    company = sheet["D" + str(i)].value
    com_domein = sheet["C" + str(i)].value
    if company != None:
        print("writing_data of " + company)    
        write = com_info(company, com_domein, i)
        if write == False:
            print("%s : domein Error." % (sheet["D" + str(i)].value))


    
