import re
import requests as rq
from selenium import webdriver
import time
from bs4 import BeautifulSoup as bs 

address_data = "〒329-2224 栃木県塩谷郡塩谷町金枝６６５"
all_address = address_data[10:]
pre_name = re.match('東京都|北海道|(?:京都|大阪)府|.{2,3}県' , all_address)
address_com = re.split('東京都|北海道|(?:京都|大阪)府|.{2,3}県', all_address)#県名とそれ以降を分離
add1 = pre_name.group()#県
add2 = address_com[1]#それ以降
print(add1)
print(add2)
#会社情報のスクレイピング

driver = webdriver.Chrome(executable_path="./chromedriver_win32/chromedriver.exe")
driver.get("https://www.google.co.jp/maps/@35.3646982,139.5381833,15z?hl=ja")
time.sleep(3)
def com_info(com_name, com_domein, index):
    search = driver.find_element_by_xpath('//*[@id="searchboxinput"]')
    search.clear()
    search.send_keys(com_name)
    search = driver.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    search.click()
    time.sleep(4)
    #ソースの抽出
    map_data = driver.page_source
    soup = bs(map_data, 'lxml')
    info = soup.find_all(class_="ugiz4pqJLAG__primary-text gm2-body-2")
    print(info[0].text.strip())
    print(info[1].text.strip())

com_info("西松接骨院", "4211.jp", 0)