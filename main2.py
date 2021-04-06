from json import load
from urllib.parse import scheme_chars
import requests as rq
from selenium import webdriver
from selenium.common.exceptions import InvalidArgumentException, NoSuchElementException, TimeoutException
from openpyxl import styles as pxstyle
import openpyxl as px 
from bs4 import BeautifulSoup as bs 
import time
import re
from selenium.webdriver.common.keys import Keys

class Job:

    def __init__(self, driverPath):
        self.driver = webdriver.Chrome(executable_path=driverPath)
        
            
    def scrap(self, url):
        try:
            self.driver.get(url)
            time.sleep(2)
        except:
            pass
        Link_names = ["会社案内", "会社概要", "会社紹介", "about", "ABOUT", "about us", "ABOUT US", "私たちについて", "会社について", "店舗案内", "店舗紹介", "当店について"]
        
        for i in Link_names:
            try:
                element = self.driver.find_element_by_link_text(i)
                element.click()
                break
            except NoSuchElementException:
                pass
        time.sleep(2)
        html = self.driver.page_source
        soup = bs(html, 'lxml')
        try:
            table = soup.find('tbody')
            print(table)
        except:
            print("テーブルが見つかりません")
            return False
        try:
            all_text = table.get_text()
            all_text_list=all_text.split("\n")
            for i, text in enumerate(all_text_list):
                if text is "":
                    all_text_list.pop(i)
            com_name_list = ["会社名", "社名", "商号", "店舗名"]
            com_tel_list = ["電話番号", "TEL", "tel", "電話", "連絡先"]
            com_add_list = ["所在地", "住所"]
            print(all_text_list)
            com_name = self.load_info(com_name_list, all_text_list)
            com_tel = self.load_info(com_tel_list, all_text_list)
            com_add = self.load_info(com_add_list, all_text_list)
            print(com_name, com_tel, com_add)
            return com_name, com_tel, com_add
        except:
            return False
    
    def load_info(self, menu_list, all_text_data):
        for i, text in enumerate(all_text_data):
            for j in menu_list:
                if j in text:
                    return all_text_data[i+1]

    def write_excel(self, index, datas):
        try:
            file_name = "./【高林様】抽出依頼リスト - コピー.xlsx"
            book = px.load_workbook(file_name)
            sheet = book.worksheets[0]
            #社名入力
            sheet["D" + str(index)].value = datas[0]
            #tel入力
            sheet["E" + str(index)].value = datas[1]
            #住所入力
            all_address = datas[2]
            all_address = re.sub('〒[0-9]{3}-[0-9]{4}', "", all_address)
            pre_name = re.search('東京都|北海道|(?:京都|大阪)府|.{2,3}県' , all_address)
            address_com = re.split('東京都|北海道|(?:京都|大阪)府|.{2,3}県', all_address)#県名とそれ以降を分離
            add1 = pre_name.group()#県
            add2 = address_com[1]#それ以降
            sheet["F" + str(index)].value = add1
            sheet["G" + str(index)].value = add2
            book.save(file_name)
            print("\t==書き込み情報==")
            print("\t" + sheet["D" + str(index)].value)
            print("\t" + sheet["E" + str(index)].value)
            print("\t" + sheet["F" + str(index)].value)
            print("\t" + sheet["G" + str(index)].value)
        except:
            print("writing error!!")
            return False

    def write_excel2(self, index, datas):
        file_name = "./【高林様】抽出依頼リスト - コピー.xlsx"
        book = px.load_workbook(file_name)
        sheet = book.worksheets[0]
            #社名入力
        sheet["D" + str(index)].value = datas[0]
            #tel入力
        sheet["E" + str(index)].value = datas[1]
            #住所入力
        all_address = datas[2]
        all_address = re.sub('〒[0-9]{3}-[0-9]{4}', "", all_address)
        pre_name = re.search('東京都|北海道|(?:京都|大阪)府|.{2,3}県' , all_address)
        address_com = re.split('東京都|北海道|(?:京都|大阪)府|.{2,3}県', all_address)#県名とそれ以降を分離
        add1 = pre_name.group()#県
        add2 = address_com[1]#それ以降
        sheet["F" + str(index)].value = add1
        sheet["G" + str(index)].value = add2
        book.save(file_name)
        print("\t==書き込み情報==")
        print("\t" + sheet["D" + str(index)].value)
        print("\t" + sheet["E" + str(index)].value)
        print("\t" + sheet["F" + str(index)].value)
        print("\t" + sheet["G" + str(index)].value)
"""
        except:
            print("writing error!!")
            return False
"""
    
job = Job('chromedriver_win32/chromedriver.exe')
def main(index):
    file_name = "./【高林様】抽出依頼リスト - コピー.xlsx"
    book = px.load_workbook(file_name)
    sheet = book.worksheets[0]
    

    print("loading No." + str(index)) 
    url = "http://www." + sheet["C" + str(index)].value
    datas = job.scrap(url)
    if datas == False:
        print("Failed")
        return False
    job.write_excel(index, datas)

for i in range(631, 2391 + 1):
    main(i)

        
 


            
