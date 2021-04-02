import requests as rq
from selenium import webdriver
from openpyxl import styles as pxstyle
import openpyxl as px 
from bs4 import BeautifulSoup as bs 
import time
import re
import pyperclip


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
driver = webdriver.Chrome(executable_path="./chromedriver_win32/chromedriver.exe")
driver.get("https://www.google.co.jp/maps/@35.3646982,139.5381833,15z?hl=ja")
time.sleep(3)
def com_info(com_name, com_domein, index):
    search = driver.find_element_by_xpath('//*[@id="searchboxinput"]')
    search.clear()
    search.send_keys(com_name)
    search = driver.find_element_by_xpath('//*[@id="searchbox-searchbutton"]')
    search.click()
    time.sleep(3)

    #ソースの抽出
    map_data = driver.page_source
    soup = bs(map_data, 'lxml')
    info = soup.find_all(class_="ugiz4pqJLAG__primary-text gm2-body-2")

    #ドメインが合っているか確認
    try:
        domein = info[1].text.strip()
    except:
        return False

    if domein == com_domein:
        pass
    else:
        return False
    #電話
    tel = info[2].text.strip()
    #住所
    address_data = info[0].text.strip()
    all_address = address_data[10:]
    pre_name = re.match('東京都|北海道|(?:京都|大阪)府|.{2,3}県' , all_address)
    address_com = re.split('東京都|北海道|(?:京都|大阪)府|.{2,3}県', all_address)#県名とそれ以降を分離
    add1 = pre_name.group()#県
    add2 = address_com[1]#それ以降
    
    #指定のセルへ書き込み
    print(tel)
    print(add1)
    print(add2)
    sheet["E" + str(index)].value = tel
    sheet["F" + str(index)].value = add1
    sheet["G" + str(index)].value = add2
    book.save(file_name)

for i in range(3, 101):
    company = sheet["D" + str(i)].value
    com_domein = sheet["C" + str(i)].value
    if company != None or company != "取得不可" or company != "不明なエラー":
        try:
            print("writing_data of " + company)    
            write = com_info(company, com_domein, i)
        except:
            pass
        if write == False:
            print("%s : domein Error." % (sheet["D" + str(i)].value))


    
