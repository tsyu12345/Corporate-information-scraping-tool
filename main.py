from openpyxl.worksheet.dimensions import SheetFormatProperties
import requests as rq
from selenium import webdriver
from openpyxl import styles as pxstyle
import openpyxl as px 
from bs4 import BeautifulSoup as bs 
import time

from selenium.webdriver.chrome.webdriver import WebDriver


#sheetの読み込み
file_name = "./【高林様】抽出依頼リスト - コピー.xlsx"
book = px.load_workbook(file_name)
sheet = book.worksheets[0]

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


