
import requests as rq
import openpyxl
from bs4 import BeautifulSoup
from requests.exceptions import SSLError 

name = "【高林様】抽出依頼リスト - コピー.xlsx"
book = openpyxl.load_workbook(name)
sheet = book.worksheets[0]

#list of urls
urls = []
for url in sheet["C"]:
    if url.value != None or url.value != "ドメイン名":
        urls.append(url.value)
        print(url.value)

print(urls[2])

#サイトが応答するか（存在するか確認）
erorr_site = []
for i in range(2, len(urls)):
    respons = rq.get("http://" + urls[i])
    site_status = respons.status_code
    print(str(i) + ". " + urls[i] + " : ", end="")
    if site_status != 200:
        print("site Erorr!!, please check this page")
        erorr_site.append(urls[i])
        sheet.cell(row=i+1, column=1, value="x")
book.save(name)

#スクレイピング
def scrap(html):
    soup = BeautifulSoup()

    



 