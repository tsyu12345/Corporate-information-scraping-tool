import re
import requests as rq
from selenium import webdriver
import time
from bs4 import BeautifulSoup as bs 

address_data = "〒251-0025愛媛県藤沢市鵠沼石上1-4-5湘南ライトビル1F"
address_data = re.sub('〒[0-9]{3}-[0-9]{4}', "", address_data)
print(address_data)
pre_name = re.search('東京都|北海道|(?:京都|大阪)府|.{2,3}県' , address_data)
address_com = re.split('東京都|北海道|(?:京都|大阪)府|.{2,3}県', address_data)#県名とそれ以降を分離
add1 = pre_name.group()#県
add2 = address_com[1]#それ以降
print(add1)
print(add2)
