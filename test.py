import requests as rq

url = "https://www.enjoy-kobo.jp"
respons = rq.get(url)
print(respons.status_code)