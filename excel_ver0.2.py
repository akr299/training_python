# 担当：大平

# import
# from selenium import webdriver
# from selenium.webdriver.chrome.options import Options

import openpyxl as excel
import subprocess
import requests
from bs4 import BeautifulSoup
import datetime

# /import

# static変数
filePass = "python/hello.xlsx"
# driverPass = "C:/Users/akira/Documents/書類/python/chromedriver"
# /static変数



# 処理内容

# webの処理

# webdriverによる情報の取得
# options = Options()
# # options.add_argument('--headless')#ヘッダーレスモード：コメントアウトで切り替え
# driver = webdriver.Chrome(driverPass,options=options)

# driver.get("https://stocks.finance.yahoo.co.jp/stocks/detail/?code=998407.O")

# finalDateRate = driver.find_element_by_css_selector(
#     "#main > div.marB6.chartFinance.clearFix > div.innerDate > div:nth-child(1) > dl > dd > strong")

# コンマがついていると数字で入れてくれないのでコンマを削除する
# /webdriverによる情報の取得

# requestsとBeautifulSoupによる情報の取得

# URLのHTMLボディの取得
r=requests.get('https://stocks.finance.yahoo.co.jp/stocks/detail/?code=998407.O')

# HTMLボディのパース処理
soup=BeautifulSoup(r.text,'html.parser')


today_index=soup.select_one('#main > div.marB6.chartFinance.clearFix > div.innerDate > div:nth-child(1) > dl > dd > strong').text

print(today_index)

rate = today_index.replace(',','')

print(rate)
# /webの処理

# excelの処理
#変数
config=1
#/変数

# 関数
# def if_test(num):
#     if:

#     else:

# /関数

# エクセルファイルのロード
excelBook = excel.load_workbook(filePass)

# シートのロード
sheet = excelBook['Sheet']
config = excelBook['config']

# 現在日時の取得
dt_now=datetime.date.today()

if sheet.cell(row=2,column=2).value == None:
    sheet.cell(row=2,column=2).value='日付'
    sheet.cell(row=2,column=3).value='終値'


# 最終行の取得と新しい行番号処理
newRowNumber=sheet.max_row+1
print(newRowNumber)
# セルの取得
cell=sheet.cell(row=newRowNumber,column=3)

# 値の書き込み
cell.value=float(rate)

# セルの値のチェック
print(cell.value)

# セーブ
excelBook.save(filePass)

# /excelの処理

# /処理内容

# 起動して確認
EXCEL = r'C:/Users/akira\Documents/書類/python/hello.xlsx'
subprocess.Popen(['start', '', EXCEL], shell=True)
# /起動して確認
