# 担当：大平

# import
from selenium import webdriver
import openpyxl as excel
import subprocess
from selenium.webdriver.chrome.options import Options
# /import

# static変数
filePass = "python/hello.xlsx"
driverPass = "C:/Users/akira/Documents/書類/python/chromedriver"
# /static変数



# 処理内容

# webの処理
options = Options()
# options.add_argument('--headless')#ヘッダーレスモード：コメントアウトで切り替え
driver = webdriver.Chrome(driverPass,options=options)

driver.get("https://stocks.finance.yahoo.co.jp/stocks/detail/?code=998407.O")

finalDateRate = driver.find_element_by_css_selector(
    "#main > div.marB6.chartFinance.clearFix > div.innerDate > div:nth-child(1) > dl > dd > strong")

# コンマがついていると数字で入れてくれないのでコンマを削除する
rate = finalDateRate.text.replace(',','')

driver.quit()
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
excelBook = excel.load_workbook(filePass)

sheet = excelBook['Sheet']
config = excelBook['config']

confCell =config["B2"]

# newRowNumber=confCell.value+1

newRowNumber=sheet.max_row+1

cell = sheet["A"+str(newRowNumber)]

cell.value = float(rate)

confCell.value=newRowNumber

print(cell.value)

excelBook.save(filePass)
# /excelの処理

# /処理内容

# 起動して確認
EXCEL = r'C:/Users/akira\Documents/書類/python/hello.xlsx'
subprocess.Popen(['start', '', EXCEL], shell=True)
# /起動して確認
