from selenium import webdriver
import openpyxl as excel
import subprocess

EXCEL = r'C:/Users/akira\Documents/書類/python/hello.xlsx'
subprocess.Popen(['start', '', EXCEL], shell=True)

# driver = webdriver.Chrome("C:/Users/akira/Documents/書類/python/chromedriver")

# driver.get("https://www.google.co.jp")


# driver.quit()

# book = excel.Workbook()
# sheet = book.active

# sheet["B2"] = "こんにちは"

# book.save("hello.xlsx")