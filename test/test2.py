import openpyxl as excel


# エクセルファイルのロード
testpass = "python/test/test.xlsx"
excelBook = excel.load_workbook(testpass)

# シートのロード
sheet = excelBook['Sheet1']

if  True:# sheet['A1'] == None:
    sheet['A2'].value='日付'
    sheet['A3'].value='終値'



# 最終行の取得と新しい行番号処理
final_row=sheet.max_row


print(final_row)

a=final_row+1

print(a)

excelBook.save(testpass)

