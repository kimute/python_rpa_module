from openpyxl import Workbook
import openpyxl
import datetime
wb = openpyxl.load_workbook('脆弱性スキャン管理表.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
today1= datetime.date.today()
today2= "{0:%Y/%m/%d}".format(today1)
##脆弱性スキャンの結果をexcelのsheetに入れる
#ip_addr : スキャンした複数のipを入力する、そのipに対して下記の作業が行う。
ip_addr = input("IP addressを入力してください：").split()
print(ip_addr)
#defaultは無に設定されている
cvss7 = input("CVSS7.0以上　有無？") or "無"
cvss4_69 = input("CVSS4.0 ~6.9　有無？") or "無"
cvss4 = input("CVSS4.0 ~6.9　有無？") or "無"

#範囲
cell_range = sheet['B6':'Z6']

for row in cell_range:
    for key in ip_addr:
        print(key) 
        if row[2].value == key:
            row[4].value = today2 
            row[5].value = "済"
            row[6].value = cvss7
            row[7].value = cvss4_69
            row[8].value = cvss4

print("結果：Done")

wb.save("201909脆弱性スキャン管理表_test3.xlsx")
