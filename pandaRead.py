import pandas as pd
"""
pip3 install pandas
pip3 install xlrd
pip3 install xlsxwriter
"""
# CSV 파일을 읽어들임 파일은 동일 폴더 내에 있음
report = pd.read_table("report.csv", sep=';', encoding='utf16', header=None)
# print(report)

# CSV 파일을 읽어서 xlsx 파일로 저장
df = pd.DataFrame(report)
exlReport = pd.ExcelWriter('excelreport.xlsx', engine='xlsxwriter')
df.to_excel(exlReport, sheet_name='CSV_Convert', header=None)
exlReport.save()

# 저장된 excel파일 읽어들임
# 지정된 레포트 포멧을 갖고있는 파일 읽어 들임
exslReport1 = pd.read_excel(
    "excelreport.xlsx", sheet_name='CSV_Convert', header=None)
dfR1 = pd.DataFrame(exslReport1)
exslReport2 = pd.read_excel("Form.xlsx", sheet_name='report', header=None)
# dfR2 = pd.DataFrame(exslReport2)
# print(exslReport2)

# CSV에서 부터 파생된 엑셀 데이터를 지정포맷 엑셀파일로 전송
exslReport2.iloc[1, 3] = exslReport1.iloc[1, 2]
exslReport2.iloc[2, 3] = exslReport1.iloc[1, 3]
exslReport2.iloc[6, 0] = exslReport1.iloc[1, 14]
shtWriter2 = pd.ExcelWriter("Form.xlsx", engine='xlsxwriter')
exslReport2.to_excel(shtWriter2, sheet_name="report",
                     header=False, index=False)
shtWriter2.save()
