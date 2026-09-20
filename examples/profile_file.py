from excel_power_engine import ExcelEngine

path = r"C:\path\to\workbook.xlsm"
profile = ExcelEngine().profile(path)
print(profile)
