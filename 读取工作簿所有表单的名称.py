from read_sheet import read_sheet
list_sheet_name = read_sheet('工作簿1.xlsx')
for i in list_sheet_name:
    print(i)
