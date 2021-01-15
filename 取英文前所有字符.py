from openpyxl import load_workbook
from position import position
from from_2D_to_1D import from_2D_to_1D
from get_col_alphabet import get_col_alphabet
from collections import Counter
import re

xlsx_name = '取英文前所有字符.xlsx'
sheet_name = 'Sheet1'
domain_read = 'A2:A8'

wb = load_workbook(xlsx_name,data_only = True)
sheet_data = wb[sheet_name]

all_data_list, all_data, all_data_list_position, all_data_position, all_data_position_str = position(xlsx_name,sheet_name,domain_read)

position_1 = 0
for i in all_data:
    select_1 = ""
    for j in i :
        if re.match(r'[a-zA-Z]',j):
            new_position = all_data_position_str[position_1].replace("A", "C")
            sheet_data[new_position].value = str(select_1)
            break
        select_1 += j
    position_1 += 1
    print(select_1)

wb.save('sample.xlsx')
