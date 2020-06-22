from openpyxl import load_workbook
from position import position
from from_2D_to_1D import from_2D_to_1D
from get_col_alphabet import get_col_alphabet

xlsx_name = '求助.xlsx'
sheet_name = 'Sheet1'
domain_read = 'A2:A3'


wb = load_workbook(xlsx_name,data_only = True)
sheet_data = wb[sheet_name]

all_data_list, all_data, all_data_list_position, all_data_position, all_data_position_str = position(xlsx_name,sheet_name,domain_read)

new_data_list = []
for i in all_data:
    data_1 = i[:2] + ':' + i[2:4] + ':' + i[4:6] + ":" + i[6:8] + ":" + i[8:10] + ":" + i[10:]
    new_data_list.append(data_1)

for i,num in enumerate(all_data_position_str):
    sheet_data['B' + num[1]].value = new_data_list[i]
print(all_data_position_str)
wb.save('sample.xlsx')
