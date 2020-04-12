# 获取位置的列字符串A、B、C、D……
# 输出为列的字符列表，单层数据结构

from openpyxl.utils import get_column_letter , column_index_from_string

def get_col_alphabet(start, end):
    alphabet_list = []
    for i in range(start,end):
        alphabet_list.append(get_column_letter(i))

    return alphabet_list

