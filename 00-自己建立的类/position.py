# 读取excel表格指定区间domain_data的数据，并范围表格里面的字符串，以及坐标位置
# 3个参数，5个返回值
# 返回值all_data_list（带有列表的数据字符串）,
# 返回值all_data（不带有列表的数据字符串）
# 返回值all_data_list_position（带有列表的数据坐标）
# 返回值all_data_position（不带有列表的数据坐标）
# 返回值all_data_position_str（不带有列表的数据坐标字符串）

def position(excel_name,form_name,domain_data):
    from openpyxl import load_workbook

    # excel_name = '统计.xlsx'  #要读取excel的名称
    # form_name = 'Sheet1'  #要读取表单的名称
    wb = load_workbook(excel_name,data_only = True)
    sheet_data = wb[form_name]

    # domain_data = 'A1:L31'  #要读取表格的区间
    data_id_1 = list(sheet_data[domain_data])
    all_data_list = []
    sub_data_list = []
    all_data = []
    all_data_position = []
    all_data_list_position = []
    sub_data_list_position = []
    all_data_position_str = []
    for i in data_id_1:
        for j in i:
            sub_data_list.append(j.value)
            all_data.append(j.value)
            all_data_position.append(j)
            sub_data_list_position.append(j)
            all_data_position_str.append(str(j).split('.')[1][:-1])

        all_data_list.append(sub_data_list)
        sub_data_list = []
        all_data_list_position.append(sub_data_list_position)
        sub_data_list_position = []

    # print(all_data_list)
    # print(all_data)
    # print(all_data_list_position)
    # print(all_data_position)
    return all_data_list, all_data, all_data_list_position, all_data_position, all_data_position_str

