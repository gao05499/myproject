# 创建颜色序列字符串号
# 返回值fill的值
import random
from openpyxl.styles import PatternFill

def random_color_fill():
    colorArr = ['1','2','3','4','5','6','7','8','9','A','B','C','D','E','F']
    color = ""
    for i in range(6):
        color += colorArr[random.randint(0,14)]
    fill = PatternFill("solid", fgColor=color)
    return fill