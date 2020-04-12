# 获取指定路径下文件夹下所有文件名
# root, dirs, files分别是目录，文件夹，文件

import os
def read_fillpath_allfillname(filePath):
    for root, dirs, files in os.walk(filePath):
        print(files)
    return files

