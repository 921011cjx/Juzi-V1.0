# @Author:橘子
# @email :2315253816@qq.com
# @Time  :2021/4/21 18:01
# @File  :run.py
from comment import juzi01
from test_data import test
import openpyxl
import requests  # 引用第三方库
import time

if __name__ == '__main__':
    juzi01.execute_fun(test.filename, test.sheetname)  # 调用函数
