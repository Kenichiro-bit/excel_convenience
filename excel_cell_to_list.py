from typing import Counter
from numpy import double
import openpyxl
from openpyxl.styles import PatternFill
from pandas.core.frame import DataFrame
import excel_open
import pandas
from tqdm import tqdm
import csv

a = excel_open.excel_file_open()               # MyClass のインスタンスを生成
a_read = a.read_excel_file_pandas()         # 変数 value に文字列 "abc" を代入 
b = excel_open.excel_file_open()
b_read = b.read_excel_file_pandas()
# a_read
#
list_cell_excel = []
list_cell_excel2 = []
Counter_cell = 0
count_cell = []
double_cell = []
print(type(b_read))

# b_cell = b.read_excel_pyxl()

class correct_cell():

    def __init__(self):
        self.cacee = b.pyxl_excel_color()
        self.c_cell = a.get_cell(a_read,list_cell_excel)
        self.d_cell = b.get_cell(b_read,list_cell_excel2)
        print(len(self.c_cell),len(self.d_cell))
    def correct_process(self):
        for i in tqdm(range(len(self.d_cell))):
            for j in range(len(self.c_cell)):
                if b_read.iat[i,11] == a_read.iat[j,28] :
                    # 数字の部分が列の番号
                    b.excel_color(i,self.cacee)
                    
                    # Counter_cell = Counter_cell + 1
                    double_cell.append(i)
                    count_cell.append(a_read.iat[j,28])
        b.close_pyxl_excel_color(self.cacee)


get_ok = correct_cell()
get_ok.correct_process()
print(double_cell)




