from typing import Counter
from numpy import double
import openpyxl
from openpyxl.styles import PatternFill
from pandas.core.frame import DataFrame
import excel_open
import pandas
from tqdm import tqdm
import csv


if __name__ == ("__main__"):
    a = excel_open.excel_file_open()               # MyClass のインスタンスを生成
    a_read = a.read_excel_file_pandas(file1,file2)         # 変数 value に文字列 "abc" を代入 
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
        self.color_excel_open = b.pyxl_excel_color()
        self.c_cell = a.get_cell(a_read,list_cell_excel)
        self.d_cell = b.get_cell(b_read,list_cell_excel2)
        print(len(self.c_cell),len(self.d_cell))
    def correct_process(self):
        get_col_num_checking = int(input('比較したい列番号を入力してください'))
        get_col_num_original = int(input('参照したい元の列番号を入力してください'))
        for i in tqdm(range(len(self.d_cell))):
            for j in range(len(self.c_cell)):
                if b_read.iat[i,get_col_num_original] != a_read.iat[j,get_col_num_checking] :
                    # 数字の部分が列の番号
                    a.excel_color(j,self.color_excel_open)
                    
                    # Counter_cell = Counter_cell + 1
                    double_cell.append(j)
                    count_cell.append(a_read.iat[j,get_col_num_checking])
        b.close_pyxl_excel_color(self.color_excel_open)


get_ok = correct_cell()
get_ok.correct_process()
print(double_cell)




