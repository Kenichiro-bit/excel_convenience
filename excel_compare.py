from openpyxl.styles import PatternFill
from pandas.core.frame import DataFrame
from excel_open import excel_get_cell
import excel_open
from tqdm import tqdm
import file_initial
import excel_color

file1 = file_initial.excel_file1
sheet1 = file_initial.excel_sheet1

file2 = file_initial.excel_file2
sheet2 = file_initial.excel_sheet2

check_num = file_initial.get_col_num_checking
color_num = file_initial.get_col_num_original

if __name__ == ("__main__"):
    a = excel_open.excel_file_open()               # MyClass のインスタンスを生成
    a_read = a.read_excel_file_pandas(file1,sheet1)         # 変数 value に文字列 "abc" を代入 
    b = excel_open.excel_file_open()
    b_read = b.read_excel_file_pandas(file2,sheet2)
    excel_cell = excel_get_cell()

list_cell_excel = []
list_cell_excel2 = []
count_cell = []
cell_number = []
print(b_read[:7])

class correct_cell():
    def __init__(self):
        self.a_cell = excel_cell.get_cell(a_read,list_cell_excel)
        self.b_cell = excel_cell.get_cell(b_read,list_cell_excel2)

    def correct_process(self): ## 行数列数が異なるとき
        get_col_num_checking = int(input('比較したい列番号を入力してください'))
        get_col_num_original = int(input('参照したい元の列番号を入力してください'))
        for i in tqdm(range(len(self.a_cell))):
            for j in range(len(self.b_cell)):
                if b_read.iat[i,get_col_num_original] == a_read.iat[j,get_col_num_checking] :
                    cell_number.append(i)
                    count_cell.append(a_read.iat[j,get_col_num_checking])

    def correct_check(self): ##同じとき
        for i in tqdm(range(len(self.a_cell))):
            if b_read.iat[i,check_num] == a_read.iat[i,color_num] :
                cell_number.append(i)
                count_cell.append(a_read.iat[i,color_num])
                print(b_read.iat[i,check_num],a_read.iat[i,color_num])
        


get_ok = correct_cell()
if len(a_read) != len(b_read):
    get_ok.correct_process()
else:
    get_ok.correct_check()
print(cell_number)

color = excel_color.excel_fill_color()
file_excel = color.open_excel_color(file1)
for i in range(len(cell_number)):
    color.excel_color(int(cell_number[i]),file_excel,sheet2,color_num)
color.close_pyxl_excel_color(file_excel)



