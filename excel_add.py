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
    a_read = a.read_excel_file_pandas() 

class excel_add_cell():
    def excel_add_ope(self):
        a.