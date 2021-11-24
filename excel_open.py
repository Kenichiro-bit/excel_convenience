import pandas as pd
import openpyxl


##　エクセルの読出し、書き込みを行う
## pandas用のファイルオープン
class excel_file_open():

    def read_excel_file_pandas(self, file, sheet):
        excel1 = pd.read_excel(file,sheet_name=sheet)
        excel_file_name1_split = file.split('.')
        excel_file_name1_split[0] = excel_file_name1_split[0].replace(excel_file_name1_split[0],excel_file_name1_split[0]+'_backup')
        excel_backup_file = '.'.join(excel_file_name1_split)
        excel1.to_excel(excel_backup_file)
        
        return excel1


## セルをリストに入れる
class  excel_get_cell():
    def get_cell(self,x_read,list_cell_excelx):
        for i in range(len(x_read)):
            list_cell_excelx.append(x_read.iat[i,0])

        return list_cell_excelx

class excel_only_open():
    def __init__(self):
        self.file_name1 = 'D:/UserArea/J0134661/Desktop/便利機能/excel_check/コピーFMEA試験正解表_次世代PDRV_OBD対象⊿8_masuo.xlsx'
        self.file_sheet_name1 = input('エクセルファイルのsheet名を記入>>')
        print(self.file_sheet_name1)

    def open_excel_file(self):
        self.excel_pyxl_open = openpyxl.load_workbook(self.file_name1)
        return self.excel_pyxl_open

    def open_excel_sheet(self,excel_pyxl_open):
        self.execel_pyxl_sheet = self.execel_pyxl_open[self.file_sheet_name1]
        return 

