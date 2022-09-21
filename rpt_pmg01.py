if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['make_202210'] # 目前路徑
    sys.path.append(config_path)

import time
import openpyxl
import pandas as pd
import PySimpleGUI as sg
from tool_excel2 import tool_excel
from tool_style import *
import tool_file
import tool_db_make
from config import *

class Report_pmg01(tool_excel):
    def __init__(self, filename):
        self.fileName = filename
        self.report_name = 'pmg01' # 工程資料明細表
        self.report_dir = config_report_dir # 資料夾名稱
        self.report_path = os.path.join(os.path.expanduser(r'~\Documents'), self.report_dir) #資料夾路徑
        self.file_tool = tool_file.File_tool() # 檔案工具並初始化資料夾
        if self.report_path is None:
            print('找不到路徑')
            sys.exit() #正式結束程式
        self.file_tool.clear(self.report_name) # 清除舊檔
        self.db =  tool_db_make.db_make()
        self.comp_base()
        self.create_excel()
        self.output()
        self.save_xls()
        self.open_xls()

    def create_excel(self):
        self.wb = openpyxl.Workbook()
        wb = self.wb
        self.df_dp = self.db.get_dp_df() # 所有部門
        for i, r in self.df_dp.iterrows():
            wb.create_sheet(r['dp002'])

        wb.remove(wb['Sheet'])
        self.xlsfile = os.path.join(self.report_path, self.fileName)
        wb.save(filename = self.xlsfile)

    def comp_base(self):
        # 基礎設定
        # name, width, sql_column_name
        lis_base = []; a = lis_base.append
        a('產品工程名稱,  36, pm002') 
        a('品號,   12, pm003')
        a('品名,   20, pm004') 
        a('規格,   12, pm005') 
        a('預設工程,      5, pm006') 
        a('換校模時間(分),     6, pm008')
        a('換校模時間(秒),     6, pm016')
        a('加工批量(1模做幾個), 6, pm009')
        a('換料時間(分),       6, pm010')
        a('換料時間(秒),       6, pm017')
        a('加工時間(分),      6, pm011')
        a('加工時間(秒),      6, pm018')
        a('量測時間比率分子,   6, pm021')
        a('量測時間比率分母,   6, pm022')
        a('換刀時間比率分子,   6, pm023')
        a('換刀時間比率分母,   6, pm024')
        a('等待時間比率分子,   6, pm025')
        a('等待時間比率分母,   6, pm026')
        a('時間成本PCS/分,    12, pm019')
        a('備註,     12,  pm020')
        a('最後編輯日期,     19, pm013')
        a('最後編輯人員,     12, pm014')
        lis_e1, lis_e2, lis_e3 = [],[],[]
        for e in lis_base:
            [e1, e2, e3]= e.split(',')
            lis_e1.append(e1.strip())
            lis_e2.append(int(e2.strip()))
            lis_e3.append(e3.strip())
        self.xls_index =dict(zip(lis_e1, [e for e in range(1,len(lis_e1)+1)]))
        self.xls_width =dict(zip(lis_e1, lis_e2))
        self.xls_sqlcn =dict(zip(lis_e1, lis_e3))

    def output(self):
        if True: # style, func
            f10g =font_10_Calibri_g
            f10 = font_10_Calibri
            f11 = font_11_Calibri
            f11gr = font_11_Calibri_green
            # fill color
            fill_color_bn061 = {
                0: cf_none, # 0未開始
                1: cf_none, # 1準備中
                2: cf_none, # 2準備好
                3: cf_none, # 3校模中
                4: cf_green, # 4生產中
                5: cf_blue} # 5已完工

            fill_color_bn066 = {
                1: cf_purple} #  1插單

            fill_color_bn062 = {
                1: cf_red} #  1.異常停機

            # func, method
            write=self.c_write; fill=self.c_fill; comm=self.c_comm; img=self.c_image2; column_w=self.c_column_width

        wb = self.wb
        # 第1部分 機台排程狀況

        df_pm = self.db.get_dp_pm()
        x_index = self.xls_index
        x_width = self.xls_width
        x_sqlcn = self.xls_sqlcn
        for dpi, dpr in self.df_dp.iterrows():
            sh = wb[dpr['dp002']]
            super().__init__(self.xlsfile, wb, sh) # 傳遞引數給父class
            cr=1; column_w(list(x_width.values())) # 設定欄寬
            for name, index in x_index.items():
                write(cr, index, name, f11, alignment=ah_wr, border=bt_border, fillcolor=cf_gray) # 欄位名稱

            df_w = df_pm.loc[df_pm['pm015']==dpr['dp001']]
            for i, r in df_w.iterrows():
                cr+=1
                for name, index in x_index.items():
                    scn = x_sqlcn[name] # sql_column_name
                    write(cr, index, r[scn], f11)

def test1():
    fileName = 'pmg01' + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    Report_pmg01(fileName)
    print('ok')

if __name__ == '__main__':
    test1()
    sys.exit() #正式結束程式