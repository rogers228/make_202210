if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['make_202210'] # 目前路徑
    sys.path.append(config_path)

import pandas as pd
import pyodbc
from config import *

class db_yst(): #讀取excel 單一零件
    def __init__(self):
        self.cn = pyodbc.connect(config_conn_YST) # connect str 連接字串\

    def get_pd_test(self, pdno):
        s = "SELECT MB001,MB002,MB003 FROM INVMB WHERE MB001 = '{0}'"
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def is_exist_pd(self, pdno): # 品號是否存在
        s = "SELECT MB001 FROM INVMB WHERE MB001 = '{0}'"
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return (len(df.index)>0)

    def get_pur_ma002(self, ma001): # 供應商簡稱
        s = "SELECT MA002 FROM PURMA WHERE MA001 = '{0}'"
        s = s.format(ma001)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['MA002'] if len(df.index) > 0 else ''

    def get_pd_one_to_dic(self, pdno):
        s = """
        SELECT TOP 1
        RTRIM(MB001) AS MB001,MB002,MB003,MB025,MB032,MB050,MB010,MB011,MB053
        FROM INVMB
        WHERE MB001 = '{0}'
        """
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0].to_dict() if len(df.index) > 0 else None

    def get_bom(self, pdno):
        s = """
            SELECT RTRIM(BOMMD.MD001) AS MD001,
                BOMMD.MD002,
                BOMMD.MD003,
                BOMMD.MD006,
                BOMMD.MD007,
                INVMB.MB002,
                INVMB.MB003,
                INVMB.MB025,
                INVMB.MB032,
                INVMB.MB050,
                INVMB.MB010,
                INVMB.MB011,
                INVMB.MB053
            FROM BOMMD
                LEFT JOIN INVMB ON BOMMD.MD003 = INVMB.MB001
            WHERE MD001 = '{0}'
            ORDER BY MD002
            """
        s = s.format(pdno)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def get_bmk(self, mf001, mf002):
        s = """
            SELECT RTRIM(MF001) AS MF001,MF002,MF003,MF004,MF005,MF006,MF007,MF008,MF022,MF019,MF009,MF024,MF017,MF018,MF023,
            MW002
            FROM BOMMF
                LEFT JOIN CMSMW ON BOMMF.MF004 = CMSMW.MW001
            WHERE MF001 = '{0}' AND MF002 = '{1}'
            ORDER BY MF003
            """
        s = s.format(mf001, mf002)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def wget_imd(self, pdno_arr=''):
        # 品號單位換算
        # pdno_arr 品號 文字陣列
        if pdno_arr == "":
            return None
        pdno_arr = str(pdno_arr).replace(' ','') # 去除空格
        pdno_inSTR = "('" + "','".join(pdno_arr.split(',')) + "')"
        s = """
            SELECT RTRIM(MD001) AS MD001,RTRIM(MD002) AS MD002,MD003,MD004 FROM INVMD
            WHERE MD001 IN {0}
            ORDER BY MD001, MD002 ASC
            """
        s = s.format(pdno_inSTR)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

def test1():
    db = db_yst()
    # df = db.get_bom('4A302001')
    # df = db.get_bmk('4A302001','01')
    # print(df)
    # print(db.get_pur_ma002('1020010'))
    print(db.is_exist_pd('4A3060010'))

if __name__ == '__main__':
    test1()        