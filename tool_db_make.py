if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['make_202210'] # 目前路徑
    sys.path.append(config_path)
    
import pandas as pd
# import pyodbc
from sqlalchemy.engine import URL
from sqlalchemy import create_engine
from config import *

class db_make(): #讀取excel 單一零件
    def __init__(self):
        # self.cn = pyodbc.connect(config_conn_MAKE) 
        self.cn = create_engine(URL.create('mssql+pyodbc', query={'odbc_connect': config_conn_MAKE})).connect()

    def get_bn_df(self, br018, bn007): #所有排程資料
        # br018 部門ID
        # bn007 預計開始
        s = """
            SELECT
                bn001,bn002,bn003,bn004,bn005,bn006,bn007,bn008,bn010,bn023,bn042,bn044,bn045,bn061,bn062,bn066,bn070,bn071,
                br002,br003,br004,br005,br006,br007,br008,br009,br010,br013,br014,br018,br019,br025,br026,
                pm002,
                bt003,bt004,bt005,bt006,
                (br002 +'-' + br003) AS sbr003,
                (bt004 +'-' + bt005) AS sbt004
            FROM
                (((rec_bn INNER JOIN rec_br ON rec_bn.bn003 = rec_br.br001) INNER JOIN rec_ma ON rec_bn.bn002 = rec_ma.ma001) INNER JOIN rec_pm ON rec_bn.bn004 = rec_pm.pm001) LEFT JOIN rec_bt ON rec_bn.bn003 = rec_bt.bt002
            WHERE
                rec_br.br010 = 1 AND
                rec_br.br018 = {0} AND
                rec_bn.bn007 > '{1}'
            ORDER BY bn005
            """
            # br010 = 1 已派工
        s = s.format(br018, bn007)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def get_fbnr_df(self, br018): #所有尚未排程 的派工資料)
        s = """
            SELECT br001,br002,br003,br004,br005,br006,br007,br008,br009,br010,br013,br014,br015,br018,br019,br023
            FROM rec_br
            WHERE
                br010 = 1 AND
                br018 = {0} AND
                (br023 Is Null Or br023 = '')
            ORDER BY br011
            """
            # br010 = 1 已派工
            # (br023 Is Null Or br023 = '') 尚未排程

        s = s.format(br018)
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def get_dp_pm(self): #所有工程
        s = """
        SELECT 
            pm001,pm002,pm003,pm004,pm005,pm006,pm007,pm008,pm009,pm010,
            pm011,pm012,pm013,pm014,pm015,pm016,pm017,pm018,pm019,pm020,
            pm021,pm022,pm023,pm024,pm025,pm026,pm027
        FROM rec_pm
        ORDER BY pm004 
        """
        df = pd.read_sql(s, self.cn)
        return df if len(df.index) > 0 else None

    def get_dp_df(self): #所有部門
        s = "SELECT dp001,dp002,dp003 FROM rec_dp ORDER BY dp002"
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def get_ma_df(self, ma004): #所有機台 
        # ma004 部門id
        s = "SELECT ma001,ma002,ma003,ma008 FROM rec_ma WHERE ma004 = {0} ORDER BY ma008"
        s = s.format(ma004)
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None

    def get_sy002(self): #顯示日期
        # ma004 部門id
        s = "SELECT sy002 FROM rec_sy WHERE sy001=1"
        df = pd.read_sql(s, self.cn) #轉pd
        return df.iloc[0]['sy002'].strftime('%Y-%m-%d') if len(df.index) > 0 else ''

def test1():
    # new id
    mk = db_make()
    # df = mk.get_ma_df(1)
    # df = mk.get_bn_df(1, '2022-09-01')
    # df1 = df[['bn001','bt004','bt005']]
    # df1 = df[['sbr003','br002','br003']]
    # pd.set_option('display.max_rows', df1.shape[0]+1) # 顯示最多列
    # pd.set_option('display.max_columns', None) #顯示最多欄位    
    # print(df1)

    df = mk.get_dp_df()
    print(df)

    # d1 =mk.get_sy002()
    # print(d1)
    # print(type(d1))

if __name__ == '__main__':
    test1()        
    print('ok')