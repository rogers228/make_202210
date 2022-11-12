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

    # def test_to_df(self):
    #     s = """
    #         SELECT
    #             br002,br003,
    #             SUBSTRING(br024,1,4) AS a1,
    #             br004,ta.TA024,
    #             SUBSTRING(br024,11, LEN(br024)-10) AS a3, MD002,
    #             br008,
    #             convert(varchar, br011, 112) AS br011,
    #             CASE 
    #                 WHEN br010 = 0 THEN '0.未派工'
    #                 WHEN br010 = 1 THEN '1.已派工'
    #                 ELSE ''
    #             END AS br010,
    #             br012,
    #             CASE 
    #                 WHEN br015 = 1 THEN '1.未排程'
    #                 WHEN br015 = 2 THEN '2.已排程'
    #                 WHEN br015 = 3 THEN '3.生產中'
    #                 WHEN br015 = 4 THEN '4.已完工'
    #                 WHEN br015 = 5 THEN '5.待排程'
    #                 WHEN br015 = 6 THEN '6.建議外包'
    #                 ELSE ''
    #             END AS br015
    #         FROM YEOSHE_MAKE.dbo.rec_br
    #             LEFT JOIN YST.dbo.CMSMD ON SUBSTRING(br024,11,LEN(br024)-10) = MD001 
    #             LEFT JOIN YST.dbo.SFCTA as ta ON br002 = ta.TA001 AND br003 = ta.TA002 AND SUBSTRING(br024,1,4)=ta.TA003

    #         WHERE
    #             br002 = {0} AND
    #             br003 = {1}
    #         ORDER BY SUBSTRING(br024,6,4),br003
    #         """
    #     s = s.format('5101', '20220418001')
    #     df = pd.read_sql(s, self.cn) #轉pd
    #     return df if len(df.index) > 0 else None

    def test_to_df(self):
        s = """
            SELECT
                ma008,ma003,pm002,
                SUBSTRING(br024,1,4) AS a1,br004,ta.TA024
                bn041,bn042,bn043,bn065,
                convert(varchar, bn012, 112) AS bn012,
                convert(varchar, bn044, 112) AS bn044,
                convert(varchar, bn045, 112) AS bn045,
                CASE 
                    WHEN bn061 = 0 THEN '0.未開始'
                    WHEN bn061 = 1 THEN '1.開始準備'
                    WHEN bn061 = 2 THEN '2.完成準備'
                    WHEN bn061 = 3 THEN '3.開始校模(校模中)'
                    WHEN bn061 = 4 THEN '4.開始加工(生產中)'
                    WHEN bn061 = 5 THEN '5.結束(完工)'
                    ELSE ''
                END AS bn061,
                bn070,bn071
            FROM YEOSHE_MAKE.dbo.rec_bn
                LEFT JOIN YEOSHE_MAKE.dbo.rec_ma ON bn002=ma001
                LEFT JOIN YEOSHE_MAKE.dbo.rec_br ON bn003=br001
                LEFT JOIN YEOSHE_MAKE.dbo.rec_pm ON bn004=pm001
                LEFT JOIN YST.dbo.SFCTA as ta ON br002 = ta.TA001 AND br003 = ta.TA002 AND SUBSTRING(br024,1,4)=ta.TA003

            WHERE
                br002 = {0} AND
                br003 = {1}
            ORDER BY SUBSTRING(br024,6,4), bn012
            """
        s = s.format('5101', '20220418001')
        df = pd.read_sql(s, self.cn) #轉pd
        return df if len(df.index) > 0 else None
def test1():
    # new id
    mk = db_make()
    # df = mk.get_ma_df(1)
    # df = mk.get_bn_df(1, '2022-09-01')
    df = mk.test_to_df()
    # df1 = df[['br004','br024','a1','a2','MW002','a3','MD002']]
    pd.set_option('display.max_rows', df.shape[0]+1) # 顯示最多列
    pd.set_option('display.max_columns', None) #顯示最多欄位    
    print(df)



    # d1 =mk.get_sy002()
    # print(d1)
    # print(type(d1))

if __name__ == '__main__':
    test1()        
    print('ok')