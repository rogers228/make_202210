if True: # 固定引用開發環境 或 發佈環境 的 路徑
    import os, sys, custom_path
    config_path = os.getcwd() if os.getenv('COMPUTERNAME')=='VM-TESTER' else custom_path.custom_path['make_202210'] # 目前路徑
    sys.path.append(config_path)

import time
import click
import tool_auth
import rpt_bng01, rpt_pmg01

@click.command() # 命令行入口
@click.option('-report_name', help='report name', required=True, type=str) # required 必要的
@click.option('-depart', help='department id.', type=int) 
def main(report_name, depart=''):
    au = tool_auth.Authorization()
    if not any([au.isqs(1), au.isqs(2)]): # 權限 1啟用派工作業 or 2啟用排程作業
        click.echo('無權限!')
        return # 無權限 退出

    global fileName; fileName = report_name + '_' + time.strftime("%Y%m%d%H%M%S", time.localtime()) + '.xlsx'
    global departId; departId = depart
    dic = {'bng01': bng01,
           'pmg01': pmg01,
          }
    func = dic.get(report_name, None)
    if func is not None:
        func()

def bng01(): # 製造排程表
    rpt_bng01.Report_bng01(fileName, departId)

def pmg01(): # 工程資料明細表
    rpt_pmg01.Report_pmg01(fileName)

if __name__ == '__main__':
    main()
    # cmd
    # C:\python_venv\python.exe \\220.168.100.104\pdm\python_program\make_202210\rpt_main.py -report_name bng01 -depart 1
