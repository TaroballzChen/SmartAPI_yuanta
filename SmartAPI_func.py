import time
import xlwings as xw
import numpy as np
import datetime
import os
import subprocess
import sys
import glob

PATH = os.getcwd()+'/'
##############  參數 ##########################
product_id = "UDFC9"      #交易商品的代碼
K_bar = int(sys.argv[1])  #策略
data_length = 120   #使用多少array空間進行儲存指標，越大值會越準但是效能會越低
DDE_file = ""   # DDE報價檔案位置
Order = PATH+'SmartAPI/Order.exe'           #下單機位置
Report = PATH+"SmartAPI/GetAccount.exe"     #查詢機位置
ChangeProd = PATH+"SmartAPI/ChangeProdid.exe"  #切換商品
check_unequal = PATH+"SmartAPI/OnOpenInterest.exe"  #查詢未平倉
###################################################

class SmartAPI():     # SmartAPI調用
    buy_option = ["B","0","1","MKT","IOC","1"]
    sell_option = ["S", "0", "1", "MKT", "IOC", "1"]
    def __init__(self):
        self.askcode = ""
        self.OrderAPI = [Order,product_id]
        self.ChangeProd = [ChangeProd,product_id]
        self.nonequalAPI = [check_unequal]
        self.buy_point = 0
        self.sell_point = 0


    def LINE_notify(self,msg):
        return subprocess.Popen("python notify.py %s" %msg, shell=True)

    def changeproduct(self):
        status = subprocess.check_output(self.ChangeProd).decode("big5").strip()
        time.sleep(2)
        if status == "ChangeSucess":
            print("切換商品代號成功：%s" % product_id)
            return True
        else:
            time.sleep(0.5)
            return False

    def Check_UnEual(self,simulation=False):
        if not simulation:
            for _ in range(3):
                if not self.changeproduct():
                    continue
                result = subprocess.check_output([check_unequal]).decode("big5").strip()
                if result:
                    B_or_S = result.split(",")[1]
                    if B_or_S in ["B","S"]:
                        return "B" if B_or_S =="B"else "S"
                    else:
                        continue
                else:
                    print("沒有未平倉的單")
                    return False
        else:
            B_or_S = SIMULATION_UnEqual
            if B_or_S in ["B","S"]:
                return "B" if B_or_S =="B" else "S"
            else:
                print("沒有未平倉的單")
                return False

    def order(self,time_value,buy=True,simulation=False,simulation_price=None,sublot=None,subpen=None,subamount=None,criticalprice=None):
        if not simulation:
            for error_times in range(3):
                if not self.changeproduct():
                    continue
                if buy:
                    self.askcode = subprocess.check_output(self.OrderAPI + self.buy_option).decode("big5").strip()
                    op = "多單"
                else:
                    self.askcode = subprocess.check_output(self.OrderAPI + self.sell_option).decode("big5").strip()
                    op = "空單"
                time.sleep(3)
                print("委託代號：%s"%self.askcode,"時間：",time_value)
                price = self.check_trade_status()
                if price:
                    print("%s 成交"%op,"點數：%d"%price)
                    p = self.LINE_notify("%d策略，下%s，時間：%s，價格：%d，成交量：%d，關鍵價格：%d" % (K_bar,op,time_value, price,subamount,criticalprice))
                    if buy:
                        self.buy_point = price
                    else:
                        self.sell_point = price
                    p.wait(10)
                    return price
                else:
                    print("交易失敗 重新交易 嘗試第%d次 時間：%s"%(error_times+1,time_value))
            else:
                print("交易未曾成功，將退出系統 時間：%s"%time_value)
                p = self.LINE_notify("%d策略，交易未曾成功，將退出系統\n時間：%s"%(K_bar,time_value))
                p.wait(10)
                sys.exit()
        else:
            global SIMULATION_UnEqual
            if buy:
                print("下多單：時間",time_value,"價格：",simulation_price)
                p = self.LINE_notify(
                    "[模擬%d策略]，下多單，時間：%s，價格：%d，成交量：%d，關鍵價格：%d" % (K_bar, time_value, simulation_price, subamount,criticalprice))
                self.buy_point = float(simulation_price)
                if SIMULATION_UnEqual == 0:
                    SIMULATION_UnEqual ="B"
                elif SIMULATION_UnEqual == "S":
                    SIMULATION_UnEqual = 0
            else:
                print("下空單：時間",time_value,"價格：",simulation_price)
                p = self.LINE_notify(
                    "[模擬%d策略]，下空單，時間：%s，價格：%d，成交量：%d，關鍵價格：%d" % (K_bar, time_value, simulation_price, subamount,criticalprice))
                self.sell_point = float(simulation_price)
                if SIMULATION_UnEqual == 0:
                    SIMULATION_UnEqual ="S"
                elif SIMULATION_UnEqual == "B":
                    SIMULATION_UnEqual = 0
            p.wait(10)

    def check_trade_status(self):
        for error_times in range(5):
            check_result = subprocess.check_output([Report,self.askcode]).decode("big5").strip().split(',')
            try:
                status = check_result[1]
                price = check_result[4]
                if status == "委託失敗":
                    print("委託失敗，重新查詢")
                    time.sleep(1)
                    continue
                elif status == "全部成交":
                    print("全部成交")
                    return float(price)
                else:
                    if error_times ==4:
                        print("交易系統錯誤，緊急退出，或許以產生交易")
                        p = self.LINE_notify("%d策略，交易系統錯誤，緊急退出，或許以產生交易" %K_bar)
                        p.wait(10)
                        sys.exit(1)
                    time.sleep(1)
                    continue

            except Exception:
                time.sleep(1)
                continue