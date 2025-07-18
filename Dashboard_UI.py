#!/usr/bin/env python
# -*- coding: utf-8 -*-

import pyodbc # MS SQL DB 연결을 위한 라이브러리

import os
import sys
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout,
    QPushButton, QLineEdit, QTableWidget, QTableWidgetItem,
    QTextBrowser, QWidget, QFileDialog, QMessageBox
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt
import pandas as pd
from datetime import datetime
import math
from shiboken6 import Shiboken
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

global ns_conn
f = open(resource_path("./config.txt"), 'r', encoding='utf-8')
lines = f.readlines()
for line in lines:
    line = line.strip()  # 줄 끝의 줄 바꿈 문자를 제거한다.
    # print(line, line[:9])
    if line[:9] == "ns_conn =":
        ns_conn = line[9:]

    if line[:11] == "app_title =":
        App_Title = line[11:]

    if line[:9] == "app_ver =":
        App_Ver = line[9:]

font_path = resource_path("./static/D2Coding.ttc")
font_name = fm.FontProperties(fname = font_path).get_name()
plt.rc('font', family = font_name)

def TM_QTY():

    try:
        conn = pyodbc.connect(ns_conn)
        cursor = conn.cursor()

        FromDate = '20250707' #CONVERT(VARCHAR(8),dateadd(MONTH,-3,GETDATE()),112 )'
        ToDate = 'CONVERT(VARCHAR(8),GETDATE(),112)'

        SQL_QTY = f"""
                    SELECT st.SaName
                        , ISNULL(sum(x.cnt1),0) as 접수cnt
                        , ISNULL(sum(x.cnt2),0) as 정상cnt
                        , ISNULL(sum(x.cnt1),0)+ISNULL(sum(x.cnt2),0) as cnt
                    from (SELECT SaBun, SaName FROM Staff WHERE PlaceofDuty ='홈쇼핑 TM' AND BranchOffice='TM1' AND OutDate ='') st 
                    LEFT JOIN
                    (
                    (SELECT me.Charge_IDP, me.MemberNo, CASE gu.G_etc_str5 WHEN 4 THEN count(me.ID)*0.25 WHEN 2 THEN count(me.ID)*0.5 END AS cnt1,0 AS cnt2,0 AS cnt3, 0 AS cnt4
                    FROM [Member] me INNER JOIN goods gu ON me.Goods  = gu.Goods_ID
                    WHERE me.EventType in ('여행','크루즈') 
                        AND me.MemType ='접수'
                        AND REPLACE(me.Rec_Date,'-','') >=  {FromDate}
                        AND REPLACE(me.Rec_Date,'-','') <= {ToDate}
                    GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x1 ON st.SaBun = x1.Charge_IDP
                    union all
                    (SELECT me.Charge_IDP, me.MemberNo, 0 AS cnt1, CASE gu.G_etc_str5 WHEN 4 THEN count(me.ID)*0.25 WHEN 2 THEN count(me.ID)*0.5 END AS cnt2,0 AS cnt3, 0 AS cnt4
                    FROM [Member] me INNER JOIN goods gu ON me.Goods  = gu.Goods_ID
                    WHERE me.EventType in ('여행','크루즈') 
                        AND me.MemType in ('정상','만기','행사')
                        AND REPLACE(me.Rec_Date,'-','') >=  {FromDate}
                        AND REPLACE(me.Rec_Date,'-','') <= {ToDate}
                    GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x2 ON st.SaBun = x2.Charge_IDP
                    ) x ON st.SaBun = x.Charge_IDP
                    GROUP BY st.SaName	
                    order by st.SaName
                    """
        print(SQL_QTY)
        cursor.execute(SQL_QTY)
        QTY_Result = cursor.fetchall()
        # print(QTY_Result)

        Emp = []
        Cnt1 = []
        Cnt2 = []
        # Cnt3 = []
        # Cnt4 = []
        SumCnt = []

        for Q in QTY_Result:
            Emp.append(Q[0])
            Cnt1.append(Q[1])
            Cnt2.append(Q[2])
            # Cnt3.append(Q[3])
            # Cnt4.append(Q[4])
            SumCnt.append(Q[3])


        Qty_Dict = {
            "매니져" : Emp,
            "접수" : Cnt1,
            "정상" : Cnt2,
            # "만기" : Cnt3,
            # "행사" : Cnt4,
            "합" : SumCnt
        }
        PD_Qty = pd.DataFrame(Qty_Dict)

        max_pd = PD_Qty.loc[PD_Qty["합"].idxmax()]
        # 각각의 변수에 담기
        MAX_EMP = max_pd["매니져"]
        MAX_QTY = max_pd["합"]

        return PD_Qty

    except Exception as Err_Search_Pay:
        print(Err_Search_Pay)
        # self.ERR_MSG("Search_Pay 오류", Err_Search_Pay, bg_color="yellow", text_color="red")
        return "Error"
    finally:
        conn.close()

        print(PD_Qty)
        # dic = {y: x for x, y in zip(x, y)}
        plt.figure(figsize=(6, 4))
        plt.title("000 계약현황")
        plt.xlabel("상담 매니져")
        plt.ylabel("채결 구좌수")
        plt.bar(PD_Qty['매니져'], PD_Qty['접수'], label = "접수")
        plt.bar(PD_Qty['매니져'], PD_Qty['정상'], bottom=PD_Qty['접수'], label = "계약")
        # plt.bar(PD_Qty['매니져'], PD_Qty['만기'], bottom=PD_Qty['접수']+PD_Qty['정상'], label = "만기")
        # plt.bar(PD_Qty['매니져'], PD_Qty['행사'], bottom=PD_Qty['접수']+PD_Qty['정상']+PD_Qty['만기'], label="행사")
        # plt.text(PD_Qty['매니져'],MAX_QTY,'%.1f' % 20, ha='center', size = 12)
        plt.legend(ncol = 2)
        # plt.tick_params(axis='y', direction='inout')
        plt.grid(True, axis='y', color = 'lightgray')
        plt.ylim(0, MAX_QTY + 3)

        for i, v in enumerate(PD_Qty['매니져']):
            plt.text(v, PD_Qty['합'][i], PD_Qty['합'][i],  # 좌표 (x축 = v, y축 = y[0]..y[1], 표시 = y[0]..y[1])
                     fontsize=12,
                     color='blue',
                     horizontalalignment='center',  # horizontalalignment (left, center, right)
                     verticalalignment='bottom')  # verticalalignment (top, center, bottom)
        plt.show()

def TM_Allowance():

    try:
        conn = pyodbc.connect(ns_conn)
        cursor = conn.cursor()

        SQL_Allow = """
                    select x.SaName, sum(x.xx)*0.25 as ETC from (
                    select me.TotPay , gu.Cash_Month, me.id, me.name, me.reg_date, st.SaName, st.SaBun, case gu.G_etc_str5 when 4 then 1 when 2 then 2 end xx
                    from member me
                    inner join staff st on me.Charge_IDP = st.SaBun
                    inner join goods gu on me.goods = gu.Goods_ID and gu.Cash>0 and gu.Goods_ID not like 'WEP%'
                    left join Allowance_DT a on me.id = a.id
                    where st.PlaceofDuty='홈쇼핑 TM'
                    and me.MemType in ('정상','만기','행사')
                    and me.TotPay / gu.Cash_Month >= 2
                    and me.TotPay > 0 
                    and a.id is null
                    and me.reg_date >='2024-01-01'
                    union all
                    select me.TotPay , gu.Cash_Month, me.id, me.name, me.reg_date, st.SaName, st.SaBun, case gu.G_etc_str5 when 4 then 1 when 2 then 2 end xx
                    from member me
                    inner join staff st on me.Charge_IDP = st.SaBun
                    inner join goods gu on me.goods = gu.Goods_ID and gu.Cash>0 and gu.Goods_ID like 'WEP%'
                    left join Allowance_DT a on me.id = a.id
                    where st.PlaceofDuty='홈쇼핑 TM'
                    and me.MemType in ('정상','만기','행사')
                    and me.TotPay > 0 
                    and a.id is null
                    and me.reg_date >='2024-01-01'
                    ) x
                    group by x.SaName
                 """
        # print(SQL_QTY)
        cursor.execute(SQL_Allow)
        QTY_Allow = cursor.fetchall()
        print(QTY_Allow)
        Emp = []
        Cnt1 = []
        for Q in QTY_Allow:
            Emp.append(Q[0])
            Cnt1.append(Q[1])

        Qty_Dict = {
            "매니져": Emp,
            "구좌수": Cnt1}
        PD_Qty = pd.DataFrame(Qty_Dict)

        return PD_Qty

    except Exception as Err_Allow:
        print(Err_Allow)
    finally:
        conn.close()


        print(PD_Qty)
        max_pd = PD_Qty.loc[PD_Qty["구좌수"].idxmax()]
        # 각각의 변수에 담기
        MAX_EMP = max_pd["매니져"]
        MAX_QTY = max_pd["구좌수"]


        plt.figure(figsize=(6, 4))
        plt.title("000 수수료 지급예정 구좌")
        plt.xlabel("상담 매니져")
        plt.ylabel("구좌수")
        plt.bar(PD_Qty['매니져'], PD_Qty['구좌수'], label = "구좌수")
        plt.legend(ncol = 2)
        plt.grid(True, axis='y', color = 'lightgray')
        plt.ylim(0, MAX_QTY+3)
        for i, v in enumerate(PD_Qty['매니져']):
            plt.text(v, PD_Qty['구좌수'][i], PD_Qty['구좌수'][i],  # 좌표 (x축 = v, y축 = y[0]..y[1], 표시 = y[0]..y[1])
                     fontsize=12,
                     color='blue',
                     horizontalalignment='center',  # horizontalalignment (left, center, right)
                     verticalalignment='bottom')  # verticalalignment (top, center, bottom)
        plt.show()


print(TM_QTY())
# print(TM_Allowance())

#https://tnqkrdmssjan.tistory.com/55

