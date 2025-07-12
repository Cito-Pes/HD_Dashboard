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


def TM_QTY():

    try:
        conn = pyodbc.connect(ns_conn)
        cursor = conn.cursor()

        FromDate = 'CONVERT(VARCHAR(8),dateadd(MONTH,-3,GETDATE()),112 )'
        ToDate = 'CONVERT(VARCHAR(8),GETDATE(),112)'

        SQL_QTY = f"""
                    SELECT st.SaName
                        , ISNULL(sum(x.cnt1),0) as 접수cnt
                        , ISNULL(sum(x.cnt2),0) as 정상cnt
                        , ISNULL(sum(x.cnt3),0) as 만기cnt
                        , ISNULL(sum(x.cnt4),0) as 행사cnt 
                        , ISNULL(sum(x.cnt1),0)+ISNULL(sum(x.cnt2),0)+ISNULL(sum(x.cnt3),0)+ISNULL(sum(x.cnt4),0) as cnt
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
                        AND me.MemType = '정상'
                        AND REPLACE(me.Rec_Date,'-','') >=  {FromDate}
                        AND REPLACE(me.Rec_Date,'-','') <= {ToDate}
                    GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x2 ON st.SaBun = x2.Charge_IDP
                    union all
                    (SELECT me.Charge_IDP, me.MemberNo, 0 AS cnt1, 0 AS cnt2, CASE gu.G_etc_str5 WHEN 4 THEN count(me.ID)*0.25 WHEN 2 THEN count(me.ID)*0.5 END AS cnt3, 0 AS cnt4
                    FROM [Member] me INNER JOIN goods gu ON me.Goods  = gu.Goods_ID
                    WHERE me.EventType in ('여행','크루즈') 
                        AND me.MemType IN ('만기')
                        AND REPLACE(me.Rec_Date,'-','') >=  {FromDate}
                        AND REPLACE(me.Rec_Date,'-','') <= {ToDate}
                    GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x3 ON st.SaBun = x3.Charge_IDP
                    union all
                    (SELECT me.Charge_IDP, me.MemberNo, 0 AS cnt1, 0 AS cnt2, 0 AS cnt3, CASE gu.G_etc_str5 WHEN 4 THEN count(me.ID)*0.25 WHEN 2 THEN count(me.ID)*0.5 END AS cnt4
                    FROM [Member] me INNER JOIN goods gu ON me.Goods  = gu.Goods_ID
                    WHERE me.EventType in ('여행','크루즈') 
                        AND me.MemType IN ('행사')
                        AND REPLACE(me.Rec_Date,'-','') >=  {FromDate}
                        AND REPLACE(me.Rec_Date,'-','') <= {ToDate}
                    GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5)) x ON st.SaBun = x.Charge_IDP
                    GROUP BY st.SaName	
                    order by st.SaName
                    """
        print(SQL_QTY)
        cursor.execute(SQL_QTY)
        QTY_Result = cursor.fetchall()
        print(QTY_Result)
    except Exception as Err_Search_Pay:
        print(Err_Search_Pay)
        # self.ERR_MSG("Search_Pay 오류", Err_Search_Pay, bg_color="yellow", text_color="red")
    finally:
        conn.close()


TM_QTY()