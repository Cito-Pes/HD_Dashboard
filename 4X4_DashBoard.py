#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys,os
import time
import threading
import pyodbc

from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import (
    QApplication, QWidget, QGridLayout, QLabel
)
from PySide6.QtGui import QIcon, QBrush
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

global ns_conn, App_Title, App_Ver
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

global TDate, WDate, MDate
TDate=""
WDate=""
MDate=""

font_path = resource_path("./static/D2Coding.ttc")
font_name = fm.FontProperties(fname = font_path).get_name()
plt.rc('font', family = font_name)

def get_db_connection():
    """MSSQL 연결 생성."""
    return pyodbc.connect(ns_conn, timeout=5)

def get_DATE_SET():
    global TDate, WDate, MDate
    conn = get_db_connection()
    cursor = conn.cursor()
    SQL_DATE = """
                            SELECT
                            CONVERT(VARCHAR, GETDATE(), 112) AS 오늘,
                            CONVERT(VARCHAR, DATEADD(wk,DATEDIFF(wk,7,GETDATE()),0), 112) AS 이번주_첫_월요일,
                            CONVERT(VARCHAR, DATEADD(mm,DATEDIFF(mm,0,GETDATE()),0), 112) AS 이번달_첫날;    
                            """
    cursor.execute(SQL_DATE)
    Date_SET = cursor.fetchone()
    TDate = Date_SET[0]
    WDate = Date_SET[1]
    MDate = Date_SET[2]
    print("Date_SET", TDate, WDate, MDate)
    cursor.close()
    conn.close()
    return TDate, WDate, MDate

# 2. 쓰레드 클래스 정의
class QueryWorker(QThread):
    data_ready = Signal(object)  # 데이터를 UI에 보낼 때 사용

    def __init__(self, sql_query: str, interval_sec: int, parent=None):
        super().__init__(parent)
        self.query = sql_query
        self.interval = interval_sec
        self._running = True

    def run(self):
        conn = get_db_connection()
        cursor = conn.cursor()
        while self._running:
            cursor.execute(self.query)
            rows = cursor.fetchall()
            self.data_ready.emit(rows)
            time.sleep(self.interval)
        cursor.close()
        conn.close()

    def stop(self):
        self._running = False


# 3. 메인 윈도우 정의
class DashboardWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(App_Title+"   "+App_Ver)
        self.setWindowIcon(QIcon(resource_path("./static/logo_icon.png")))  # 로고 파일 경로 설정 필요
        self.resize(1920, 1080)
        self._setup_ui()
        self._start_workers()

    def _setup_ui(self):
        layout = QGridLayout(self)

        # 4개의 캔버스를 생성
        self.canvas_daily = self._create_canvas("일간 실적")
        self.canvas_weekly = self._create_canvas("주간 실적")
        self.canvas_monthly = self._create_canvas("월간 실적")
        self.canvas_fees = self._create_canvas("수수료 정산 실적")


        # 그리드 배치 (2행 2열)
        layout.addWidget(self.canvas_daily,   0, 0)
        layout.addWidget(self.canvas_weekly,  0, 1)
        layout.addWidget(self.canvas_monthly, 1, 0)
        layout.addWidget(self.canvas_fees,    1, 1)

        self.setLayout(layout)

    def _create_canvas(self, title: str) -> FigureCanvas:
        """Matplotlib 캔버스 생성 및 타이틀 설정."""
        fig = Figure(figsize=(9.6, 5.4), dpi=100)  # 1920x1080 해상도 기준
        ax = fig.add_subplot(111)
        ax.set_title(title)
        canvas = FigureCanvas(fig)
        canvas.ax = ax
        return canvas

    def Make_SQL(self, Type: str):
        global TDate, WDate, MDate

        self.Type = Type

        TDate, WDate, MDate = get_DATE_SET()

        print("SET", TDate, WDate, MDate)


        if self.Type == 'DAY':
            self.FromDate = TDate #'20250718'  # CONVERT(VARCHAR(8),dateadd(MONTH,-3,GETDATE()),112 )'
        elif self.Type == 'WEEK':
            self.FromDate = WDate #'20250714'  # CONVERT(VARCHAR(8),dateadd(MONTH,-3,GETDATE()),112 )'
        elif self.Type =='MONTH':
            self.FromDate = MDate #'20250607'  # CONVERT(VARCHAR(8),dateadd(MONTH,-3,GETDATE()),112 )'


        self.ToDate = TDate #'CONVERT(VARCHAR(8),GETDATE(),112)'

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
                                AND REPLACE(me.Rec_Date,'-','') >=  '{self.FromDate}'
                                AND REPLACE(me.Rec_Date,'-','') <= '{self.ToDate}'
                            GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x1 ON st.SaBun = x1.Charge_IDP
                            union all
                            (SELECT me.Charge_IDP, me.MemberNo, 0 AS cnt1, CASE gu.G_etc_str5 WHEN 4 THEN count(me.ID)*0.25 WHEN 2 THEN count(me.ID)*0.5 END AS cnt2,0 AS cnt3, 0 AS cnt4
                            FROM [Member] me INNER JOIN goods gu ON me.Goods  = gu.Goods_ID
                            WHERE me.EventType in ('여행','크루즈') 
                                AND me.MemType in ('정상','만기','행사')
                                AND REPLACE(me.Rec_Date,'-','') >=  '{self.FromDate}'
                                AND REPLACE(me.Rec_Date,'-','') <= '{self.ToDate}'
                            GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5) --x2 ON st.SaBun = x2.Charge_IDP
                            ) x ON st.SaBun = x.Charge_IDP
                            GROUP BY st.SaName	
                            order by st.SaName
                                    """
        return SQL_QTY
    def _start_workers(self):
        """각각의 SQL과 인터벌을 설정하여 쓰레드를 시작."""


        # 예시 SQL. 실제 환경에 맞게 수정하세요.
        daily_sql   = self.Make_SQL('DAY')
        weekly_sql  = self.Make_SQL('WEEK')
        monthly_sql = self.Make_SQL('MONTH')
        fees_sql    = """
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

        # 각각 업데이트 주기(초) 설정
        self.worker_daily = QueryWorker(daily_sql,   interval_sec=30)
        self.worker_weekly = QueryWorker(weekly_sql, interval_sec=30)
        self.worker_monthly = QueryWorker(monthly_sql, interval_sec=30)
        self.worker_fees = QueryWorker(fees_sql,      interval_sec=400)

        # 시그널 연결
        self.worker_daily.data_ready.connect(self._update_daily_plot)
        self.worker_weekly.data_ready.connect(self._update_weekly_plot)
        self.worker_monthly.data_ready.connect(self._update_monthly_plot)
        self.worker_fees.data_ready.connect(self._update_fees_plot)

        # 쓰레드 시작
        for w in (self.worker_daily, self.worker_weekly,
                  self.worker_monthly, self.worker_fees):
            w.start()


    def _update_daily_plot(self, rows):
        global TDate, WDate, MDate
        print(rows)
        self._update_plot(self.canvas_daily, rows, xlabel="", ylabel="구좌", Barnum=2, title='일간실적', T_Date = TDate, F_Date=TDate, GType = "DAY")

    def _update_weekly_plot(self, rows):
        self._update_plot(self.canvas_weekly, rows, xlabel="", ylabel="구좌", Barnum=2, title='주간실적', T_Date = TDate, F_Date=WDate, GType = "WEEK")

    def _update_monthly_plot(self, rows):
        self._update_plot(self.canvas_monthly, rows, xlabel="", ylabel="구좌", Barnum=2, title='월간실적', T_Date = TDate, F_Date=MDate, GType = "MONTH")

    def _update_fees_plot(self, rows):
        self._update_plot(self.canvas_fees, rows, xlabel="", ylabel="구좌", Barnum=1, title='정산 예정 구좌 (2회차 입금 구좌)', T_Date = "", F_Date="", GType = "ALLOW")

    def _update_plot(self, canvas: FigureCanvas, rows, xlabel: str, ylabel: str , Barnum: int, title:str, F_Date:str, T_Date:str, GType:str):
        xs = [r[0] for r in rows]
        if GType == "ALLOW":
            ys = [r[1] for r in rows]
            ys_max = ys
        else:
            ys = [r[1] for r in rows]
            ys2 = [r[2] for r in rows]
            ys3 = [r[3] for r in rows]
            ys_max = ys3

        if GType == "DAY":
            title = f"{T_Date[:4]}년 {T_Date[4:6]}월 {T_Date[6:8]}일 실적"
        if GType == "MONTH":
            title = f"{F_Date[:4]}년 {F_Date[4:6]}월 월간실적"
        if GType == "WEEK":
            title = f"주간실적 ({F_Date[:4]}년{F_Date[4:6]}월{F_Date[6:8]}일 ~ {T_Date[:4]}년{T_Date[4:6]}월{T_Date[6:8]}일)"
        if GType == "ALLPW":
            title = title

        print("Gtype:", GType, F_Date, T_Date)

        ax = canvas.ax
        ax.clear()

        # 1) 타이틀 설정
        ax.set_title(title, fontsize=24)

        # 막대그래프(bar chart)로 변경
        if Barnum == 1:
            ax.bar(xs, ys, label = "인정구좌")
            ax.axhspan(0,30, color='red', alpha=0.1)
        else:
            ax.bar(xs, ys, label = '접수')  # , color='skyblue', edgecolor='gray')
            ax.bar(xs, ys2, bottom=ys, label='계약')  # , color='skyblue', edgecolor='gray')
        # x축 레이블 글자가 겹치지 않도록 회전
        ax.set_xticks(xs)
        ax.set_xticklabels(xs, rotation=30, ha='center', fontsize=16)

        ax.set_xlabel(xlabel)
        ax.set_ylabel(ylabel)

        ax.grid(True, axis='y', linestyle='--', alpha=0.7)
        if ys_max == 0:
            Max_Y = 10
        else:
            Max_Y = float(max(ys_max))+float(max(ys_max))*0.1

        ax.set_ylim(0,Max_Y)


        # 5) 범례 표시
        ax.legend(loc='upper right', fontsize=9, frameon=True)

        # 레이아웃 조정
        box1 = {'boxstyle': 'round,pad=0.2', 'ec': (1.0, 0.5, 0.5), 'fc': (1.0, 0.8, 0.8), 'linewidth': 2}
        canvas.figure.tight_layout()
        for x, y in zip(xs, ys_max):
            if y == max(ys_max):
                ax.text(x, y, f"{y}", fontsize=16,
                        color='#dc00ff',
                        horizontalalignment='center',  # horizontalalignment (left, center, right)
                        verticalalignment='bottom', bbox=box1)  # verticalalignment (top, center, bottom))
            else:
                ax.text(x, y, f"{y}",fontsize=16,
                         color='blue',
                         horizontalalignment='center',  # horizontalalignment (left, center, right)
                         verticalalignment='bottom')  # verticalalignment (top, center, bottom))
        canvas.draw()


    def closeEvent(self, event):
        """윈도우 닫힐 때 쓰레드 안전하게 중지."""
        for w in (self.worker_daily, self.worker_weekly,
                  self.worker_monthly, self.worker_fees):
            w.stop()
            w.wait()
        super().closeEvent(event)


# 4. 실행부
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = DashboardWindow()
    # win.show()
    win.showFullScreen()
    sys.exit(app.exec())