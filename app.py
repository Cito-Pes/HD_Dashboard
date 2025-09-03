#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import time
import pyodbc
import pymssql
from io import BytesIO
from flask import Flask, Response, render_template
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
import datetime

global TDate, WDate, MDate, now, ToDay
TDate=""
WDate=""
MDate=""
now = ""
ToDay = 0

# ─── 리소스 경로 헬퍼 ─────────────────────────────────────────────
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ─── config.txt 읽기 ──────────────────────────────────────────────
ns_conn = App_Title = App_Ver = None
with open(resource_path("config.txt"), "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if line.startswith("ns_conn ="):
            ns_conn = line.split("=", 1)[1]
        elif line.startswith("app_title ="):
            App_Title = line.split("=", 1)[1]
        elif line.startswith("app_ver ="):
            App_Ver = line.split("=", 1)[1]
        elif line.startswith("pysql_conn ="):
            pysql_conn = line.split("=", 1)[1]

# ─── 사용자 정의 색상 맵 ─────────────────────────────────────────
# 각 period에 대해 접수(rec)·계약(con)·(수수료는 rec만) 색상을 지정
COLOR_MAP = {
    "day":   {"rec": "#1f77b4", "con": "#ff7f0e"},
    "week":  {"rec": "#1f77b4", "con": "#ff7f0e"},
    "month": {"rec": "#1f77b4", "con": "#ff7f0e"},
    "fees":  {"rec": "#008000"}
}

# ─── 한글 폰트 설정 ──────────────────────────────────────────────
font_path = resource_path("static/D2Coding.ttc")
font_prop = fm.FontProperties(fname=font_path)
plt.rc("font", family=font_prop.get_name())

# ─── DB 연결 & 날짜 세트 가져오기 ─────────────────────────────────
def get_db_connection():
    return pyodbc.connect(ns_conn, timeout=5)
    # pysql_conn = "host=211.239.164.106,50106;database=HDTourZone;user=sa;password=sjpw_hdtourzone106"
    # print(pysql_conn)
    # return pymssql.connect(host='211.239.164.106:50106', database='HDTourZone', user='sa', password='sjpw_hdtourzone106', tds_version='7.0')


def get_DATE_SET():
    """오늘(TDate), 이번주 첫 월요일(WDate), 이번달 1일(MDate) 반환 (YYYYMMDD)"""
    SQL_DATE = """
        SELECT
          CONVERT(VARCHAR, GETDATE(), 112),
          CONVERT(VARCHAR, DATEADD(wk,DATEDIFF(wk,0,GETDATE()),0), 112),
          CONVERT(VARCHAR, DATEADD(mm,DATEDIFF(mm,0,GETDATE()),0), 112);
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(SQL_DATE)
    TDate, WDate, MDate = cursor.fetchone()
    cursor.close()
    conn.close()
    return TDate, WDate, MDate

# ─── SQL 생성 함수 ───────────────────────────────────────────────
def make_sql(period: str):
    global TDate, WDate, MDate, now, ToDay
    # TDate, WDate, MDate = get_DATE_SET()

    now = datetime.datetime.now()
    today = datetime.datetime.today()
    weekday = today.weekday()

    TDate = now.strftime("%Y%m%d")
    WDate = today - datetime.timedelta(days=weekday)
    WDate = WDate.strftime("%Y%m%d")
    MDate = datetime.datetime(now.year, now.month, 1).strftime("%Y%m%d")
    ToDay = datetime.datetime.today()

    if period == "day":
        FromDate, ToDate = TDate, TDate
    elif period == "week":
        FromDate, ToDate = WDate, TDate
    elif period == "month":
        FromDate, ToDate = MDate, TDate
    else:
        # 수수료 정산 실적 SQL
        if ToDay.day <= 7:
            return """
                select st.SaName, (a.cnt-isnull(ra.rcnt,0))*0.25 as ETC
                from
                staff st 
                left join (select sabun, count(id) as cnt FROM  Allowance_GCnt_DT where apmonth=CONVERT(CHAR(6), DATEADD(month, -1, GETDATE()), 112) and aptype='신규구좌' group by sabun) a on st.SaBun = a.SaBun
                left join (select sabun, count(id) as rcnt from Allowance_GCntr_DT where apmonth = CONVERT(CHAR(6), DATEADD(month, -1, GETDATE()), 112) group by sabun) ra on st.SaBun = ra.SaBun
                where 
                st.PlaceofDuty='홈쇼핑 TM' 
                and st.SaBun not in ('015101','015102','CJ-SHOP','HN-SHOP','K-SHOP','LO-SHOP')
                and st.OutDate =''
                order by st.SaName
                """
        else:
            return """
                SELECT x.SaName, SUM(x.xx)*0.25 AS ETC
                FROM (
                  SELECT me.TotPay, gu.Cash_Month, me.id, me.name, me.reg_date,
                         st.SaName, st.SaBun,
                         CASE gu.G_etc_str5 WHEN 4 THEN 1 WHEN 2 THEN 2 END xx
                  FROM member me
                  INNER JOIN staff st
                    ON me.Charge_IDP = st.SaBun
                  INNER JOIN goods gu
                    ON me.goods = gu.Goods_ID
                   AND gu.Cash > 0
                   AND gu.Goods_ID NOT LIKE 'WEP%'
                  LEFT JOIN Allowance_DT a ON me.id = a.id
                  WHERE st.PlaceofDuty='홈쇼핑 TM'
                    AND me.MemType IN ('정상','만기','행사')
                    AND me.TotPay / gu.Cash_Month >= 2
                    AND me.TotPay > 0
                    AND a.id IS NULL
                    AND me.reg_date >= '2024-01-01'
    
                  UNION ALL
    
                  SELECT me.TotPay, gu.Cash_Month, me.id, me.name, me.reg_date,
                         st.SaName, st.SaBun,
                         CASE gu.G_etc_str5 WHEN 4 THEN 1 WHEN 2 THEN 2 END xx
                  FROM member me
                  INNER JOIN staff st
                    ON me.Charge_IDP = st.SaBun
                  INNER JOIN goods gu
                    ON me.goods = gu.Goods_ID
                   AND gu.Cash > 0
                   AND gu.Goods_ID LIKE 'WEP%'
                  LEFT JOIN Allowance_DT a ON me.id = a.id
                  WHERE st.PlaceofDuty='홈쇼핑 TM'
                    AND me.MemType IN ('정상','만기','행사')
                    AND me.TotPay > 0
                    AND a.id IS NULL
                    AND me.reg_date >= '2024-01-01'
                ) x
                GROUP BY x.SaName
                --
            """

    # 일간/주간/월간 실적 SQL
    return f"""
        SELECT st.SaName
             , ISNULL(SUM(x.cnt1),0) AS cnt1
             , ISNULL(SUM(x.cnt2),0) AS cnt2
             , ISNULL(SUM(x.cnt1),0)+ISNULL(SUM(x.cnt2),0) AS cnt
        FROM (
          SELECT SaBun, SaName
          FROM Staff
          WHERE PlaceofDuty='홈쇼핑 TM'
            AND BranchOffice='TM1'
            AND OutDate=''
        ) st
        LEFT JOIN (
          (SELECT me.Charge_IDP, me.MemberNo,
                  CASE gu.G_etc_str5 WHEN 4 THEN COUNT(me.ID)*0.25
                                     WHEN 2 THEN COUNT(me.ID)*0.5 END AS cnt1,
                  0 AS cnt2, 0 AS cnt3, 0 AS cnt4
           FROM [Member] me
           INNER JOIN goods gu ON me.Goods = gu.Goods_ID
           WHERE me.EventType IN ('여행','크루즈')
             AND me.MemType='접수'
             AND REPLACE(me.Rec_Date,'-','') BETWEEN '{FromDate}' AND '{ToDate}'
           GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5)

          UNION ALL

          (SELECT me.Charge_IDP, me.MemberNo, 0 AS cnt1,
                  CASE gu.G_etc_str5 WHEN 4 THEN COUNT(me.ID)*0.25
                                     WHEN 2 THEN COUNT(me.ID)*0.5 END AS cnt2,
                  0 AS cnt3, 0 AS cnt4
           FROM [Member] me
           INNER JOIN goods gu ON me.Goods = gu.Goods_ID
           WHERE me.EventType IN ('여행','크루즈')
             AND me.MemType IN ('정상','만기','행사')
             AND REPLACE(me.Rec_Date,'-','') BETWEEN '{FromDate}' AND '{ToDate}'
           GROUP BY me.Charge_IDP, me.MemberNo, gu.G_etc_str5)
        ) x
        ON st.SaBun = x.Charge_IDP
        GROUP BY st.SaName
        ORDER BY st.SaName
    """

# ─── 그래프 생성 함수 ───────────────────────────────────────────────
def create_figure(period: str):
    global TDate, WDate, MDate, now, ToDay
    sql    = make_sql(period)
    # print(sql)
    conn   = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(sql)
    rows   = cursor.fetchall()
    cursor.close()
    conn.close()

    # 이름(x축) 추출
    names = [r[0] for r in rows]
    cmap = COLOR_MAP.get(period, {})
    # print(rows)
    # y값을 float로 변환
    if period == "fees":
        ys      = [float(r[1]) for r in rows]  # ETC
        ys2     = None
        barnum  = 1
        ys_max  = ys
        rec_color = cmap["rec"]

    else:
        ys      = [float(r[1]) for r in rows]  # cnt1 (접수)
        ys2     = [float(r[2]) for r in rows]  # cnt2 (계약)
        barnum  = 2
        rec_color = cmap["rec"]
        con_color = cmap["con"]

        # cnt = cnt1+cnt2는 4번째 컬럼(r[3])에 이미 합산되어 있으므로 float 변환
        ys_max  = [float(r[3]) for r in rows]

    # Figure/Axes 준비
    fig, ax = plt.subplots(figsize=(9.6, 5.4), dpi=100)

    # 제목 설정 (기존 로직 유지)
    # TDate, WDate, MDate = get_DATE_SET()
    if period == "day":
        title = f"{TDate[:4]}년 {TDate[4:6]}월 {TDate[6:8]}일 실적 "
    elif period == "week":
        title = f"주간실적 ({WDate[:4]}년{WDate[4:6]}월{WDate[6:8]}일 ~ "
        title += f"{TDate[:4]}년{TDate[4:6]}월{TDate[6:8]}일)"
    elif period == "month":
        title = f"{MDate[:4]}년 {MDate[4:6]}월 월간실적"
    else:
        if ToDay.day <=7 :
            title = f'{str(int(MDate[4:6])-1)}월 정산 예정 구좌 (2회차 입금 구좌)'
        else:
            title = "정산 예정 구좌 (2회차 입금 구좌)"
    ax.set_title(title, fontsize=24)

    # 막대그래프 그리기
    if barnum == 1:
        ax.bar(names, ys, color=rec_color, label="인정구좌")
        ax.axhspan(0, 30, color="red", alpha=0.1)
    else:
        ax.bar(names, ys, color=rec_color, label="접수")
        ax.bar(names, ys2, bottom=ys, color=con_color, label="계약")

    # X축 설정
    ax.set_xticks(names)
    ax.set_xticklabels(names, rotation=30, ha="center", fontsize=16)
    ax.set_ylabel("구좌")
    ax.grid(True, axis="y", linestyle="--", alpha=0.7)
    ax.legend(loc="upper right", fontsize=9, frameon=True)

    # Y축 한계: float 계산
    max_y = max(ys_max) if ys_max else 0.0
    upper = max_y * 1.1 if max_y > 0 else 10.0
    ax.set_ylim(0, upper)

    # 가장 큰 값 강조
    box1 = dict(boxstyle="round,pad=0.2",
                ec=(1.0,0.5,0.5), fc=(1.0,0.8,0.8), linewidth=2)
    for x, y in zip(names, ys_max):
        color = "#dc00ff" if y == max_y else "blue"
        ax.text(x, y, f"{y:.2f}", fontsize=16, color=color,
                ha="center", va="bottom",
                bbox=box1 if y == max_y else None)

    fig.tight_layout()
    return fig

# ─── Flask 앱 정의 ─────────────────────────────────────────────────
app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True

@app.route("/")
def index():
    title = f"{App_Title}   {App_Ver}"
    return render_template("index.html", title=title)

@app.route("/plot/<period>.png")
def plot_png(period):
    # period는 "day", "week", "month", "fees" 중 하나
    fig = create_figure(period)
    buf = BytesIO()
    FigureCanvas(fig).print_png(buf)
    buf.seek(0)
    return Response(buf.getvalue(), mimetype="image/png")

if __name__ == "__main__":
    app.run('0.0.0.0',port=2500,debug=False)