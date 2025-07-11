from flask import Flask, render_template, jsonify
import pyodbc
import sys, os
app = Flask(__name__)

# MSSQL 연결 설정
# conn = pyodbc.connect("DRIVER={SQL Server};SERVER=your_server;DATABASE=your_db;UID=your_user;PWD=your_password")
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
conn = ns_conn
def fetch_data(query):
    cursor = conn.cursor()
    cursor.execute(query)
    data = [{"emp_name": row[0], "qty": row[1]} for row in cursor.fetchall()]
    return data

# 라우트 설정
@app.route('/month_chart')
def month_chart():
    return render_template('month_chart.html')

@app.route('/week_chart')
def week_chart():
    return render_template('week_chart.html')

@app.route('/day_chart')
def day_chart():
    return render_template('day_chart.html')

@app.route('/pie_chart')
def pie_chart():
    return render_template('pie_chart.html')

if __name__ == '__main__':
    app.run(debug=True)