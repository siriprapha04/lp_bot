from flask import Flask, request, send_file
from datetime import datetime
import openpyxl
import psycopg2
import os
import io

app = Flask(__name__)

# ดึง URL ฐานข้อมูลจาก Environment Variable
DB_URL = os.environ.get("DATABASE_URL")


def get_db_connection():
    return psycopg2.connect(DB_URL)


# สร้างตารางถ้ายังไม่มี
def init_db():
    if DB_URL:
        conn = get_db_connection()
        c = conn.cursor()

        c.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id SERIAL PRIMARY KEY,
                timestamp TEXT,
                branch TEXT,
                status TEXT,
                detail TEXT
            )
        """)

        conn.commit()
        c.close()
        conn.close()


init_db()


@app.route("/")
def home():
    return "LINE Bot Running"


@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json

    for event in data.get("events", []):

        if event.get("type") == "message":

            if event["message"]["type"] == "text":

                msg = event["message"]["text"]

                if msg.startswith("#ตรวจ"):

                    branch = ""
                    status = ""
                    detail = ""

                    for line in msg.split("\n"):

                        if line.startswith("สาขา:"):
                            branch = line.replace("สาขา:", "").strip()

                        elif line.startswith("สถานะ:"):
                            status = line.replace("สถานะ:", "").strip()

                        elif line.startswith("รายละเอียด:"):
                            detail = line.replace("รายละเอียด:", "").strip()

                    # บันทึกลงฐานข้อมูล
                    if DB_URL:
                        conn = get_db_connection()
                        c = conn.cursor()

                        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                        c.execute(
                            """
                            INSERT INTO reports
                            (timestamp, branch, status, detail)
                            VALUES (%s, %s, %s, %s)
                            """,
                            (timestamp, branch, status, detail)
                        )

                        conn.commit()
                        c.close()
                        conn.close()

                        print(f"บันทึกลง Cloud Database สำเร็จ: {branch}")

    return "OK"


@app.route("/download")
def download_file():

    if not DB_URL:
        return "ขาดการเชื่อมต่อ Database"

    conn = get_db_connection()
    c = conn.cursor()

    c.execute("""
        SELECT timestamp, branch, status, detail
        FROM reports
        ORDER BY id ASC
    """)

    rows = c.fetchall()

    c.close()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    ws.append(["วันที่-เวลา", "สาขา", "สถานะ", "รายละเอียด"])

    for row in rows:
        ws.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    return send_file(
        out,
        as_attachment=True,
        download_name="LP_Report_Cloud.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
