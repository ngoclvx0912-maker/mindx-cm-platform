#!/usr/bin/env python3
import json
import os
import sqlite3
import sys
from urllib.parse import parse_qs

# Database path - relative to project directory (where scripts run from)
DB_PATH = "data.db"

BU_NAMES = [
    "HCM1 - PVT", "HCM1 - PXL", "HCM1 - TK",
    "HCM2 - LVV", "HCM2 - NX", "HCM2 - PVD", "HCM2 - SH",
    "HCM3 - 3T2", "HCM3 - HL", "HCM3 - HTLO", "HCM3 - PMH", "HCM3 - PNL",
    "HCM4 - LBB", "HCM4 - TC", "HCM4 - TL", "HCM4 - TT",
    "HN - HĐT", "HN - MK", "HN - NCT", "HN - NHT", "HN - NPS",
    "HN - NVC", "HN - OCP", "HN - TP", "HN - VHHN", "HN - VP",
    "K18 - HCM", "K18 - HN",
    "MB1 - BN", "MB1 - HP", "MB1 - QN", "MB1 - TS",
    "MB2 - PT", "MB2 - TN", "MB2 - VP",
    "MN - BD - DA", "MN - BD - TA", "MN - BD - TDM", "MN - BH - PVT",
    "MN - CT - THD", "MN - VT - LHP",
    "MT - ĐN", "MT - NA", "MT - TH",
    "ONL - ART", "ONL - COD"
]

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bu TEXT NOT NULL,
            month TEXT NOT NULL,
            week INTEGER NOT NULL,
            report_type TEXT NOT NULL,
            data TEXT NOT NULL,
            saved_at TEXT NOT NULL,
            UNIQUE(bu, month, week, report_type)
        )
    """)
    conn.commit()
    return conn

def send_response(data, status=200):
    print(f"Status: {status}")
    print("Content-Type: application/json")
    print("Access-Control-Allow-Origin: *")
    print("Access-Control-Allow-Methods: GET, POST, OPTIONS")
    print("Access-Control-Allow-Headers: Content-Type")
    print()
    print(json.dumps(data, ensure_ascii=False))

def handle_options():
    print("Status: 200")
    print("Access-Control-Allow-Origin: *")
    print("Access-Control-Allow-Methods: GET, POST, OPTIONS")
    print("Access-Control-Allow-Headers: Content-Type")
    print()

def handle_get(params):
    action = params.get("action", [""])[0]
    db = get_db()

    if action == "bus":
        send_response(BU_NAMES)

    elif action == "get":
        bu = params.get("bu", [""])[0]
        month = params.get("month", [""])[0]
        week = params.get("week", [""])[0]

        result = {"T1": [], "T2": [], "T3": [], "meta": {}}

        for rt in ["T1", "T2", "T3"]:
            row = db.execute(
                "SELECT data, saved_at FROM reports WHERE bu=? AND month=? AND week=? AND report_type=?",
                (bu, month, week, rt)
            ).fetchone()
            if row:
                result[rt] = json.loads(row["data"])
                result["meta"][f"{rt}_saved_at"] = row["saved_at"]

        send_response(result)

    elif action == "list":
        month = params.get("month", [""])[0]
        week = params.get("week", [""])[0]

        rows = db.execute(
            "SELECT bu, report_type, saved_at FROM reports WHERE month=? AND week=?",
            (month, week)
        ).fetchall()

        bu_map = {}
        for row in rows:
            bu = row["bu"]
            rt = row["report_type"]
            if bu not in bu_map:
                bu_map[bu] = {"bu": bu, "hasT1": False, "hasT2": False, "hasT3": False}
            bu_map[bu][f"has{rt}"] = True
            bu_map[bu][f"{rt}_saved_at"] = row["saved_at"]

        # Build full list with all BUs
        result = []
        for bu_name in BU_NAMES:
            if bu_name in bu_map:
                result.append(bu_map[bu_name])
            else:
                result.append({"bu": bu_name, "hasT1": False, "hasT2": False, "hasT3": False})

        send_response(result)

    elif action == "all_data":
        month = params.get("month", [""])[0]
        week = params.get("week", [""])[0]

        rows = db.execute(
            "SELECT bu, report_type, data FROM reports WHERE month=? AND week=?",
            (month, week)
        ).fetchall()

        result = {}
        for row in rows:
            bu = row["bu"]
            rt = row["report_type"]
            if bu not in result:
                result[bu] = {"T1": [], "T2": [], "T3": []}
            result[bu][rt] = json.loads(row["data"])

        send_response(result)

    else:
        send_response({"error": "Unknown action"}, 400)

    db.close()

def handle_post():
    content_length = int(os.environ.get("CONTENT_LENGTH", 0))
    body = sys.stdin.read(content_length) if content_length > 0 else sys.stdin.read()

    try:
        payload = json.loads(body)
    except Exception as e:
        send_response({"error": f"Invalid JSON: {str(e)}"}, 400)
        return

    action = payload.get("action", "")

    if action == "save":
        bu = payload.get("bu", "")
        month = payload.get("month", "")
        week = payload.get("week", 0)
        report_type = payload.get("report_type", "")
        data = payload.get("data", [])

        if not bu or not month or not week or not report_type:
            send_response({"error": "Missing required fields"}, 400)
            return

        import datetime
        saved_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        db = get_db()
        db.execute(
            """INSERT OR REPLACE INTO reports (bu, month, week, report_type, data, saved_at)
               VALUES (?, ?, ?, ?, ?, ?)""",
            (bu, month, str(week), report_type, json.dumps(data, ensure_ascii=False), saved_at)
        )
        db.commit()
        db.close()

        send_response({"success": True, "saved_at": saved_at})

    else:
        send_response({"error": "Unknown action"}, 400)

def main():
    method = os.environ.get("REQUEST_METHOD", "GET").upper()
    query_string = os.environ.get("QUERY_STRING", "")
    params = parse_qs(query_string)

    if method == "OPTIONS":
        handle_options()
    elif method == "GET":
        handle_get(params)
    elif method == "POST":
        handle_post()
    else:
        send_response({"error": "Method not allowed"}, 405)

if __name__ == "__main__":
    main()
