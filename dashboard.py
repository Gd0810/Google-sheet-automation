from datetime import datetime

import gspread
from flask import Blueprint, jsonify, render_template, request
from oauth2client.service_account import ServiceAccountCredentials


dashboard_bp = Blueprint("dashboard", __name__)

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Girdharan_daily_work_report"

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
client = gspread.authorize(creds)
spreadsheet = client.open(SPREADSHEET_NAME)


def is_blank_row(row):
    return not row or all(str(cell).strip() == "" for cell in row)


def parse_date_value(value):
    return datetime.strptime(value.strip(), "%d %B %Y")


def normalize_date_string(value):
    dt = parse_date_value(value)
    return dt.strftime("%d %B %Y").lstrip("0")


def normalize_time_string(value):
    cleaned = " ".join(value.strip().upper().split())
    dt = datetime.strptime(cleaned, "%I:%M %p")
    return dt.strftime("%I:%M %p").lstrip("0")


def get_month_index(name):
    return datetime.strptime(name, "%B").month


def get_sheet(month_name):
    return spreadsheet.worksheet(month_name)


def month_matches_date(month_name, date_value):
    return get_month_index(month_name) == parse_date_value(date_value).month


def get_report_rows(sheet):
    rows = sheet.get_all_values()
    entries = []

    for row_number, row in enumerate(rows[1:], start=2):
        normalized = (row + ["", "", "", ""])[:4]
        if is_blank_row(normalized):
            continue

        date_value, task_value, from_value, to_value = [str(cell).strip() for cell in normalized]
        if date_value.lower() == "date" or not date_value:
            continue

        entries.append({
            "row_number": row_number,
            "date": date_value,
            "task": task_value,
            "from_time": from_value,
            "to_time": to_value,
        })

    return entries


def cleanup_separators(sheet):
    rows = sheet.get_all_values()
    rows_to_delete = []

    for row_number in range(2, len(rows) + 1):
        row = rows[row_number - 1]
        if not is_blank_row(row):
            continue

        prev_row = rows[row_number - 2] if row_number - 2 >= 1 else None
        next_row = rows[row_number] if row_number < len(rows) else None

        valid_separator = (
            prev_row is not None
            and next_row is not None
            and not is_blank_row(prev_row)
            and not is_blank_row(next_row)
            and str(prev_row[0]).strip() != str(next_row[0]).strip()
        )

        if not valid_separator:
            rows_to_delete.append(row_number)

    for row_number in reversed(rows_to_delete):
        sheet.delete_rows(row_number)


@dashboard_bp.route("/dashboard", methods=["GET"])
def dashboard_home():
    return render_template("dashboard.html")


@dashboard_bp.route("/dashboard/api/months", methods=["GET"])
def get_months():
    months = []

    for worksheet in spreadsheet.worksheets():
        try:
            month_order = get_month_index(worksheet.title)
        except ValueError:
            continue

        entries = get_report_rows(worksheet)
        if not entries:
            continue

        unique_dates = {entry["date"] for entry in entries}
        months.append({
            "name": worksheet.title,
            "month_order": month_order,
            "report_count": len(entries),
            "date_count": len(unique_dates),
        })

    months.sort(key=lambda item: item["month_order"])
    return jsonify({"months": months})


@dashboard_bp.route("/dashboard/api/dates", methods=["GET"])
def get_dates():
    month_name = request.args.get("month", "").strip()
    if not month_name:
        return jsonify({"error": "Month is required."}), 400

    try:
        sheet = get_sheet(month_name)
    except gspread.WorksheetNotFound:
        return jsonify({"error": "Month sheet not found."}), 404

    entries = get_report_rows(sheet)
    grouped = {}
    for entry in entries:
        grouped.setdefault(entry["date"], 0)
        grouped[entry["date"]] += 1

    dates = [
        {"date": date_value, "report_count": count}
        for date_value, count in grouped.items()
    ]
    dates.sort(key=lambda item: parse_date_value(item["date"]))

    return jsonify({"month": month_name, "dates": dates})


@dashboard_bp.route("/dashboard/api/reports", methods=["GET"])
def get_reports():
    month_name = request.args.get("month", "").strip()
    date_value = request.args.get("date", "").strip()

    if not month_name or not date_value:
        return jsonify({"error": "Month and date are required."}), 400

    try:
        sheet = get_sheet(month_name)
    except gspread.WorksheetNotFound:
        return jsonify({"error": "Month sheet not found."}), 404

    normalized_date = normalize_date_string(date_value)
    reports = [
        entry
        for entry in get_report_rows(sheet)
        if normalize_date_string(entry["date"]) == normalized_date
    ]

    reports.sort(key=lambda item: datetime.strptime(item["from_time"], "%I:%M %p"))
    return jsonify({"month": month_name, "date": normalized_date, "reports": reports})


@dashboard_bp.route("/dashboard/api/report/update", methods=["POST"])
def update_report():
    payload = request.get_json(silent=True) or {}

    month_name = str(payload.get("month", "")).strip()
    date_value = str(payload.get("date", "")).strip()
    task_value = str(payload.get("task", "")).strip()
    from_time = str(payload.get("from_time", "")).strip()
    to_time = str(payload.get("to_time", "")).strip()
    row_number = int(payload.get("row_number", 0) or 0)

    if not month_name or not date_value or not task_value or not from_time or not to_time or row_number < 2:
        return jsonify({"error": "Month, date, task, from time, to time, and row number are required."}), 400

    try:
        normalized_date = normalize_date_string(date_value)
        normalized_from = normalize_time_string(from_time)
        normalized_to = normalize_time_string(to_time)
        sheet = get_sheet(month_name)
        if not month_matches_date(month_name, normalized_date):
            return jsonify({"error": f"Date must stay inside the {month_name} sheet."}), 400
    except ValueError:
        return jsonify({"error": "Invalid date or time format."}), 400
    except gspread.WorksheetNotFound:
        return jsonify({"error": "Month sheet not found."}), 404

    sheet.update(range_name=f"A{row_number}:D{row_number}", values=[[
        normalized_date,
        task_value,
        normalized_from,
        normalized_to,
    ]])

    line_count = max(1, task_value.count("\n") + 1)
    sheet.spreadsheet.batch_update({
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet.id,
                        "dimension": "ROWS",
                        "startIndex": row_number - 1,
                        "endIndex": row_number,
                    },
                    "properties": {"pixelSize": 28 * line_count},
                    "fields": "pixelSize",
                }
            }
        ]
    })

    return jsonify({
        "success": True,
        "message": "Report updated successfully.",
        "report": {
            "row_number": row_number,
            "date": normalized_date,
            "task": task_value,
            "from_time": normalized_from,
            "to_time": normalized_to,
        },
    })


@dashboard_bp.route("/dashboard/api/report/delete", methods=["POST"])
def delete_report():
    payload = request.get_json(silent=True) or {}

    month_name = str(payload.get("month", "")).strip()
    row_number = int(payload.get("row_number", 0) or 0)

    if not month_name or row_number < 2:
        return jsonify({"error": "Month and row number are required."}), 400

    try:
        sheet = get_sheet(month_name)
    except gspread.WorksheetNotFound:
        return jsonify({"error": "Month sheet not found."}), 404

    sheet.delete_rows(row_number)
    cleanup_separators(sheet)

    return jsonify({
        "success": True,
        "message": "Report deleted successfully.",
    })
