from flask import Flask, render_template, request
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

app = Flask(__name__)

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open("Girdharan_daily_work_report")


def format_sheet(sheet):

    # Header style (full width)
    sheet.format("A1:Z1", {
        "backgroundColor": {
            "red": 0.2,
            "green": 0.4,
            "blue": 0.8
        },
        "textFormat": {
            "bold": True,
            "foregroundColor": {
                "red": 1,
                "green": 1,
                "blue": 1
            }
        },
        "horizontalAlignment": "CENTER"
    })

    # Center align all columns (full width)
    sheet.format("A:Z", {
        "horizontalAlignment": "CENTER"
    })

    # Set column width for a wider range
    sheet.columns_auto_resize(0, 26)


def get_month_sheet(date_string):

    date_obj = datetime.strptime(date_string.strip(), "%d %B %Y")
    month_name = date_obj.strftime("%B")

    try:
        sheet = spreadsheet.worksheet(month_name)
        print("Using existing sheet:", month_name)

    except:

        print("Creating new sheet:", month_name)

        sheet = spreadsheet.add_worksheet(title=month_name, rows=1000, cols=4)

        header = ["Date", "Task", "From", "To"] + [""] * 22
        sheet.update(values=[header], range_name="A1:Z1")

        format_sheet(sheet)

    return sheet


def add_row_border(sheet, row_number):

    sheet.format(f"A{row_number}:Z{row_number}", {
        "borders": {
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"}
        }
    })


def insert_separator_row(sheet, row_number):
    # Insert a visible gap row with no borders and a taller height
    sheet.insert_row(["", "", "", ""] + [""] * 22, row_number)

    requests = [
        {
            "updateBorders": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_number - 1,
                    "endRowIndex": row_number,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "top": {"style": "NONE"},
                "bottom": {"style": "NONE"},
                "left": {"style": "NONE"},
                "right": {"style": "NONE"},
                "innerHorizontal": {"style": "NONE"},
                "innerVertical": {"style": "NONE"}
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "ROWS",
                    "startIndex": row_number - 1,
                    "endIndex": row_number
                },
                "properties": {"pixelSize": 24},
                "fields": "pixelSize"
            }
        }
    ]

    # Remove the bottom border of the row above the separator
    if row_number > 1:
        requests.append({
            "updateBorders": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_number - 2,
                    "endRowIndex": row_number - 1,
                    "startColumnIndex": 0,
                    "endColumnIndex": 26
                },
                "bottom": {"style": "NONE"}
            }
        })

    sheet.spreadsheet.batch_update({"requests": requests})


def parse_report(report_text):

    date_match = re.search(r'Date:\s*(.*)', report_text)
    date = date_match.group(1).replace("\r","").strip()

    print("Parsed Date:", date)

    lines = report_text.splitlines()

    rows = []
    start_time = None
    end_time = None
    task_text = ""

    time_pattern = re.compile(r'(\d{1,2}:\d{2}\s*[AP]M)\s*[–-]\s*(\d{1,2}:\d{2}\s*[AP]M)')

    for line in lines:

        line = line.strip()

        if not line:
            continue

        time_match = time_pattern.search(line)

        if time_match:

            if start_time and task_text:
                rows.append([date, task_text.strip(), start_time, end_time])

            start_time = time_match.group(1)
            end_time = time_match.group(2)
            task_text = ""

        else:
            task_text += " " + line

    if start_time and task_text:
        rows.append([date, task_text.strip(), start_time, end_time])

    return rows, date


def normalize_time_string(value):
    cleaned = re.sub(r"\s+", " ", value.strip())
    dt = datetime.strptime(cleaned, "%I:%M %p")
    return dt.strftime("%I:%M %p").lstrip("0")


def time_to_minutes(value):
    cleaned = re.sub(r"\s+", " ", value.strip())
    dt = datetime.strptime(cleaned, "%I:%M %p")
    return dt.hour * 60 + dt.minute


def date_to_key(value):
    cleaned = value.strip()
    dt = datetime.strptime(cleaned, "%d %B %Y")
    return dt.date()


def is_blank_row(row):
    return not row or all(cell.strip() == "" for cell in row)


def get_row_date(row):
    if not row or len(row) == 0:
        return ""
    cell = row[0].strip()
    if cell.lower() == "date":
        return ""
    return cell


def find_insert_index(existing_data, date_value, start_minutes):
    # existing_data is 0-based list of rows; return 1-based row index
    new_date_key = date_to_key(date_value)

    for idx, row in enumerate(existing_data[1:], start=2):
        row_date = get_row_date(row)
        if row_date == "":
            continue

        try:
            row_date_key = date_to_key(row_date)
        except Exception:
            continue

        if new_date_key < row_date_key:
            # Insert before this later date (and before its separator if present)
            if idx > 2 and is_blank_row(existing_data[idx - 2]):
                return idx - 1
            return idx

        if new_date_key == row_date_key:
            try:
                row_start = time_to_minutes(row[2])
            except Exception:
                row_start = None

            if row_start is not None and start_minutes < row_start:
                return idx

    return len(existing_data) + 1


def ensure_separator_between(sheet, existing_data, upper_index, lower_index):
    # upper_index and lower_index are 1-based, and should be adjacent
    if lower_index != upper_index + 1:
        return

    upper_row = existing_data[upper_index - 1]
    lower_row = existing_data[lower_index - 1]

    if is_blank_row(lower_row):
        return

    upper_date = get_row_date(upper_row)
    lower_date = get_row_date(lower_row)

    if upper_date and lower_date and upper_date != lower_date:
        insert_separator_row(sheet, lower_index)
        existing_data.insert(lower_index - 1, [""] * 26)


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        report = request.form["report"]

        rows, date = parse_report(report)

        sheet = get_month_sheet(date)

        existing_data = sheet.get_all_values()

        last_date = None
        last_data_row_index = 0

        # Track the last non-empty row and last date (ignore header row)
        for idx, row in enumerate(existing_data, start=1):
            if any(cell.strip() != "" for cell in row):
                last_data_row_index = idx

        for row in reversed(existing_data[1:]):
            if len(row) > 0 and row[0].strip() != "" and row[0].strip().lower() != "date":
                last_date = row[0].strip()
                break

        print("Last date in sheet:", last_date)
        print("Current report date:", date)

        # Build a set of existing (date, from, to) to avoid duplicates
        existing_slots = set()
        for row in existing_data[1:]:
            if len(row) >= 4 and row[0].strip() != "":
                existing_slots.add((
                    row[0].strip(),
                    normalize_time_string(row[2]),
                    normalize_time_string(row[3])
                ))

        rows_to_insert = []
        for row in rows:
            slot_key = (
                row[0].strip(),
                normalize_time_string(row[2]),
                normalize_time_string(row[3])
            )
            if slot_key in existing_slots:
                print("Duplicate slot skipped:", row)
                continue
            existing_slots.add(slot_key)
            rows_to_insert.append(row)

        if not rows_to_insert:
            return "No new rows to insert (duplicate time slots)."

        rows_to_insert.sort(key=lambda r: (date_to_key(r[0]), time_to_minutes(r[2])))

        for row in rows_to_insert:
            insert_at = find_insert_index(existing_data, row[0].strip(), time_to_minutes(row[2]))
            sheet.insert_row(row + [""] * 22, insert_at)
            add_row_border(sheet, insert_at)
            print("Inserted Row:", row)
            # Keep local data in sync for subsequent inserts
            existing_data.insert(insert_at - 1, row + [""] * 22)
            # Ensure blank separator between different dates
            if insert_at > 2:
                ensure_separator_between(sheet, existing_data, insert_at - 1, insert_at)
            if insert_at < len(existing_data):
                ensure_separator_between(sheet, existing_data, insert_at, insert_at + 1)

        sheet.columns_auto_resize(0, 4)

        return "Report Saved Successfully!"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
