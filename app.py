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

    # Header style
    sheet.format("A1:D1", {
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

    # Center align all columns
    sheet.format("A:D", {
        "horizontalAlignment": "CENTER"
    })

    # Set column width
    sheet.columns_auto_resize(0, 4)


def get_month_sheet(date_string):

    date_obj = datetime.strptime(date_string.strip(), "%d %B %Y")
    month_name = date_obj.strftime("%B")

    try:
        sheet = spreadsheet.worksheet(month_name)
        print("Using existing sheet:", month_name)

    except:

        print("Creating new sheet:", month_name)

        sheet = spreadsheet.add_worksheet(title=month_name, rows=1000, cols=4)

        sheet.update(values=[["Date", "Task", "From", "To"]], range_name="A1:D1")

        format_sheet(sheet)

    return sheet


def add_row_border(sheet, row_number):

    sheet.format(f"A{row_number}:D{row_number}", {
        "borders": {
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"}
        }
    })


def insert_separator_row(sheet, row_number):
    # Insert a visible gap row with no borders and a taller height
    sheet.insert_row(["", "", "", ""], row_number)

    requests = [
        {
            "updateBorders": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_number - 1,
                    "endRowIndex": row_number,
                    "startColumnIndex": 0,
                    "endColumnIndex": 4
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
                    "endColumnIndex": 4
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

        # If new date detected, insert blank row
        next_row_index = last_data_row_index + 1

        if last_date and last_date != date and last_data_row_index > 1:
            print("New day detected → inserting blank row")
            insert_separator_row(sheet, next_row_index)
            next_row_index += 1

        for row in rows:
            sheet.insert_row(row, next_row_index)
            add_row_border(sheet, next_row_index)
            print("Inserted Row:", row)
            next_row_index += 1

        sheet.columns_auto_resize(0, 4)

        return "Report Saved Successfully!"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
