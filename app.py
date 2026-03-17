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

    # Base alignment for all columns (full width)
    sheet.format("A:Z", {
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "MIDDLE"
    })

    # Task column: left align for readability
    sheet.format("B:B", {
        "horizontalAlignment": "LEFT",
        "verticalAlignment": "MIDDLE"
    })

    # Wrap task text so line breaks show
    sheet.format("B:B", {
        "wrapStrategy": "WRAP"
    })

    # Set column widths for readability
    requests = [
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1
                },
                "properties": {"pixelSize": 140},
                "fields": "pixelSize"
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 1,
                    "endIndex": 2
                },
                "properties": {"pixelSize": 700},
                "fields": "pixelSize"
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 2,
                    "endIndex": 4
                },
                "properties": {"pixelSize": 120},
                "fields": "pixelSize"
            }
        },
        {
            "updateDimensionProperties": {
                "range": {
                    "sheetId": sheet.id,
                    "dimension": "COLUMNS",
                    "startIndex": 4,
                    "endIndex": 26
                },
                "properties": {"pixelSize": 80},
                "fields": "pixelSize"
            }
        }
    ]
    sheet.spreadsheet.batch_update({"requests": requests})


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


TOTAL_COLS = 26


def pad_row(values, total_cols=TOTAL_COLS):
    if len(values) >= total_cols:
        return values[:total_cols]
    return values + [""] * (total_cols - len(values))


def build_insert_row_request(sheet_id, row_number):
    return {
        "insertDimension": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row_number - 1,
                "endIndex": row_number
            },
            "inheritFromBefore": False
        }
    }


def build_update_cells_request(sheet_id, row_number, values):
    padded = pad_row(values)
    cell_values = []
    for value in padded:
        cell_values.append({"userEnteredValue": {"stringValue": str(value)}})
    return {
        "updateCells": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_number - 1,
                "endRowIndex": row_number,
                "startColumnIndex": 0,
                "endColumnIndex": TOTAL_COLS
            },
            "rows": [{"values": cell_values}],
            "fields": "userEnteredValue"
        }
    }


def build_row_border_request(sheet_id, row_number, style="SOLID"):
    return {
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row_number - 1,
                "endRowIndex": row_number,
                "startColumnIndex": 0,
                "endColumnIndex": TOTAL_COLS
            },
            "top": {"style": style},
            "bottom": {"style": style},
            "left": {"style": style},
            "right": {"style": style}
        }
    }


def build_row_height_request(sheet_id, row_number, pixel_size):
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row_number - 1,
                "endIndex": row_number
            },
            "properties": {"pixelSize": pixel_size},
            "fields": "pixelSize"
        }
    }


def insert_separator_row_batch(requests, sheet_id, row_number):
    # Insert a visible gap row with no borders and a taller height
    requests.append(build_insert_row_request(sheet_id, row_number))
    requests.append(build_update_cells_request(sheet_id, row_number, [""] * TOTAL_COLS))
    requests.append(build_row_border_request(sheet_id, row_number, style="NONE"))
    requests.append(build_row_height_request(sheet_id, row_number, 24))


def add_data_row_batch(requests, sheet_id, row_number, row_values):
    requests.append(build_insert_row_request(sheet_id, row_number))
    requests.append(build_update_cells_request(sheet_id, row_number, row_values))
    requests.append(build_row_border_request(sheet_id, row_number, style="SOLID"))
    line_count = max(1, row_values[1].count("\n") + 1)
    requests.append(build_row_height_request(sheet_id, row_number, 28 * line_count))


def parse_report(report_text):

    date_match = re.search(r'Date:\s*(.*)', report_text)
    date = date_match.group(1).replace("\r","").strip()

    print("Parsed Date:", date)

    lines = report_text.splitlines()

    rows = []
    start_time = None
    end_time = None
    task_lines = []

    time_pattern = re.compile(r'(\d{1,2}:\d{2}\s*[AP]M)\s*[–-]\s*(\d{1,2}:\d{2}\s*[AP]M)')

    for line in lines:

        line = line.strip()

        if not line:
            continue

        time_match = time_pattern.search(line)

        if time_match:

            if start_time and task_lines:
                rows.append([date, "\n".join(task_lines).strip(), start_time, end_time])

            start_time = time_match.group(1)
            end_time = time_match.group(2)
            task_lines = []

        else:
            task_lines.append(line)

    if start_time and task_lines:
        rows.append([date, "\n".join(task_lines).strip(), start_time, end_time])

    return rows, date


def normalize_time_string(value):
    cleaned = re.sub(r"\s+", " ", value.strip())
    dt = datetime.strptime(cleaned, "%I:%M %p")
    return dt.strftime("%I:%M %p").lstrip("0")


def time_to_minutes(value):
    cleaned = re.sub(r"\s+", " ", value.strip())
    dt = datetime.strptime(cleaned, "%I:%M %p")
    return dt.hour * 60 + dt.minute


def time_ranges_overlap(start_a, end_a, start_b, end_b):
    # Overlap if ranges intersect (end is exclusive)
    return start_a < end_b and start_b < end_a


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


def ensure_separator_between(requests, sheet_id, existing_data, upper_index, lower_index):
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
        insert_separator_row_batch(requests, sheet_id, lower_index)
        existing_data.insert(lower_index - 1, [""] * TOTAL_COLS)
        # Re-apply borders to the date rows above and below the separator
        requests.append(build_row_border_request(sheet_id, upper_index, style="SOLID"))
        if lower_index + 1 <= len(existing_data):
            requests.append(build_row_border_request(sheet_id, lower_index + 1, style="SOLID"))


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        report = request.form["report"]

        rows, date = parse_report(report)

        sheet = get_month_sheet(date)

        existing_data = sheet.get_all_values()

        print("Current report date:", date)

        # Build a set of existing (date, from, to) and time ranges to avoid duplicates/overlaps
        existing_slots = set()
        existing_ranges = {}
        for row in existing_data[1:]:
            if len(row) >= 4 and row[0].strip() != "":
                row_date = row[0].strip()
                from_norm = normalize_time_string(row[2])
                to_norm = normalize_time_string(row[3])
                existing_slots.add((row_date, from_norm, to_norm))
                try:
                    start_m = time_to_minutes(row[2])
                    end_m = time_to_minutes(row[3])
                except Exception:
                    continue
                existing_ranges.setdefault(row_date, []).append((start_m, end_m))

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
            try:
                new_start = time_to_minutes(row[2])
                new_end = time_to_minutes(row[3])
            except Exception:
                print("Invalid time range skipped:", row)
                continue

            overlaps = False
            for existing_start, existing_end in existing_ranges.get(row[0].strip(), []):
                if time_ranges_overlap(new_start, new_end, existing_start, existing_end):
                    overlaps = True
                    break

            if overlaps:
                print("Overlapping slot skipped:", row)
                continue

            existing_slots.add(slot_key)
            existing_ranges.setdefault(row[0].strip(), []).append((new_start, new_end))
            rows_to_insert.append(row)

        if not rows_to_insert:
            return "No new rows to insert (duplicate time slots)."

        rows_to_insert.sort(key=lambda r: (date_to_key(r[0]), time_to_minutes(r[2])))

        batch_requests = []

        for row in rows_to_insert:
            insert_at = find_insert_index(existing_data, row[0].strip(), time_to_minutes(row[2]))
            add_data_row_batch(batch_requests, sheet.id, insert_at, row)
            print("Inserted Row:", row)
            # Keep local data in sync for subsequent inserts
            existing_data.insert(insert_at - 1, pad_row(row))
            # Ensure blank separator between different dates
            if insert_at > 2:
                ensure_separator_between(batch_requests, sheet.id, existing_data, insert_at - 1, insert_at)
            if insert_at < len(existing_data):
                ensure_separator_between(batch_requests, sheet.id, existing_data, insert_at, insert_at + 1)

        if batch_requests:
            sheet.spreadsheet.batch_update({"requests": batch_requests})

        return "Report Saved Successfully!"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
