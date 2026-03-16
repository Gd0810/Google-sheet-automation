from flask import Flask, render_template, request
import gspread
import re
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

print("Starting Flask App...")

# Google Sheets connection
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

print("Loading credentials...")

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

print("Connecting to Google Sheet...")

sheet = client.open("Girdharan_daily_work_report").sheet1

print("Connected to sheet successfully")


def parse_report(report_text):

    print("\n--- REPORT RECEIVED ---")
    print(report_text)

    date_match = re.search(r'Date:\s*(.*)', report_text)
    date = date_match.group(1) if date_match else ""

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

            # Save previous task if exists
            if start_time and task_text:
                rows.append([date, task_text.strip(), start_time, end_time])
                print("Saved Task:", task_text)

            start_time = time_match.group(1)
            end_time = time_match.group(2)
            task_text = ""

            print("Found Time:", start_time, "-", end_time)

        else:
            task_text += " " + line

    # Save last task
    if start_time and task_text:
        rows.append([date, task_text.strip(), start_time, end_time])
        print("Saved Task:", task_text)

    print("Rows to Insert:", rows)

    return rows


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        report = request.form["report"]

        rows = parse_report(report)

        for row in rows:

            print("Inserting Row:", row)

            sheet.append_row(row)

        print("Data inserted successfully")

        return "Report Saved Successfully!"

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)