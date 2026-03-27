from datetime import datetime
import json
from pathlib import Path

import gspread
import requests
from flask import Blueprint, jsonify, redirect, render_template, request
from oauth2client.service_account import ServiceAccountCredentials


user_bp = Blueprint("user", __name__)

SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Girdharan_daily_work_report"

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", SCOPE)
client = gspread.authorize(creds)
spreadsheet = client.open(SPREADSHEET_NAME)
OPEN_LOG_PATH = Path("user_open_log.json")


def read_open_log():
    if not OPEN_LOG_PATH.exists():
        return {}

    try:
        with OPEN_LOG_PATH.open("r", encoding="utf-8") as file:
            data = json.load(file)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def write_open_log(data):
    with OPEN_LOG_PATH.open("w", encoding="utf-8") as file:
        json.dump(data, file, indent=2)


def get_last_open_for_user(email):
    log = read_open_log()
    return log.get(email.lower())


def record_open_for_user(email):
    log = read_open_log()
    now_utc = datetime.utcnow().strftime("%d %B %Y %I:%M:%S %p UTC")
    log[email.lower()] = {
        "email": email.lower(),
        "opened_at": now_utc,
    }
    write_open_log(log)
    return log[email.lower()]


def get_file_metadata():
    access_token = creds.get_access_token().access_token
    fields = ",".join([
        "id",
        "name",
        "webViewLink",
        "modifiedTime",
        "lastModifyingUser(displayName,emailAddress)",
        "owners(displayName,emailAddress)",
    ])
    url = f"https://www.googleapis.com/drive/v3/files/{spreadsheet.id}?fields={fields}"
    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"},
        timeout=20,
    )
    response.raise_for_status()
    return response.json()


def normalize_permission(permission):
    email = permission.get("emailAddress") or "Unknown"
    return {
        "id": permission.get("id"),
        "email": email,
        "display_name": permission.get("displayName") or email,
        "role": permission.get("role") or "unknown",
        "type": permission.get("type") or "unknown",
        "is_owner": permission.get("role") == "owner",
        "permission_details": permission.get("permissionDetails", []),
    }


@user_bp.route("/users", methods=["GET"])
def users_home():
    return render_template("users.html")


@user_bp.route("/users/api/list", methods=["GET"])
def users_list():
    permissions = [normalize_permission(item) for item in spreadsheet.list_permissions()]
    permissions.sort(key=lambda item: (not item["is_owner"], item["email"].lower()))

    return jsonify({
        "users": permissions,
        "owner_emails": [item["email"] for item in permissions if item["is_owner"]],
    })


@user_bp.route("/users/api/detail", methods=["GET"])
def user_detail():
    email = request.args.get("email", "").strip().lower()
    if not email:
        return jsonify({"error": "Email is required."}), 400

    permissions = [normalize_permission(item) for item in spreadsheet.list_permissions()]
    selected = next((item for item in permissions if item["email"].lower() == email), None)
    if selected is None:
        return jsonify({"error": "User email not found in sharing list."}), 404

    metadata = get_file_metadata()
    modified_time = metadata.get("modifiedTime")
    formatted_modified = modified_time
    if modified_time:
        formatted_modified = datetime.strptime(modified_time, "%Y-%m-%dT%H:%M:%S.%fZ").strftime("%d %B %Y %I:%M %p UTC")

    return jsonify({
        "user": selected,
        "sheet": {
            "id": metadata.get("id"),
            "name": metadata.get("name"),
            "url": metadata.get("webViewLink"),
            "modified_time": formatted_modified,
            "last_modifying_user": metadata.get("lastModifyingUser", {}),
            "owners": metadata.get("owners", []),
            "worksheet_count": len(spreadsheet.worksheets()),
            "worksheet_titles": [worksheet.title for worksheet in spreadsheet.worksheets()],
        },
        "tracking": {
            "last_open_via_app": get_last_open_for_user(selected["email"]),
        },
        "notes": {
            "last_view": "Direct Google Sheets last-view time is not available. Last open time below is tracked only when the sheet is opened from this app.",
        },
    })


@user_bp.route("/users/open", methods=["GET"])
def open_sheet_for_user():
    email = request.args.get("email", "").strip().lower()
    if not email:
        return "Email is required.", 400

    permissions = [normalize_permission(item) for item in spreadsheet.list_permissions()]
    selected = next((item for item in permissions if item["email"].lower() == email), None)
    if selected is None:
        return "User email not found in sharing list.", 404

    record_open_for_user(selected["email"])
    metadata = get_file_metadata()
    return redirect(metadata.get("webViewLink") or f"https://docs.google.com/spreadsheets/d/{spreadsheet.id}/edit")
