from flask import Flask, render_template, request, redirect, url_for
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
from datetime import datetime
from functools import lru_cache
import time
import json

app = Flask(__name__)

# ---- Google Sheet Config ----

# Lokasi file service_account.json
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), "service_account.json")

# === Tambahan penting: Buat file dari ENV jika belum ada ===
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    service_json = os.getenv("SERVICE_ACCOUNT_JSON")
    if service_json:
        with open(SERVICE_ACCOUNT_FILE, "w") as f:
            f.write(service_json)

SPREADSHEET_ID = "1Bj_ot_RMEs_usrIxZH8JPV99yThUUPR8CRHgJvbZ890"

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Load credentials setelah file dijamin tersedia
creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
client = gspread.authorize(creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

AVAILABLE_SHEETS = {
    "AIA": "AIA",
    "OCR": "OCR",
    "MYR": "MYR",
    "SGD": "SGD",
    "CRM": "CRM",
    "CTX": "CTX",
    "GRP": "GRP",
    "Suggestions": "Suggestions"
}

@lru_cache(maxsize=50)
def cached_get_records(sheet_name):
    sheet = spreadsheet.worksheet(sheet_name)
    return sheet.row_values(1), sheet.get_all_records()

# ---- Duplicate check ----
def is_duplicate_case(case_id):
    case_id = str(case_id).strip().lower()
    spreadsheet = client.open_by_key(SPREADSHEET_ID)
    for sheet in spreadsheet.worksheets():
        records = sheet.get_all_records()
        for row in records:
            existing = str(row.get("Case id", "")).strip().lower()
            if existing == case_id:
                return True
    return False

# === Cache reset tiap 30 detik ===
LAST_CACHE_RESET = time.time()

@app.before_request
def clear_cache_periodically():
    global LAST_CACHE_RESET
    if time.time() - LAST_CACHE_RESET > 30:
        cached_get_records.cache_clear()
        LAST_CACHE_RESET = time.time()

# ---- Add case route ----
@app.route("/add", methods=["GET", "POST"])
def add_case():
    if request.method == "POST":
        case_id = request.form.get("case_id")

        # Check duplicate
        if is_duplicate_case(case_id):
            return render_template(
                "add_case.html",
                sheet_options=AVAILABLE_SHEETS.keys(),
                error=f"Case ID '{case_id}' already exists!"
            )

        datetime_str = request.form.get("datetime")
        combined_datetime = datetime.strptime(datetime_str, '%Y-%m-%dT%H:%M')

        brand_name = request.form.get("brand_name")
        channel = request.form.get("channel")
        description = request.form.get("description")
        assigned_to = request.form.get("assigned_to")
        status = request.form.get("status")
        remark = request.form.get("remark")
        category = request.form.get("category")
        telegram_link = request.form.get("telegram_link")

        sheet = spreadsheet.worksheet(AVAILABLE_SHEETS[category])
        sheet.append_row([
            case_id,
            combined_datetime.strftime("%Y-%m-%d %H:%M"),
            brand_name,
            channel,
            description,
            assigned_to,
            status,
            remark,
            telegram_link
        ])

        return redirect(url_for("index", sheet=category))

    return render_template("add_case.html", sheet_options=AVAILABLE_SHEETS.keys())

# ---- Add suggestion route (POST via modal) ----
@app.route("/add_suggestion", methods=["POST"])
def add_suggestion():
    suggestion_text = request.form.get("suggestion")
    user_name = request.form.get("user_name", "Anonymous")
    department = request.form.get("department","")
    product = request.form.get("product", "")

    if suggestion_text.strip():
        # Pastikan sheet Suggestions ada
        try:
            suggestion_sheet = spreadsheet.worksheet("Suggestions")
        except gspread.WorksheetNotFound:
            suggestion_sheet = spreadsheet.add_worksheet(title="Suggestions", rows="1000", cols="5")
            suggestion_sheet.append_row(["Timestamp", "User","Department", "Product", "Suggestion"])

        # Append new suggestion
        suggestion_sheet.append_row([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            user_name,
            department,
            product,
            suggestion_text
        ])

    # Redirect ke dashboard dengan query param untuk toast
    return redirect(url_for("index", suggestion_sent=1))

# ---- Dashboard route ----
@app.route("/")
def index():
    selected_key = request.args.get("sheet", "AIA")
    search_query = request.args.get("search", "").strip()
    page = int(request.args.get("page", 1))
    per_page = 5

    all_records = []
    headers = None

    # === Jika memilih menu "Main": tampilkan semua data dari semua sheet ===
    if selected_key == "Main":
        for key, sheet_name in AVAILABLE_SHEETS.items():
            sheet = spreadsheet.worksheet(sheet_name)
            sheet_headers = sheet.row_values(1)

            # Jangan pakai header AIA; untuk Main kita buat header universal
            if headers is None:
                headers = sheet_headers + ["_sheet"]

            records = sheet.get_all_records()

            for r in records:
                r["_sheet"] = key

                for h in headers:
                    if h not in r:
                        r[h] = ""

            all_records.extend(records)

    else:
        sheet = spreadsheet.worksheet(AVAILABLE_SHEETS.get(selected_key, "AIA"))
        headers = sheet.row_values(1)

        if selected_key == "Suggestions":
            headers.append("_sheet")

        all_records = sheet.get_all_records()

        for r in all_records:
            r["_sheet"] = selected_key

            for h in headers:
                if h not in r or r[h] is None:
                    r[h] = ""
    def matches(r):
        try:
            safe_values = []
            for v in r.values():
                if v is None:
                    safe_values.append("")
                else:
                    safe_values.append(str(v))

            text = " ".join(safe_values)
            return search_query.lower() in text.lower()

        except Exception as e:
            print("SEARCH ERROR:", e)
            return False

    filtered = [r for r in all_records if matches(r)]

    total = len(filtered)
    start = (page - 1) * per_page
    end = start + per_page
    paginated_records = filtered[start:end]
    total_pages = (total + per_page - 1) // per_page

    return render_template(
        "dashboard.html",
        headers=headers,
        records=paginated_records,
        selected_sheet=selected_key,
        search_query=search_query,
        page=page,
        total_pages=total_pages
    )

if __name__ == "__main__":
    app.run(debug=True)
