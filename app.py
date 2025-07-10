import os
import sys
import json
import bcrypt
import webbrowser
import threading
import time

from flask import Flask, render_template, request, redirect, url_for, session, flash
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from werkzeug.utils import secure_filename

from form_builder import (
    create_form_and_link_sheet,
    get_linked_sheet_url,
    load_form_metadata,
    inject_script_to_sheet,
    get_linked_sheet_id_from_form,
    cancel_booking_by_phone,
    trigger_form_refresh,
    get_script_id_from_metadata,
    update_sheet_url_in_metadata,
    update_metadata_script_id
)

# Helper for PyInstaller compatibility
def get_resource_path(relative_path):
    if getattr(sys, '_MEIPASS', False):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)

# IMPORTANT:
# Use get_resource_path ONLY for read-only files (client_secret.json, templates)
# Writeable files must be in the working directory
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

CLIENT_SECRETS_FILE = get_resource_path("client_secret.json")
FORMS_JSON_FILE = os.path.join(BASE_DIR, "forms.json")
ADMIN_AUTH_FILE = os.path.join(BASE_DIR, "admin_auth.json")
GOOGLE_CREDS_FILE = os.path.join(BASE_DIR, "google_creds.json")

SCOPES = [
    "https://www.googleapis.com/auth/forms.body",
    "https://www.googleapis.com/auth/forms.responses.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.projects",
    "https://www.googleapis.com/auth/script.deployments",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/script.external_request",
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile"
]

app = Flask(
    __name__,
    template_folder=get_resource_path('templates'),
    static_folder=get_resource_path('static')
)
app.secret_key = "your_super_secret_key"

def ensure_google_credentials():
    if "credentials" not in session:
        if os.path.exists(GOOGLE_CREDS_FILE):
            with open(GOOGLE_CREDS_FILE, "r") as f:
                session["credentials"] = json.load(f)
            return True
        return False
    return True

def upload_pdf_to_drive(creds, filepath, filename):
    drive_service = build("drive", "v3", credentials=creds)
    file_metadata = {"name": filename, "mimeType": "application/pdf"}
    media = MediaFileUpload(filepath, mimetype="application/pdf")
    file = drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    drive_service.permissions().create(fileId=file["id"], body={"type": "anyone", "role": "reader"}).execute()
    return f"https://drive.google.com/uc?id={file['id']}&export=download"

@app.route("/")
@app.route("/dashboard")
def dashboard():
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    if not ensure_google_credentials():
        return redirect(url_for("login"))
    forms = load_form_metadata()
    return render_template("dashboard.html", forms=forms)

@app.route("/set_password", methods=["GET", "POST"])
def set_password():
    if os.path.exists(ADMIN_AUTH_FILE):
        return redirect(url_for("admin_login"))
    error = None
    if request.method == "POST":
        password = request.form["password"]
        confirm = request.form["confirm_password"]
        if password != confirm:
            error = "Passwords do not match."
        elif len(password) < 4:
            error = "Password too short."
        else:
            hashed = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
            with open(ADMIN_AUTH_FILE, "w") as f:
                json.dump({"password": hashed}, f)
            return redirect(url_for("admin_login"))
    return render_template("set_password.html", error=error)

@app.route("/admin_login", methods=["GET", "POST"])
def admin_login():
    if not os.path.exists(ADMIN_AUTH_FILE):
        return redirect(url_for("set_password"))
    error = None
    if request.method == "POST":
        password = request.form["password"]
        with open(ADMIN_AUTH_FILE, "r") as f:
            data = json.load(f)
        if bcrypt.checkpw(password.encode(), data["password"].encode()):
            session["admin_logged_in"] = True
            return redirect(url_for("dashboard"))
        error = "Incorrect password."
    return render_template("admin_login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    if os.path.exists(GOOGLE_CREDS_FILE):
        os.remove(GOOGLE_CREDS_FILE)
    return redirect(url_for("admin_login"))

@app.route("/login")
def login():
    # Allow HTTP for local development
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    auth_url, _ = flow.authorization_url(prompt="consent")
    return redirect(auth_url)

@app.route("/oauth2callback")
def oauth2callback():
    # Allow HTTP for local development
    os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE,
        scopes=SCOPES,
        redirect_uri=url_for("oauth2callback", _external=True)
    )
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    creds_data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes
    }
    session["credentials"] = creds_data
    with open(GOOGLE_CREDS_FILE, "w") as f:
        json.dump(creds_data, f)
    return redirect(url_for("dashboard"))

@app.route("/create_form", methods=["POST"])
def create_form():
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])
    class_name = request.form["class_name"]

    slot_names = request.form.getlist("slot_name[]")
    slot_limits = request.form.getlist("slot_limit[]")
    slot_dates = request.form.getlist("slot_date[]")

    slots = []
    for n, l, d in zip(slot_names, slot_limits, slot_dates):
        if n.strip():
            slots.append({"name": n.strip(), "limit": int(l.strip()), "date": d.strip()})

    notes_url = ""
    pdf = request.files.get("notes_pdf")
    if pdf and pdf.filename:
        filename = secure_filename(pdf.filename)
        path = os.path.join(UPLOAD_FOLDER, filename)
        pdf.save(path)
        notes_url = upload_pdf_to_drive(creds, path, filename)
        os.remove(path)

    form_info = {
        "title": class_name,
        "class_name": class_name,
        "slots": slots,
        "limit": int(request.form.get("limit", "10")),
        "meet_link": request.form.get("meet_link", "").strip(),
        "notes": notes_url.strip()
    }

    form_url, edit_url, form_id = create_form_and_link_sheet(creds, form_info)
    time.sleep(5)
    update_sheet_url_in_metadata(creds, form_id)

    flash("Form created successfully.", "success")
    return redirect(url_for("dashboard"))

@app.route("/inject_script/<form_id>", methods=["POST"])
def inject_script(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])
    sheet_id = get_linked_sheet_id_from_form(creds, form_id)
    forms = load_form_metadata()
    target = next((f for f in forms if f["form_id"] == form_id), None)
    if not target:
        flash("Form not found.", "danger")
        return redirect(url_for("dashboard"))
    slot_limits = {s["name"]: {"limit": s["limit"], "expiry": s["date"]} for s in target["slots"]}
    script_id = inject_script_to_sheet(
        creds,
        sheet_id,
        target["class_name"],
        target["form_edit_url"],
        slot_limits,
        form_id,
        target.get("meet_link", ""),
        target.get("notes", "")
    )
    update_metadata_script_id(form_id, script_id)
    flash("Script injected successfully.", "success")
    return redirect(url_for("dashboard"))

@app.route("/edit_metadata/<form_id>", methods=["GET", "POST"])
def edit_metadata(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    forms = load_form_metadata()
    target = next((f for f in forms if f["form_id"] == form_id), None)
    if not target:
        flash("Form not found.", "danger")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        target["meet_link"] = request.form.get("meet_link", "").strip()
        target["notes"] = request.form.get("notes", "").strip()
        with open(FORMS_JSON_FILE, "w") as f:
            json.dump(forms, f, indent=2)
        flash("Metadata updated. Re-inject script to apply changes.", "success")
        return redirect(url_for("dashboard"))
    return render_template("edit_metadata.html", form=target)

@app.route("/change_password", methods=["GET", "POST"])
def change_password():
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    error = None
    success = None
    if request.method == "POST":
        current = request.form["current_password"]
        new = request.form["new_password"]
        confirm = request.form["confirm_password"]
        with open(ADMIN_AUTH_FILE, "r") as f:
            data = json.load(f)
        stored_hash = data.get("password", "")
        if not bcrypt.checkpw(current.encode(), stored_hash.encode()):
            error = "Current password is incorrect."
        elif new != confirm:
            error = "New passwords do not match."
        elif len(new) < 4:
            error = "New password must be at least 4 characters."
        else:
            new_hash = bcrypt.hashpw(new.encode(), bcrypt.gensalt()).decode()
            with open(ADMIN_AUTH_FILE, "w") as f:
                json.dump({"password": new_hash}, f)
            success = "Password changed successfully."
    return render_template("change_password.html", error=error, success=success)

# 🟢 Continue in next message (limit reached)
@app.route("/view_submissions/<form_id>")
def view_submissions(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])

    forms = load_form_metadata()
    target = next((f for f in forms if f["form_id"] == form_id), None)
    if not target:
        flash("Form not found.", "danger")
        return redirect(url_for("dashboard"))

    if target.get("sheet_url"):
        sheet_id = target["sheet_url"].split("/d/")[1].split("/")[0]
    else:
        sheet_id = get_linked_sheet_id_from_form(creds, form_id)

    sheets_service = build("sheets", "v4", credentials=creds)
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range="A1:Z1000"
    ).execute()
    values = result.get("values", [])

    if not values or len(values) < 2:
        return render_template("view_submissions.html", form=target, submissions=[], chart_data={})

    headers = values[0]
    submissions = [dict(zip(headers, row)) for row in values[1:]]

    slot_counts = {}
    form_slots = {s["name"]: s["limit"] for s in target["slots"]}
    for slot in form_slots:
        slot_counts[slot] = 0

    for submission in submissions:
        raw_slot = submission.get("Choose a Slot", "").strip()
        slot_clean = raw_slot.split(" (")[0].strip()
        status = submission.get("Status", "").strip().lower()
        if status != "cancelled" and slot_clean in slot_counts:
            slot_counts[slot_clean] += 1

    chart_data = {
        "slots": list(slot_counts.keys()),
        "booked": list(slot_counts.values()),
        "limits": [form_slots[s] for s in slot_counts]
    }

    return render_template(
        "view_submissions.html",
        form=target,
        submissions=submissions,
        chart_data=chart_data
    )


@app.route("/update_sheet_url/<form_id>")
def update_sheet_url(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    if not ensure_google_credentials():
        return redirect(url_for("login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])

    forms = load_form_metadata()
    target = next((f for f in forms if f["form_id"] == form_id), None)
    if not target:
        flash("Form not found.", "danger")
        return redirect(url_for("dashboard"))

    sheet_id = get_linked_sheet_id_from_form(creds, form_id)
    if not sheet_id:
        flash("Could not find linked Google Sheet. Make sure you linked the form to a Sheet in Google Forms > Responses.", "danger")
        return redirect(url_for("dashboard"))

    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    target["sheet_url"] = sheet_url
    with open(FORMS_JSON_FILE, "w") as f:
        json.dump(forms, f, indent=2)

    flash("Sheet URL updated successfully.", "success")
    return redirect(url_for("dashboard"))


@app.route("/refresh_slots/<form_id>", methods=["POST"])
def refresh_slots(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])
    script_id = get_script_id_from_metadata(form_id)
    if not script_id:
        flash("Script ID not found.", "danger")
        return redirect(url_for("dashboard"))
    trigger_form_refresh(creds, script_id)
    flash("Slots refreshed successfully.", "success")
    return redirect(url_for("dashboard"))


@app.route("/cancel_booking", methods=["POST"])
def cancel_booking():
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])
    form_id = request.form["form_id"]
    mobile = request.form["mobile_number"].strip()
    result = cancel_booking_by_phone(creds, form_id, mobile)
    flash(result, "info")
    return redirect(url_for("dashboard"))


@app.route("/delete_form/<form_id>")
def delete_form(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    if not ensure_google_credentials():
        return redirect(url_for("login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])

    drive_service = build("drive", "v3", credentials=creds)
    forms_service = build("forms", "v1", credentials=creds)

    try:
        drive_service.files().delete(fileId=form_id).execute()
        flash("Form file deleted from Drive.", "success")
    except Exception as e:
        flash(f"Error deleting Form file: {e}", "danger")

    try:
        form = forms_service.forms().get(formId=form_id).execute()
        sheet_id = form.get("linkedSheetId")
        if sheet_id:
            drive_service.files().delete(fileId=sheet_id).execute()
            flash("Linked Sheet deleted from Drive.", "success")
        else:
            flash("No linked Sheet found.", "info")
    except Exception as e:
        flash(f"Error deleting linked Sheet: {e}", "danger")

    forms = load_form_metadata()
    forms = [f for f in forms if f["form_id"] != form_id]
    with open(FORMS_JSON_FILE, "w") as f:
        json.dump(forms, f, indent=2)

    flash("Metadata removed.", "success")
    return redirect(url_for("dashboard"))


@app.route("/update_metadata/<form_id>", methods=["POST"])
def update_metadata(form_id):
    if "admin_logged_in" not in session:
        return redirect(url_for("admin_login"))
    creds = Credentials.from_authorized_user_info(info=session["credentials"])

    forms = load_form_metadata()
    target = next((f for f in forms if f["form_id"] == form_id), None)
    if not target:
        flash("Form not found.", "danger")
        return redirect(url_for("dashboard"))

    meet_link = request.form.get("meet_link", "").strip()
    if meet_link:
        target["meet_link"] = meet_link

    pdf_file = request.files.get("notes_pdf")
    if pdf_file and pdf_file.filename:
        filename = secure_filename(pdf_file.filename)
        path = os.path.join(UPLOAD_FOLDER, filename)
        pdf_file.save(path)
        notes_url = upload_pdf_to_drive(creds, path, filename)
        target["notes"] = notes_url
        os.remove(path)

    with open(FORMS_JSON_FILE, "w") as f:
        json.dump(forms, f, indent=2)

    sheet_id = get_linked_sheet_id_from_form(creds, form_id)
    if not sheet_id:
        flash("Linked sheet not found.", "danger")
        return redirect(url_for("dashboard"))

    slot_limits = {s["name"]: {"limit": s["limit"], "expiry": s["date"]} for s in target["slots"]}
    script_id = inject_script_to_sheet(
        creds,
        sheet_id,
        target["class_name"],
        target["form_edit_url"],
        slot_limits,
        form_id,
        target.get("meet_link", ""),
        target.get("notes", "")
    )
    target["script_id"] = script_id
    with open(FORMS_JSON_FILE, "w") as f:
        json.dump(forms, f, indent=2)

    flash("Metadata updated and script re-injected.", "success")
    return redirect(url_for("dashboard"))

# Automatically open browser after server starts
def open_browser():
    webbrowser.open("http://localhost:5000")

if __name__ == "__main__":
    threading.Timer(1.5, open_browser).start()
    app.run(debug=False)
