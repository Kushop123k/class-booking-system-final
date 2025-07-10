import json
import os
from datetime import datetime
from googleapiclient.discovery import build


MASTER_FORM_ID = "1XqnWTpsgR8gUyz2H7R_tWdlJVxZSj2xMd7cg4eEmwwo"
METADATA_FILE = "forms.json"


def load_form_metadata():
    if not os.path.exists(METADATA_FILE):
        return []
    with open(METADATA_FILE, "r") as f:
        return json.load(f)


def save_form_metadata(new_entry):
    data = load_form_metadata()
    data.append(new_entry)
    with open(METADATA_FILE, "w") as f:
        json.dump(data, f, indent=2)


def get_script_id_from_metadata(form_id):
    data = load_form_metadata()
    for entry in data:
        if entry["form_id"] == form_id:
            return entry.get("script_id")
    return None


def update_metadata_script_id(form_id, script_id):
    data = load_form_metadata()
    for entry in data:
        if entry["form_id"] == form_id:
            entry["script_id"] = script_id
    with open(METADATA_FILE, "w") as f:
        json.dump(data, f, indent=2)


def create_form_and_link_sheet(creds, form_info):
    drive_service = build("drive", "v3", credentials=creds)
    forms_service = build("forms", "v1", credentials=creds)

    copied_form = drive_service.files().copy(
        fileId=MASTER_FORM_ID,
        body={"name": f"{form_info['class_name']} Booking Form"}
    ).execute()
    form_id = copied_form["id"]

    slot_names = []
    today = datetime.now()
    for slot in form_info["slots"]:
        slot_date = datetime.strptime(slot.get("date", ""), "%Y-%m-%dT%H:%M")

        if slot_date and slot_date < today:
            continue
        if slot.get("limit", 0) == 0:
            continue
        slot_names.append(slot["name"])

    if not slot_names:
        raise ValueError("No valid slots available to display.")

    forms_service.forms().batchUpdate(formId=form_id, body={
        "requests": [
            {
                "updateFormInfo": {
                    "info": {"title": form_info["class_name"]},
                    "updateMask": "title"
                }
            },
            {
                "createItem": {
                    "item": {
                        "title": "Choose a Slot",
                        "questionItem": {
                            "question": {
                                "required": True,
                                "choiceQuestion": {
                                    "type": "RADIO",
                                    "options": [{"value": slot} for slot in slot_names],
                                    "shuffle": False
                                }
                            }
                        }
                    },
                    "location": {"index": 0}
                }
            }
        ]
    }).execute()

    form_url = f"https://docs.google.com/forms/d/{form_id}/viewform"
    edit_url = f"https://docs.google.com/forms/d/{form_id}/edit"

    save_form_metadata({
        "class_name": form_info["class_name"],
        "form_id": form_id,
        "form_url": form_url,
        "form_edit_url": edit_url,
        "slots": form_info["slots"],
        "meet_link": form_info.get("meet_link", ""),
        "notes": form_info.get("notes", ""),
        "sheet_url": ""
    })

    return form_url, edit_url, form_id


def inject_script_to_sheet(creds, sheet_id, form_title, form_edit_url, slot_limits_dict, form_id, meet_link, notes_url):
    from googleapiclient.discovery import build
    import json

    drive_service = build("drive", "v3", credentials=creds)
    script_service = build("script", "v1", credentials=creds)

    # ✅ Check if script_id already exists
    existing_script_id = get_script_id_from_metadata(form_id)

    if existing_script_id:
        project_id = existing_script_id
        print(f"Updating existing script project: {project_id}")
    else:
        # Create new project first time
        drive_service.files().get(fileId=sheet_id, fields="name").execute()
        project = script_service.projects().create(body={
            "title": f"{form_title} Script",
            "parentId": sheet_id
        }).execute()
        project_id = project["scriptId"]
        update_metadata_script_id(form_id, project_id)
        print(f"Created new script project: {project_id}")

    # ✅ Convert Python data for JS
    slot_limits_js = json.dumps(slot_limits_dict)
    form_url_js = json.dumps(form_edit_url)
    meet_link_js = json.dumps(meet_link)
    notes_url_js = json.dumps(notes_url)

    # ✅ Your Apps Script code
    script_code = f"""
function setupTrigger() {{
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {{
    ScriptApp.deleteTrigger(triggers[i]);
  }}
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  ScriptApp.newTrigger("onTimeTrigger")
    .timeBased()
    .everyMinutes(5)
    .create();
}}

function onFormSubmit(e) {{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var phoneCol = headers.indexOf("Mobile Number");
  var emailCol = headers.indexOf("Email Address");
  var slotCol = headers.indexOf("Choose a Slot");
  var statusCol = headers.indexOf("Status");

  if (statusCol === -1) {{
    sheet.getRange(1, headers.length + 1).setValue("Status");
    statusCol = headers.length;
  }}

  var newRow = data[data.length - 1];
  var newPhone = newRow[phoneCol];
  var newEmail = newRow[emailCol];

  var isDuplicate = false;

  for (var i = 1; i < data.length - 1; i++) {{
    var row = data[i];
    var phone = row[phoneCol];
    var email = row[emailCol];
    if (phone === newPhone || email === newEmail) {{
      isDuplicate = true;
      break;
    }}
  }}

  if (isDuplicate) {{
    sheet.getRange(data.length, statusCol + 1).setValue("Duplicate");
    return;
  }}

  refreshSlots();
}}

function onTimeTrigger() {{
  var now = new Date();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  var headers = data[0];
  var phoneCol = headers.indexOf("Mobile Number");
  var notifiedCol = headers.indexOf("ExpiryNotified");
  var slotCol = headers.indexOf("Choose a Slot");

  if (notifiedCol === -1) {{
    sheet.getRange(1, headers.length + 1).setValue("ExpiryNotified");
    notifiedCol = headers.length;
  }}

  var limits = {slot_limits_js};

  for (var i = 1; i < data.length; i++) {{
    var row = data[i];
    var phone = row[phoneCol];
    var slotRaw = row[slotCol];
    if (!phone || !slotRaw) continue;

    var slot = slotRaw.toString().replace(/\\s*\\(.*\\)/, "").trim();
    if (!limits[slot]) continue;

    var expiry = new Date(limits[slot].expiry);
    if (now >= expiry && row[notifiedCol] !== "YES") {{
      var url = "https://bhashsms.com/api/sendmsgutil.php"
        + "?user=RCclasses_BW"
        + "&pass=123456"
        + "&sender=BUZWAP"
        + "&phone=" + encodeURIComponent(phone)
        + "&priority=wa"
        + "&stype=normal";

      var meetLink = {meet_link_js};
      var notesUrl = {notes_url_js};

      if (meetLink && notesUrl) {{
        url += "&text=bookmeet"
              + "&htype=document"
              + "&fname=" + encodeURIComponent("notes.pdf")
              + "&url=" + encodeURIComponent(notesUrl)
              + "&Params=" + encodeURIComponent(meetLink);
      }} else if (meetLink && !notesUrl) {{
        url += "&text=" + encodeURIComponent("meet " + meetLink);
      }} else {{
        url += "&text=" + encodeURIComponent("tex1");
      }}

      UrlFetchApp.fetch(url);
      sheet.getRange(i + 1, notifiedCol + 1).setValue("YES");
    }}
  }}
}}

function refreshSlots() {{
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var slotCol = headers.indexOf("Choose a Slot");
  var phoneCol = headers.indexOf("Mobile Number");
  var emailCol = headers.indexOf("Email Address");
  var statusCol = headers.indexOf("Status");
  if (slotCol === -1) return;

  var counts = {{}};
  var uniqueSet = {{}};

  for (var i = 1; i < data.length; i++) {{
    var row = data[i];
    var slotRaw = row[slotCol];
    var phone = row[phoneCol];
    var email = row[emailCol];
    var status = statusCol !== -1 ? row[statusCol] : "";
    if (!slotRaw || status === "Cancelled" || status === "Duplicate") continue;

    var slot = slotRaw.toString().replace(/\\s*\\(.*\\)/, "").trim();
    var key = (phone + "_" + email + "_" + slot).toLowerCase();
    if (uniqueSet[key]) continue;

    uniqueSet[key] = true;
    counts[slot] = (counts[slot] || 0) + 1;
  }}

  var limits = {slot_limits_js};

  var form = FormApp.openByUrl({form_url_js});
  var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  var slotQuestion = null;
  for (var i = 0; i < items.length; i++) {{
    if (items[i].getTitle().toLowerCase().indexOf("slot") !== -1) {{
      slotQuestion = items[i].asMultipleChoiceItem();
      break;
    }}
  }}
  if (!slotQuestion) return;

  var choices = [];
  var now = new Date();
  for (var slot in limits) {{
    var data = limits[slot];
    var limit = parseInt(data.limit);
    var expiry = new Date(data.expiry);
    var current = counts[slot] || 0;
    if (now <= expiry && current < limit) {{
      choices.push(slotQuestion.createChoice(slot + " (" + (limit - current) + " left)", true));
    }}
  }}

  if (choices.length === 0) {{
    choices.push(slotQuestion.createChoice("No slots available", false));
  }}

  slotQuestion.setChoices(choices);
  form.setAcceptingResponses(choices.length > 0);
}}
"""

    # Manifest
    manifest = '{ "timeZone": "Asia/Kolkata", "exceptionLogging": "STACKDRIVER" }'

    # ✅ Update content
    script_service.projects().updateContent(
        scriptId=project_id,
        body={
            "files": [
                {"name": "Code", "type": "SERVER_JS", "source": script_code},
                {"name": "appsscript", "type": "JSON", "source": manifest}
            ]
        }
    ).execute()

    return project_id


def trigger_form_refresh(creds, script_id):
    script_service = build("script", "v1", credentials=creds)
    script_service.scripts().run(
        scriptId=script_id,
        body={"function": "refreshSlots"}
    ).execute()
def get_linked_sheet_id_from_form(creds, form_id):
    forms_service = build("forms", "v1", credentials=creds)
    res = forms_service.forms().get(formId=form_id).execute()
    linkedSheetId = res.get("linkedSheetId")
    if not linkedSheetId:
        raise ValueError("No linked Sheet found. Please create it first in Google Forms.")
    return linkedSheetId
def cancel_booking_by_phone(creds, form_id, phone):
    sheets_service = build("sheets", "v4", credentials=creds)
    sheet_id = get_linked_sheet_id_from_form(creds, form_id)
    if not sheet_id:
        return "Linked sheet not found."

    # Get sheet metadata to get the sheet/tab name
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheet_title = spreadsheet["sheets"][0]["properties"]["title"]

    # Read values
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range=f"{sheet_title}!A1:Z1000"
    ).execute()
    values = result.get("values", [])
    if not values or len(values) < 2:
        return "No data found."

    headers = values[0]
    phone_col = headers.index("Mobile Number") if "Mobile Number" in headers else -1
    status_col = headers.index("Status") if "Status" in headers else -1

    if phone_col == -1:
        return "Mobile Number column not found."

    updates = []
    for i, row in enumerate(values[1:], start=2):
        if len(row) > phone_col and row[phone_col] == phone:
            if status_col == -1:
                # If no Status column, create it
                sheets_service.spreadsheets().values().update(
                    spreadsheetId=sheet_id,
                    range=f"{sheet_title}!{chr(65 + len(headers))}1",
                    valueInputOption="RAW",
                    body={"values": [["Status"]]}
                ).execute()
                status_col = len(headers)
            # Update the status cell
            sheets_service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=f"{sheet_title}!{chr(65 + status_col)}{i}",
                valueInputOption="RAW",
                body={"values": [["Cancelled"]]}
            ).execute()
            return f"Booking for {phone} marked as Cancelled."
    return "Booking not found."
def update_sheet_url_in_metadata(form_id, sheet_url):
    data = load_form_metadata()
    for entry in data:
        if entry["form_id"] == form_id:
            entry["sheet_url"] = sheet_url
            break
    with open(METADATA_FILE, "w") as f:
        json.dump(data, f, indent=2)
from googleapiclient.discovery import build

def get_linked_sheet_url(creds, form_id):
    drive_service = build("drive", "v3", credentials=creds)

    # Search for the spreadsheet whose parents include this Form
    response = drive_service.files().list(
        q=f"'{form_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)"
    ).execute()
    files = response.get("files", [])

    if not files:
        return None

    # Usually there is only 1 linked Sheet
    sheet_id = files[0]["id"]
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    return sheet_url