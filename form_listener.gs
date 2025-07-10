const LIMIT = 5;

function onFormSubmit(e) {
  const responses = e.values;
  const phone = responses[2]; // assuming name, phone
  const name = responses[1];

  const classLink = "https://zoom.us/class_link_here";

  // Count responses
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowCount = sheet.getLastRow();

  if (rowCount > LIMIT) {
    // Disable the form
    const form = FormApp.openById("YOUR_FORM_ID_HERE");
    form.setAcceptingResponses(false);
  }

  // Send WhatsApp (use your WhatsApp API here)
  const message = `Hi ${name}, your seat is confirmed. Join the class here: ${classLink}`;
  sendWhatsApp(phone, message);
}

function sendWhatsApp(phone, message) {
  const url = "https://api.ultramsg.com/instanceXXX/messages/chat"; // example
  const payload = {
    token: "YOUR_API_TOKEN",
    to: `+91${phone}`,
    body: message
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(url, options);
}
