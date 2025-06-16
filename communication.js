function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CW Email Tools')
    .addItem('Send Email to Selected Row', 'sendCWEmail')
    .addToUi();
}

function sendCWEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  const data = sheet.getRange(row, 1, 1, 11).getValues()[0]; // A to K

 const [
    email, team, requestType, rawStartDate, rawEndDate,
    reason, submittedBy, status, tlNotes, tlName, emailStatus
  ] = data;

  // Only proceed if status is Approved or Rejected
  if (status !== "Approved" && status !== "Rejected") {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "⚠️ Cannot send email. Status must be 'Approved' or 'Rejected'.",
    "CW Email Status",
    5
  );
  return;
}


  const startDate = formatDate(rawStartDate);
  const endDate = formatDate(rawEndDate); 

  const emailCol = 11; // Column K

  if (emailStatus === "Email Sent") {
    SpreadsheetApp.getUi().alert("❌ Email already sent for this row.");
    return;
  }

  sheet.getRange(row, emailCol).setValue("📨 Sending...");
  SpreadsheetApp.flush();

  const approverEmail = Session.getActiveUser().getEmail();
  const approverName = formatNameFromEmail(approverEmail);
  const cwFirstName = formatNameFromEmail(email);

  let statusEmoji = "⏳";
  let subjectPrefix = "⏳ Request Received";

  if (status === "Approved") {
    statusEmoji = "✅";
    subjectPrefix = "✅ Request Approved";
  } else if (status === "Rejected") {
    statusEmoji = "❌";
    subjectPrefix = "❌ Request Rejected";
  }

  const htmlBody = `
  <div style="font-family: Arial, sans-serif; font-size: 14px; line-height: 1.6;">
    <p>Hi <strong>${cwFirstName}</strong>,</p>

    <p>Thank you for submitting your general communication request. We’ve reviewed it and would like to update you on the status:</p>

    <p><strong>Request Type:</strong> ${requestType}<br>
    <strong>Team:</strong> ${team}<br>

    <strong>Duration:</strong> ${startDate} to ${endDate}</p>

    <p style="font-size: 16px;">${statusEmoji} <strong>Status: ${status}</strong></p>

    ${reason ? `<p>📝 <strong>Reason Provided:</strong> ${reason}</p>` : ''}
    ${tlNotes ? `<p>💬 <strong>TL Notes:</strong> ${tlNotes}</p>` : ''}

    <p>If you have any questions or need further clarification, feel free to reply directly to this email.</p>

    <p>Kind regards,<br><strong>${approverName}</strong></p>
  </div>
`;


  const plainTextBody = `
      Hi ${cwFirstName},

      Your request for ${requestType} from ${startDate} to ${endDate} has been processed.

      Status: ${status}

      ${reason ? 'Reason: ' + reason : ''}
      ${tlNotes ? '\nTL Notes: ' + tlNotes : ''}

      Regards,
      ${approverName}
`;

  MailApp.sendEmail({
    to: email,
    cc: "peter.osoo@cloudfactory.com", // Add yourself in CC
    subject: `Avid Strongroom | ${subjectPrefix} – ${requestType} for ${submittedBy}`,
    htmlBody: htmlBody,
    body: plainTextBody
  });

  sheet.getRange(row, emailCol).setValue("📧 Email Sent");
}

function formatNameFromEmail(email) {
  if (!email) return "Approver";
  const username = email.split("@")[0];
  const parts = username.split(".");
  const name = parts[0];
  return name.charAt(0).toUpperCase() + name.slice(1);
}

function formatDate(rawDate) {
  const date = new Date(rawDate);
  return date.toLocaleDateString('en-GB', {
    day: '2-digit',
    month: 'long',
    year: 'numeric'
  }); // e.g. "14 June 2025"
}

