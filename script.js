if (status !== "Approved" && status !== "Rejected") {
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "⚠️ Cannot send email. Status must be 'Approved' or 'Rejected'.",
    "CW Email Status",
    5
  );
  return;
}
