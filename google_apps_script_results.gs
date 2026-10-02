function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results")
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet("Results");

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Timestamp",
      "Activity",
      "Student 1",
      "Student 2",
      "Mode",
      "Score 1",
      "Score 2",
      "Matches",
      "Moves / Checks",
      "Time (seconds)",
      "Sets / Details",
      "Submitted At"
    ]);
  }

  const p = e.parameter;

  sheet.appendRow([
    new Date(),
    p.activity || "",
    p.student1 || "",
    p.student2 || "",
    p.mode || "",
    p.score1 || "",
    p.score2 || "",
    p.matches || "",
    p.moves || "",
    p.time_seconds || "",
    p.sets || "",
    p.submitted_at || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({status:"ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}
