function doGet(e) {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Core");

  const issueKey = e.parameter.key;
  const jiraStatus = (e.parameter.status || "").toUpperCase();

  let finalStatus = null;

  // All statuses that should map to "In Progress"
  const toInProgress = [
    "BLOCKED / REJECTED",
    "TESTING",
    "IN-PROGRESS",
    "STORYBOOK",
    "PLANNING WEB",
    "PLANNING APP"
  ];

  if (toInProgress.includes(jiraStatus)) {
    finalStatus = "In Progress";
  } else if (jiraStatus === "UAT") {
    finalStatus = "Done";
  } else {
    return ContentService.createTextOutput("IGNORED");
  }

  const lastRow = sheet.getLastRow();
  const colO = sheet.getRange(1, 15, lastRow).getValues();
  const colR = sheet.getRange(1, 18, lastRow).getValues();

  for (let i = 0; i < lastRow; i++) {
    if (colO[i][0] === issueKey) {
      sheet.getRange(i + 1, 16).setValue(finalStatus);
      return ContentService.createTextOutput("UPDATED");
    }
    if (colR[i][0] === issueKey) {
      sheet.getRange(i + 1, 19).setValue(finalStatus);
      return ContentService.createTextOutput("UPDATED");
    }
  }

  return ContentService.createTextOutput("KEY_NOT_FOUND");
}