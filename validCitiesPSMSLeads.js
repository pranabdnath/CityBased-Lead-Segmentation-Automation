function validCitiesPSMSLeads() {

  const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
  const SUBJECT_LINE = "mention your mail subject";
  const NORTH_SHEET = "North-Sheet";
  const SOUTH_SHEET = "South-Sheet";

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const northSheet = ss.getSheetByName(NORTH_SHEET);
    const southSheet = ss.getSheetByName(SOUTH_SHEET);

    const threads = GmailApp.search(`subject:"${SUBJECT_LINE}"`, 0, 1);
    const message = threads[0].getMessages().pop();
    const csvBlob = message.getAttachments()[0];
    const csvDataRaw = Utilities.parseCsv(csvBlob.getDataAsString("UTF-8"));

    const headers = csvDataRaw[0].slice(0, 19);
    const rows = csvDataRaw.slice(1).map(r => r.slice(0, 19));

    const validConditions = ["AS-Y", "AS-N", "Online Without Date", "Online with Date"];
    const northCities = ["Ahmedabad", "Gandhinagar", "Mumbai", "Pune", "Thane", "Navi Mumbai"];
    const southCities = ["Bangalore", "Hyderabad", "Chennai", "Kolkata"];

    const northFiltered = rows.filter(r =>
      northCities.includes(String(r[9]).trim()) &&
      validConditions.includes(String(r[18]).trim())
    );

    const southFiltered = rows.filter(r =>
      southCities.includes(String(r[9]).trim()) &&
      validConditions.includes(String(r[18]).trim())
    );

    writeBelowHeader(northSheet, addFormattedDates(northFiltered));
    writeBelowHeader(southSheet, addFormattedDates(southFiltered));

    markYEvery30Days(northSheet);
    markYEvery30Days(southSheet);

  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}

function addFormattedDates(data) {
  return data.map(r => {
    const dt = tryParseDate(r[11]);
    const formatted = dt ? Utilities.formatDate(dt, "GMT+5:30", "dd-MMM-yyyy") : "";
    return r.concat([formatted]);
  });
}

function writeBelowHeader(sheet, data) {
  const startRow = 2;
  sheet.getRange(startRow, 1, sheet.getLastRow(), sheet.getLastColumn()).clearContent();
  if (data.length > 0) {
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
  }
}

function markYEvery30Days(sh) {
  const buyerCol = 2, projectCol = 8, dateCol = 20, outCol = 21;
  const startRow = 2;
  const numRows = sh.getLastRow() - 1;

  const buyers = sh.getRange(startRow, buyerCol, numRows, 1).getValues();
  const projects = sh.getRange(startRow, projectCol, numRows, 1).getValues();
  const dates = sh.getRange(startRow, dateCol, numRows, 1).getValues();

  const out = [];
  const lastAction = {};

  for (let i = 0; i < numRows; i++) {
    const b = String(buyers[i][0]).trim();
    const p = String(projects[i][0]).trim();
    const raw = dates[i][0];
    if (!b || !p || !(raw instanceof Date)) {
      out.push([""]);
      continue;
    }

    const key = b + "|" + p;
    const dt = new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());

    if (!lastAction[key]) {
      out.push(["Y"]);
      lastAction[key] = dt;
    } else {
      const diff = Math.floor((dt - lastAction[key]) / 86400000);
      out.push([diff >= 30 ? "Y" : "N"]);
      if (diff >= 30) lastAction[key] = dt;
    }
  }

  sh.getRange(startRow, outCol, numRows, 1).setValues(out);
}

function tryParseDate(v) {
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
