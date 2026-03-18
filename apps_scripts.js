function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  // On first submission, write headers if the sheet is empty
  if (sheet.getLastRow() === 0) {
    const headers = ["Timestamp", ...Object.keys(data)];
    sheet.appendRow(headers);
  }

  // Write one row: timestamp + cards per category (joined by line breaks)
  const row = [new Date(), ...Object.values(data).map(cards => cards.join("\n"))];
  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateCardSortCounts() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Données brutes");
  const countSheet = ss.getSheetByName("Analyse");

  const lastRow = rawSheet.getLastRow();
  const jsonCells = rawSheet.getRange(2, 2, lastRow - 1, 1).getValues();

  let counts = {};
  let categories = new Set();

  jsonCells.forEach(row => {

    if (!row[0]) return;

    const obj = JSON.parse(row[0]);

    Object.keys(obj).forEach(category => {

      if (category === "Cartes à trier") return; // Skip unassigned cards bucket

      categories.add(category);

      obj[category].forEach(card => {

        if (!counts[card]) counts[card] = {};
        if (!counts[card][category]) counts[card][category] = 0;

        counts[card][category]++;

      });

    });

  });

  const catArray = [...categories].sort((a, b) => a.localeCompare(b, "fr")); // Sort categories alphabetically (French-aware)
  const cardArray = Object.keys(counts).sort((a, b) => a.localeCompare(b, "fr")); // Sort cards alphabetically (French-aware)

  let table = [["Carte", ...catArray]];

  cardArray.forEach(card => {

    let row = [card];

    catArray.forEach(cat => {
      row.push(counts[card][cat] || 0);
    });

    table.push(row);

  });

  countSheet.clear();
  countSheet.getRange(1, 1, table.length, table[0].length).setValues(table);

}
