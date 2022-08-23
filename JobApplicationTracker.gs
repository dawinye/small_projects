function refreshSheet() {
  const data = []
  const labelCheck = ["Applied", "Rejected", "OAs/Interviewing", "Offers"] //change this if your labels are different
  const spreadsheetId = "" //enter your own id of the spreadsheet, appears after https://docs.google.com/spreadsheets/d/...
  const sheetName = "" //enter the sheetname of the spreadsheet that you want to populate (on the bottom bar of sheets)
  const resource = {
    valueInputOption: "USER_ENTERED",
    data: data,
  };
  const sheetLink = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName)
  const labels = GmailApp.getUserLabels();
  //start row is 2 because first row should be Date, Subject, From, Status
  //freeze the first row by clicking on it and doing View -> Freeze -> 1 Row
  const colNumber = 1
  const startRow = 2

  const numRows = sheetLink.getLastRow() - startRow + 1;
  const numCols = sheetLink.getLastColumn() - colNumber + 1;
  //console.log(labels)
  const range = sheetLink.getRange(startRow, colNumber, numRows, numsCols);
  range.clear();

  // 
  const OA = labels[0]
  const Offers = labels[1]
  const Applied = labels[2]
  const Rejected = labels[3]
  labels[0] = Rejected
  labels[1] = Offers
  labels[2] = OA
  labels[3] = Applied

  var emails = new Set();
  var colorMap = new Map([
    ["Rejected", "red"],
    ["Offers", "green"],
    ["Applied", "blue"],
    ["OAs/Interviewing", "yellow"]
  ]);
  
  for (let i = 0; i < labels.length; i++) {
    const labelNames = labels[i].getName()
    let threads = labels[i].getThreads()
    console.log(labelNames)
    for (let j = 0; j < threads.length; j++) {
      let messages = threads[j].getMessages();
      
      for (let k = 0; k < messages.length; k++) {
        const message = messages[k];

        if (labelCheck.includes(labelNames)) {
          const subject = message.getSubject();
          const date = message.getDate();
          const email = message.getFrom();
          // don't want to see emails under the applying label for company X when we're already in the interview process
          if (!emails.has(email)){
            sheetLink.getRange(sheetLink.getLastRow() + 1, 1, 1, 4).setValues([[date,subject,email,labelNames]]);
            sheetLink.getRange(sheetLink.getLastRow(),4).setBackground(colorMap.get(labelNames))
          }
          emails.add(email)
        }
      }
      console.log("sheet data:", data)
      console.log("---------")
    }
    Sheets.Spreadsheets.Values.batchUpdate(resource, spreadsheetId);
  }
  range.sort({column: 1, ascending: false});
  console.log("done")
}
