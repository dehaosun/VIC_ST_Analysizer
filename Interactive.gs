function goToCell(rng) {
  var sh = sheet.getSheet(); 
  SpreadsheetApp.setActiveSheet(sh);
  SpreadsheetApp.setActiveRange(rng);
}

function getSheetUrl() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId(); 
  return url;
}

function getSheetUrlByName(shName) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getSheetByName(shName);
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId(); 
  return url;
}



function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getAddress(e){
  var rng = e.range;
  var address = range.getA1Notation(); 
  return address;
}