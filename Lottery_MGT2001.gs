var sheetApp = SpreadsheetApp.getActiveSpreadsheet();  
var sheet1 = sheetApp.getSheetByName('表單回應 1');  

var firstPrizeRowNum = 2;
var lastPrizeRowNum = 2;

var randColumnRange = "D2:D84";

function setRand() {
  sheet1.getRange(randColumnRange).setFormula("=RAND()");
  sheet1.getRange(randColumnRange).setValues(sheet1.getRange(randColumnRange).getValues());
}

function getPrize() {  
  sheet1.getRange("E" + firstPrizeRowNum + ":E" + lastPrizeRowNum).clearContent();

  SpreadsheetApp.getUi().alert('Drawing...');
  
  var j = 2;
  
  // get prize
  for (var i = firstPrizeRowNum; i <= lastPrizeRowNum; i++){
    sheet1.getRange("E" + i).setFontColor("white").setFormula("=INDEX(B:B,MATCH(LARGE(D:D,ROW(B" + j + ")),D:D,0))");
    j++;
  }

  // display
  j = 2;
  for (var i = firstPrizeRowNum; i <= lastPrizeRowNum; i++) {
    sheet1.getRange("E" + i).setValue(sheet1.getRange("E" + i).getValue()).setFontColor("black");
    j++;
  }
}

