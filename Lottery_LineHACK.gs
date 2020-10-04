var sheetApp = SpreadsheetApp.getActiveSpreadsheet();  
var sheet1 = sheetApp.getSheetByName('表單回應 1');  

var firstPrizeRowNum = 3;
var lastPrizeRowNum = 3;

var randColumnRange = "S2:S434";

function setRand() {
  sheet1.getRange(randColumnRange).setFormula("=RAND()");
  sheet1.getRange(randColumnRange).setValues(sheet1.getRange(randColumnRange).getValues());
}

function getPrize() {  
  sheet1.getRange("W" + firstPrizeRowNum + ":W" + lastPrizeRowNum).clearContent();

  SpreadsheetApp.getUi().alert('抽獎中...誰會獲得哈士奇呢🐶');
  
  var j = 2;
  
  // get prize
  for (var i = firstPrizeRowNum; i <= lastPrizeRowNum; i++){
    sheet1.getRange("W" + i).setFontColor("white").setFormula("=INDEX(B:B,MATCH(LARGE(S:S,ROW(B" + j + ")),S:S,0))");
    j++;
  }

  // display
  j = 2;
  for (var i = firstPrizeRowNum; i <= lastPrizeRowNum; i++) {
    sheet1.getRange("W" + i).setValue(sheet1.getRange("W" + i).getValue()).setFontColor("black");
    j++;
  }
}
