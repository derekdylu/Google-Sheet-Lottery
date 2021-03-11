var sheetApp = SpreadsheetApp.getActiveSpreadsheet();
var sheet1 = sheetApp.getSheetByName('表單回應 1');

var product = ["IM Night '21 (橢圓)貼紙 20元",
                "Spring '21 貼紙(三角形）20元",
                "I'm in love with myself 貼紙(標籤形）20元",
                "I ❤️ I M 貼紙(菱形) 20元",
                "貼紙套組四張 50元",
                "帆布袋 200元",
                "In Love 成對情侶徽章組 50元",
                "Family徽章 (黃) 25元",
                "一拍即合徽章 (綠) 25元",
                "I'm in love with...徽章 (藍) 25元",
                "明信片 20元"]

var i = 2
var col = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "R"]
// product column index

function getReceipt(){

  var k = sheet1.getRange("T2").getValue() + 1
  // last row index
  
  for(i; i < k; i++){
    var receipt = "臺大資管之夜 商品預購明細\n"

    for(var j = 0; j < 11; j++){
      if(sheet1.getRange(col[j] + i).getValue() == "" || sheet1.getRange(col[j] + i).getValue() == "0"){
        continue
      }else{
        receipt = receipt + sheet1.getRange(col[j] + i).getValue() + "* --- " + product[j] + "\n"
      }
    }

    receipt = receipt + "\n[ 總計 $" + sheet1.getRange(col[11] + i).getValue() + " ]"

    sheet1.getRange("S" + i).setValue(receipt)
    // receipt present cell column index
  }
}