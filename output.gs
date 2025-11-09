function AddDataLine( data ){

  const sheet = SpreadsheetApp.getActiveSheet()

  //最終行にデータを追加
  sheet.appendRow(data);

}
function ClearCells(){

  const sheet = SpreadsheetApp.getActiveSheet()

  let clearRange = sheet.getDataRange()

  sheet.clearContents()
}