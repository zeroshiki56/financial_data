function stockAnalyze() {
  
  var today = new Date()
  if (isBusinessDay(today)){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for ( var i = 0 ; i < sheets.length ; i++ ){
      var sheet = sheets[i];
      var code = sheet.getName();
      var historicalOhlc = getHistoricalOhlcFromKabutan(code,today,today);
      var lastRow = sheet.getLastRow();
      var dateCol = findCol(sheet,'Date',1);
      var dateCel = sheet.getRange(lastRow+1,dateCol,historicalOhlc.length,historicalOhlc[0].length);
      dateCel.setValues(historicalOhlc.reverse());
    }
  }
}
