function stockAnalyze() {
  
  var today = new Date()
  if (isBusinessDay(today)){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('0000');
    var code = sheet.getName();
    var historicalOhlc = getHistoricalOhlcFromKabutan(code,today,today);
    var lastRow = sheet.getLastRow();
    var dateCol = findCol(sheet,'Date',1);
    var dateCel = sheet.getRange(lastRow+1,dateCol,historicalOhlc.length,historicalOhlc[0].length);
    dateCel.setValues(historicalOhlc.reverse());
    Logger.log(getHistoricalOhlcFromKabutan(code));    
  }
}
