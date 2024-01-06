function stockAnalyze() {
  
  var date = new Date()
  var today = new Date(date.getFullYear(),date.getMonth(),date.getDate())
  if (isBusinessDay(today)){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    for ( var i = 0 ; i < sheets.length ; i++ ){
      var sheet = sheets[i];
      var code = getStockCode(sheet);
      var lastRow = sheet.getLastRow();
      var dateCol = findCol(sheet,'Date',1);
      var lastDateCel = sheet.getRange(lastRow,dateCol);
      var lastDateCelValue = lastDateCel.getValue();
      if ( lastDateCelValue == 'Date' ){
        var historicalOhlc = getHistoricalOhlcFromKabutan(code);
      }else{
        var startDate = new Date(lastDateCelValue.getTime());
        startDate.setDate( startDate.getDate() + 1 );
        var historicalOhlc = getHistoricalOhlcFromKabutan(code,startDate,today);
      }
      
      if ( historicalOhlc.length != 0 ){
        var ohlcCel = sheet.getRange(lastRow+1,dateCol,historicalOhlc.length,historicalOhlc[0].length);
        ohlcCel.setValues(historicalOhlc);
      }
    }
  }
}

function getStockCode(sheet) {
  return sheet.getName();
}

function findCol(sheet,val,row){
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat[0].length;i++){
    if(dat[row-1][i-1] == val){
      return i;
    }
  }
  return 0;
}

function findRow(sheet,val,col){ 
  var dat = sheet.getDataRange().getValues(); //受け取ったシートのデータを二次元配列に取得
  for(var i=1;i<dat.length;i++){
    if(dat[i-1][col-1] == val){
      return i;
    }
  }
  return 0;
}

function isBusinessDay(date){
  if (date.getDay() == 0 || date.getDay() == 6) {
    return false;
  }
  var calJa = CalendarApp.getCalendarById('ja.japanese#holiDay@group.v.calendar.google.com');
  if(calJa.getEventsForDay(date).length > 0){
    return false;
  }
  return true;
}