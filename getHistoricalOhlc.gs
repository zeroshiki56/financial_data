function getHistoricalOhlcFromKabutan(code,startDate,endDate){
  if(typeof(endDate) == "undefined") {
    endDate = new Date();
  }
  if(typeof(startDate) == "undefined") {
    startDate = new Date();
    startDate.setMonth( startDate.getMonth() -12 );
  }
  var regOhlc = /<th scope="row"><time datetime="[\s\S]*?<\/time><\/th>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>/g;
  var regDate = /<time datetime="[\s\S]*?<\/time>/;
  const maxPageInKabutan = 10;
  var p = 1;
  var ohlc = [];
  var searchDate = endDate;

  while ( searchDate.getTime() >= startDate.getTime() ){
    var getUrl = 'https://kabutan.jp/stock/kabuka?code=' + code + '&ashi=Date&page=' + p;
    var response = UrlFetchApp.fetch(getUrl);
    var html = response.getContentText( 'UTF-8' );
    var stockRow = html.match( regOhlc );
    for(var i=0 ; i < stockRow.length ; i++){
      var dateElement = stockRow[i].match( regDate );
      var dateString = dateElement[0]
        .replace( '<time datetime="',"" )
        .replace( /">.*?<\/time>/g,"" )
        .replace( /-/g,"/" );
      var date = new Date(dateString);
      if( searchDate.getTime() >= date.getTime() ) {
        var tdElements = stockRow[i].match( /<td>([\d|.|,]*?)<\/td>/g )
        for(var j=0 ; j < tdElements.length ; j++){
          tdElements[j] = tdElements[j]
          .replace( '<td>',"" )
          .replace( '<\/td>',"" );            
        }
        tdElements.unshift ( Utilities.formatDate(date, "JST", "yyyy/MM/dd") );
        ohlc.push( tdElements );
        searchDate = date;
      }
      if ( searchDate.getTime() <= startDate.getTime() ) {
        break;
      }    
    }
    if ( p < maxPageInKabutan ) {
      p++;
    } else {
      break;
    }
  }
  return ohlc;
}

function isBusinessDay(date){
  if (date.getDate() == 0 || date.getDate() == 6) {
    return false;
  }
  var calJa = CalendarApp.getCalendarById('ja.japanese#holiDay@group.v.calendar.google.com');
  if(calJa.getEventsForDay(date).length > 0){
    return false;
  }
  return true;
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