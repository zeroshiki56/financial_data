function getHistoricalOhlcFromKabutan(code,startDay,endDay){
  if(typeof(endDay) == "undefined") {
    endDay = new Date();
  }
  if(typeof(startDay) == "undefined") {
    startDay = new Date();
    startDay.setMonth( startDay.getMonth() -12 );
  }
  var regOhlc = /<th scope="row"><time datetime="[\s\S]*?<\/time><\/th>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>\r\n<td>[\s\S]*?<\/td>/g;
  var regDate = /<time datetime="[\s\S]*?<\/time>/;
  const maxPageInKabutan = 10;
  var p = 1;
  var tagElement = [];
  var searchDay = endDay;

  while ( searchDay.getTime() >= startDay.getTime() ){
    var getUrl = 'https://kabutan.jp/stock/kabuka?code=' + code + '&ashi=day&page=' + p;
    var response = UrlFetchApp.fetch(getUrl);
    var html = response.getContentText( 'UTF-8' );
    var stockRow = html.match( regOhlc );
    for(var i=0 ; i < stockRow.length ; i++){
      var dateElement = stockRow[i].match( regDate );
      var dateString = dateElement[0]
        .replace( '<time datetime="',"" )
        .replace( /">.*?<\/time>/g,"" )
        .replace( /-/g,"/" );
      var rowDate = new Date(dateString);
      if( searchDay.getTime() >= rowDate.getTime() ) {
        var tdElements = stockRow[i].match( /<td>([\d|.|,]*?)<\/td>/g )
        for(var j=0 ; j < tdElements.length ; j++){
          tdElements[j] = tdElements[j]
          .replace( '<td>',"" )
          .replace( '<\/td>',"" );            
        }
        tdElements.unshift ( Utilities.formatDate(rowDate, "JST", "yyyy/MM/dd") );
        tagElement.push( tdElements );
        searchDay = rowDate;
      }
      if ( searchDay.getTime() <= startDay.getTime() ) {
        break;
      }    
    }
    if ( p < maxPageInKabutan ) {
      p++;
    } else {
      break;
    }
  }
  return tagElement;
}


function isBusinessDay(date){
  if (date.getDay() == 0 || date.getDay() == 6) {
    return false;
  }
  var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
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