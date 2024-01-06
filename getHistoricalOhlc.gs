/**
 * get ohlc stock price  
 * @param  {string} code Stock code of Japan.
 * @param  {Date} startDate A date to start getting stock price. 
 * @param  {Date} endDate    A date to end getting stock price. 
 * @return {array} Array of ohlc stock price.
 */
function getHistoricalOhlcFromKabutan(code,startDate,endDate){

  /**
   * set default date if no arguments   
   */
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
  var page = 1;
  var ohlc = [];
  var lastSearchDate = endDate;
  
  while ( 
     lastSearchDate.getTime() > startDate.getTime()　
  && page < maxPageInKabutan
  ){
    var getUrl = 'https://kabutan.jp/stock/kabuka?code=' + code + '&ashi=Date&page=' + page;
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

      if ( date.getTime() < startDate.getTime() ) {
        break;
      }    

      if ( date.getTime() > lastSearchDate.getTime() ) {
        continue;
      }    
      
      var tdElements = stockRow[i].match( /<td>([\d|.|,]*?)<\/td>/g )
      for(var j=0 ; j < tdElements.length ; j++){
        tdElements[j] = tdElements[j]
        .replace( '<td>',"" )
        .replace( '<\/td>',"" )
        .replace( ',',"" ); 
      }
      tdElements.unshift ( Utilities.formatDate(date, "JST", "yyyy/MM/dd") );
      ohlc.push( tdElements );
      lastSearchDate = date;
    }
    page++;
  }
  return ohlc.reverse();
}
