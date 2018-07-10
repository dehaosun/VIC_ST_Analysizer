function onEdit(e){
  var rng = e.range;
  processChange(rng);
}

function onOpen(e){
  processReload();
};


function uniTest(){  
  processSyncFromRecorder();
}

function processReload(){
  
  changeStatusDirectly(3, true);
  //check if Stock NAME is empty, set a default value
  var stockRng = getRangeByNamedTag('_I_ST_NAME');
  var stockName = stockRng.getValue();
  if(stockName === null || stockName.trim() === ''){
    stockRng.setValue('FB');
  }
  
  //SyncAll
  syncCacheAll();
  
  //LoadData
  processSyncFromRecorder();
  changeStatusDirectly(0, true);
}


function changeStatus(value, wait, fromCache){
  var tag = '_CFG_C_STATUS';
  if(fromCache){
    var rng = getCachedRangeByNamedTag(tag);
  }else{
    var rng = getRangeByNamedTag(tag);
  }
  rng.setValue(value);
  
  if(wait){
    SpreadsheetApp.flush();
  }
}

function changeStatusFromCache(value, wait){
  changeStatus(value ,wait,true);
}

function changeStatusDirectly(value, wait){
  changeStatus(value ,wait, false);
}


function processChange(rng){
  
  syncCacheAuto();
  var range = rng;
  var namedTag = getNamedTagFromRange("_I_",range);
   
  if(namedTag !== null){
    
    
  
    if(namedTag === "_I_ST_NAME"){
      
      range.activate()
      syncCache_StockChange();  
      var stockName = getCache_NamedTagValue('_I_ST_NAME');
      //if not empty and valid
      if( stockName !== null && stockName.trim() !==''){
        //Sync if StockNameChange    
        //change stock
        processSyncFromRecorder();
      }else{
        //clean all value?
        //Logger.log('Empty StockName.');
      }
      
      
    }else{
      //if changed cell is in recorder
      processSyncToRecorderByNamedTag(namedTag);
    }
     
  }
  
}




function processSyncToWatchList(){
  
  changeStatusDirectly(2,true);
  syncCacheAuto();
//  syncDataToRecorder();  
  syncDataToWatchList();     
  changeStatusDirectly(0,true);
}

function processSyncToRecorderByNamedTag(namedTag){  
  changeStatusDirectly(1,false);  
  syncDataToRecorderByNamedTag(namedTag);     
  changeStatusDirectly(0,false);
}

function processSyncFromRecorder(){
  changeStatusFromCache(-1,true);
  syncRecorderToData();     
  changeStatusFromCache(0,true);
}

function processSyncToRecorder(){
  syncCacheAuto();
  changeStatusFromCache(1,true);  
  syncDataToRecorder();     
  changeStatusFromCache(0,true);
}

function syncDataToRecorderByNamedTag(namedTag){
  syncDataToSheetByNamedTag('ListRecorder',namedTag);
}


function syncDataToSheetByNamedTag(shName,namedTag){

  //put value from range to sh
  var stockName = getCache_NamedTagValue('_I_ST_NAME');
  if(namedTag === '_I_ST_NAME'){return;}
  
  //get row
  var trgRow = getStockLocation(stockName,shName); 
  if(trgRow<=0){  
    trgRow = getNewStockLocation(stockName,shName);
  }
  
  //get col
  var trgCol = getCache_NamedTagColumn(shName, namedTag);
  var trgSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName);
  
  //src data
  var namedRng = getCachedRangeByNamedTag(namedTag)
  if( namedRng!== null && trgSh !== null){
    namedRng.copyValuesToRange(trgSh,trgCol,trgCol,trgRow,trgRow);
  }
  
}

function syncDataToRecorder(){
  syncDataWithSheet('ListRecorder',1);
}


function syncDataToWatchList(){
  syncDataWithSheet('WatchList',1);
}

function syncRecorderToData(){
  syncDataWithSheet('ListRecorder',-1);
}

function syncDataWithSheet(shName,direct){

  var stockName = getCache_NamedTagValue('_I_ST_NAME');
  var recorderRow = getStockLocation(stockName,shName);
  var newStock = false;
  
  if(recorderRow<=0){
    //create mode
    
    recorderRow = getNewStockLocation(stockName,shName);
    newStock =true;
  }
  
  
  var cache = CacheService.getDocumentCache();
  var rngTags = getNamedTagSetBySheet(shName);
  
  if( rngTags === null){
    return;
  }
  
  for( var i = 0 ; i<rngTags.length; i++){
    var curTag =rngTags[i];
    if( newStock ){
      if(curTag === '_I_ST_NAME' && direct == -1){
      continue;
      }
    }
    var key = shName+'!'+ curTag + '.Column';
    var curColumnS = cache.get(key);
    if(curColumnS === null ){
      continue;
    }
    var curColumn = Number(curColumnS);
    
    var key2 = 'NamedTag.'+ curTag;
    var addressString = cache.get(key2);
    if(addressString !== null ){
      var tmpStrings = addressString.split('!');
      var namedShName = tmpStrings[0];
      var namedAddress = tmpStrings[1];
      var namedCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namedShName).getRange(namedAddress)   
    }else{
      var namedCell = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(curTag);
    }
    
    var valueRng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange(recorderRow,curColumn);  
//    var valueRngAddress = valueRng.getA1Notation();
//    var namedCellAddress = namedCell.getA1Notation();
//    var toBeValue = valueRng.getValue();
    
    if(direct == -1){
      var srcSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName);
      var srcCell = valueRng;    
      var trgSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namedShName);
      var trgCell = namedCell;      
    }else if(direct == 1){
      var srcSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namedShName);
      var srcCell = namedCell;    
      var trgSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName);
      var trgCell = valueRng;     
    }else{
      return;
    }
    
    var trgColumn = trgCell.getColumn();
    var trgRow = trgCell.getRow(); 
    
    
    //method 1 not working if cell contain formula
    //srcCell.copyTo(trgCell, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    //method 2 working working
    //var toBeValue = srcCell.getValue();
    srcCell.copyValuesToRange(trgSh,trgColumn,trgColumn,trgRow,trgRow);

  }
}

function getRangeByNamedTag(curTag){
  var rng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(curTag);
  return rng;
}





function getCachedRangeByNamedTag(curTag){
  var cache = CacheService.getDocumentCache();
  var key2 = 'NamedTag.'+ curTag;
  
  var addressString = cache.get(key2);
  if(addressString !== null ){
    var tmpStrings = addressString.split('!');
    var shName = tmpStrings[0];
    var address = tmpStrings[1];
    var namedCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange(address);
    return namedCell;
  }else{
    return null;
  }

}

function getNamedTagSetBySheet(shName){
  var cache = CacheService.getDocumentCache();
  var key = shName +'!*.NamedTag';
  var rngTagString = cache.get(key);
  var rngTags =  rngTagString.split(',');
  return rngTags;
}


function getNamedTagFromRange( prefix, trgRng){
  
  var cache = CacheService.getDocumentCache();
  
  var shName = trgRng.getSheet().getName();
  var address = trgRng.getA1Notation();
  
  var key = shName + "!" + address;
  
  var curTag = cache.get(key);
  
  if( curTag !== null && ( prefix === null ||curTag.indexOf(prefix)>=0 ) ){
   return curTag;  
  }
  return null;
}



function getStockLocation(stockName,shName) {  
  
  //hard code here
  if(shName === 'WatchList'){
    var keyTag = 'W';
  }else{
    var keyTag = 'C';
  }
  var curTag = '_CFG_'+keyTag+'_ST_C_ROW';
  
  var curRow = -1;
  var curRowS = getCache_NamedTagValue(curTag);
  
  if(curRowS !== null ){    
    curRow = Number(curRowS);  
  }
  
//  if(curRow <=0){
//    
//   var searchRng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange("B:B");
//   var numRows = searchRng.getNumRows();
//   var numCols = searchRng.getNumColumns();
//  
//    for (var i = 1; i <= numRows; i++) {
//      for (var j = 1; j <= numCols; j++) {
//        var curRng = searchRng.getCell(i,j)     
//        if(!curRng.isBlank() && curRng.getValue() === stockName){        
//          return i;
//        } 
//      }
//    }
//  }
 
  return curRow;
}

function getNewStockLocation(stockName,shName) {

  

  //hard code here
  if(shName === 'WatchList'){
    var keyTag = 'W';
  }else{
    var keyTag = 'C';
  }
  var curTag = '_CFG_'+keyTag+'_ST_NEW_ROW';
  
  var curRow = -1;
  var curRowS = getCache_NamedTagValue(curTag);
  
  if(curRowS !== null ){    
    curRow = Number(curRowS);
    var curCol = getCache_NamedTagColumn(shName,'_I_ST_NAME');   
    var curCell = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange(curRow, curCol);
    curCell.setValue(stockName)    
  }
  
  
//  if(curRow <= 0){
//   
//    var searchRng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange("B:B");
//    
//    var numRows = searchRng.getNumRows();
//    var numCols = searchRng.getNumColumns();
//    var locIdx = -1;
//    
//    for (var i = 1; i <= numRows; i++) {
//      for (var j = 1; j <= numCols; j++) {
//        var curRng = searchRng.getCell(i,j)     
//        if(curRng.isBlank()){        
//          //curRng.setValue(stockName);
//          locIdx = i;
//          return locIdx;
//        } 
//      }
//    }
//    
//    if(locIdx == -1){
//      
//      var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName);
//      var lastrow = sh.getLastRow(); 
//      sh.insertRowAfter(lastrow);
//      locIdx = lastrow + 1;
//      //searchRng.getCell(numRows, numCols).offset(1, 0).setValue(stockName);
//      return locIdx;
//    }  
//    
//  }
  
  return curRow;
}




function syncCache_StockChange(){
  var stockName = getCachedRangeByNamedTag('_I_ST_NAME').getValue().toUpperCase();
  syncCache_NamedTagValue('_I_ST_NAME',stockName);
  
  var curTag = '_CFG_C_ST_C_ROW';
  syncCache_NamedTagValue(curTag,getCachedRangeByNamedTag(curTag).getValue());
  
  var curTag = '_CFG_C_ST_NEW_ROW';
  syncCache_NamedTagValue(curTag,getCachedRangeByNamedTag(curTag).getValue());
  
  var curTag = '_CFG_W_ST_C_ROW';
  syncCache_NamedTagValue(curTag,getCachedRangeByNamedTag(curTag).getValue());
  
  var curTag = '_CFG_W_ST_NEW_ROW';
  syncCache_NamedTagValue(curTag,getCachedRangeByNamedTag(curTag).getValue());
  
}




function syncCacheAuto(){
  var cache = CacheService.getDocumentCache();
  var flag = cache.get('_Done_');
  if(flag === null){
    syncCacheAll();
  }
}






function syncCacheAll(){
  changeStatusDirectly(3,true);

  syncCache_NamedRngToSheetCol('ListRecorder');
  syncCache_NamedRngToSheetCol('WatchList');
  syncCache_AddressToNamedRng();
  var cache = CacheService.getDocumentCache();
  
  syncCache_StockChange();
  
  //var stockName2 = getCache_NamedTagValue('_I_ST_NAME')
  cache.put('_Done_','1',86400);
  changeStatusFromCache(0,true);
}


function getCache_NamedTagColumn(shName, namedTag){
  
  var cache = CacheService.getDocumentCache();
  var key1 = shName +'!'+ namedTag + '.Column';
  var colStr = cache.get(key1);
  var rtn = -1;
  
  if(colStr !== null){
    rtn = Number(colStr);
  }else{
    var tagRng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListRecorder').getRange('1:1')
    var numRows = tagRng.getNumRows();
    var numCols = tagRng.getNumColumns();
  
    for (var i = 1; i <= numRows; i++) {
      for (var j = 1; j <= numCols; j++) {
        var curRng = tagRng.getCell(i,j)     
        if(!curRng.isBlank() && curRng.getValue()== namedTag){        
          return j;
        }
        
      }
    }   
  }
   
  return rtn;
}


function getCache_NamedTagValue(tagName){
  var cache = CacheService.getDocumentCache();
  var key2 = 'NamedTag.'+ tagName + '.value';
  var value = cache.get(key2);
  return value;
}


function syncCache_NamedTagValue(tagName,value){
  var cache = CacheService.getDocumentCache();
  var key2 = 'NamedTag.'+ tagName + '.value';
  cache.put(key2,value,86400);
}


function syncCache_NamedRngToSheetCol(shName){
  var cache = CacheService.getDocumentCache();
  var tagRng = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(shName).getRange('1:1')
  if(tagRng === null){return;}
  var numRows = tagRng.getNumRows();
  var numCols = tagRng.getNumColumns();
  var scanedTag = '';
  var scanedCol = '';
  
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var curRng = tagRng.getCell(i,j)     
      if(!curRng.isBlank()){
        var curTag = curRng.getValue();
        var testRng = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(curTag);
        if(testRng === null ){
          //color rng
          if(curTag.indexOf('_')>=0){
            curRng.setBackground('#ff0000').setFontColor('#ffffff');
          }
        }else{
                  
          var key1 = shName +'!'+ curTag + '.Column';
          var key2 = shName +'!'+ j + '.NamedTag';
          
          cache.put(key1, j,86400);
          cache.put(key2, curTag,86400);
          
          //append scanTag
          if(scanedTag === ''){
            scanedTag = curRng.getValue()
            scanedCol = j;
          }else{
            scanedTag = scanedTag +","+ curTag;
            scanedCol = scanedCol +","+ j;
          }
          
        }
      }   
    }
  }
  
  if( scanedTag !==''){
    cache.put(shName +'!*.Column',scanedCol,86400);
    cache.put(shName +'!*.NamedTag',scanedTag,86400);
    
  }


}


function syncCache_AddressToNamedRng(){
  var cache = CacheService.getDocumentCache();
  var namedRngs = SpreadsheetApp.getActiveSpreadsheet().getNamedRanges();
  for(var j=0;j<namedRngs.length;j++){
    var namedTag = namedRngs[j].getName();
    var namedRng = namedRngs[j].getRange();
    var shName = namedRng.getSheet().getName();
    var address = namedRng.getA1Notation();
    var key = shName + "!" + address
    var key2 = 'NamedTag.'+ namedTag;
    cache.put(key,namedTag,86400);
    cache.put(key2,key,86400);
  }

}


function syncStockInfo(){
  
  var trgSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stock_TmpInfo');
  
  var tag = '_CFG_ST_INFO_URL';
  var rng = getRangeByNamedTag(tag);
  
  var url = rng.getValue();
  
  var srcWB = SpreadsheetApp.openByUrl(url);
  var srcSh = srcWB.getSheetByName('Summary');
  
  
  
  var SRange = srcSh.getDataRange();

  // get A1 notation identifying the range
  var A1Range = SRange.getA1Notation();

  // get the data values in range
  var SData = SRange.getValues();

  // Clear the Google Sheet before copy
  trgSh.clear({contentsOnly: true});

  // set the target range to the values of the source data
  trgSh.getRange(A1Range).setValues(SData);
  
  
  
  
}


