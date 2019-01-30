/**
    This script is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    any later version.

    This script is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along.  If not, see <https://opensource.org/licenses/GPL-3.0>.
 **/

function goToCell(rng) {
  var sh = sheet.getSheet(); 
  SpreadsheetApp.setActiveSheet(sh);
  SpreadsheetApp.setActiveRange(rng);
}

function getSheetUrl() {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId(); 
  return url;
}

function getSheetUrlByName(shName) {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getSheetByName(shName);
  var url = '';
  url += SS.getUrl();
  url += '#gid=';
  url += ss.getSheetId(); 
  return url;
}



function columnToLetter(column)
{
  var temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function getAddress(e){
  var rng = e.range;
  var address = range.getA1Notation(); 
  return address;
}



function test(){
  openUrl("http://www.google.com")
}
//
function openUrl( url ){
  var html = HtmlService.createHtmlOutput('<html><script>'
  +'window.close = function(){window.setTimeout(function(){google.script.host.close()},9)};'
  +'var a = document.createElement("a"); a.href="'+url+'"; a.target="_blank";'
  +'if(document.createEvent){'
  +'  var event=document.createEvent("MouseEvents");'
  +'  if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'                          
  +'  event.initEvent("click",true,true); a.dispatchEvent(event);'
  +'}else{ a.click() }'
  +'close();'
  +'</script>'
  // Offer URL as clickable link in case above code fails.
  +'<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="'+url+'" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>'
  +'<script>google.script.host.setHeight(40);google.script.host.setWidth(410)</script>'
  +'</html>')
  .setWidth( 90 ).setHeight( 1 );
  SpreadsheetApp.getUi().showModalDialog( html, "Opening ..." );
}

function specialUrl(){
  showAnchor('清除Jerry血液','https://www.youtube.com/watch?v=9bxyezJ_5BE');
}

function showAnchor(name,url) {
  var html = '<html><body><a href="'+url+'" target="blank" onclick="google.script.host.close()">'+name+'</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
  SpreadsheetApp.getUi().showModelessDialog(ui," ");
}


function gplUrls(){
  var map = new Array();

  map[0]=  
    {
    Name: 'This googlesheet is free software: <br/>' +
    'you can redistribute it and/or modify <br/>' +
    'it under the terms of the GNU GPL, <br/>' +
    'either version 3 or later of the License. <br/>' +
    'This googlesheet is distributed in the hope that it will be useful, <br/>' +
    'but WITHOUT ANY WARRANTY; <br/>'+
    'Click for more details.'  
    ,Url: 'https://opensource.org/licenses/GPL-3.0'
    };
  
    showAnchors(map);
}



function specialUrls(){
  
  //TODO: read name url form config sheet
  var map = new Array();

  map[0]= 
    {
    Name: '我頭好熱～～'    
    ,Url: 'https://www.youtube.com/watch?v=9bxyezJ_5BE'
    };
  
  
  map[1]= 
    {
    Name: '我要努力向上'    
    ,Url: 'https://valueinvestingcollege.tw/blog/'
    };
  
  map[2]= 
    {
    Name: '我想被醍醐灌頂'    
    ,Url: 'https://www.facebook.com/groups/taiwanvic/learning_content/'
    };
    
  map[3]= 
    {
    Name: '我想砍掉重練或造福世人'    
    ,Url: 'http://valueinvestingcollege.tw/schedule/'
    };
  
  showAnchors(map);
}

function showAnchors(map) {
  var html_prefix = '<html><body>'
  var html_suffix ='</body></html>';
  var cnt = map.length;
  var html_content = '';
  for( var i = 0; i < cnt ; i++){
    var cur_link = map[i];
    var tmp_html = '<a href="'+ cur_link.Url +'" target="blank" onclick="google.script.host.close()">'+ cur_link.Name +'</a><p>';
    html_content = html_content + tmp_html;
  }
  var html_final = html_prefix + html_content + html_suffix;
  var ui = HtmlService.createHtmlOutput(html_final);
  SpreadsheetApp.getUi().showModelessDialog(ui," ");
}