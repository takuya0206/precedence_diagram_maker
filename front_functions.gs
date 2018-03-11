function createDiagram(){
  var list = getListSheet();
  var diagram = getDiagramSheet();
  if(!diagram){
    var ss = getSpreadSheet();
    ss.insertSheet('diagram', 2);
    initForDiagram();
    diagram = getDiagramSheet()
  };
  var data = list.getRange(1, 1, list.getLastRow(), list.getLastColumn()).getValues();
  var startRow = 4;
  var startCol = 2;
  var rowBase = 8;
  var colBase = 6;
  var data = analyzeList(data);
  diagram.clear();
  for(key in data){
    var row = data[key]['pos']['row']*rowBase + startRow;
    var col = data[key]['pos']['col']*colBase + startCol;
    makeBox(row, col, data[key]['id'], data[key]['title'], data[key]['duration'], data[key]['ES'], data[key]['EF'], data[key]['LS'], data[key]['LF']);
    //draw lines
    for(var i = 0, len = data[key]['precedentAct'].length; i < len; i++){
      var startPos = {'row': data[key]['precedentAct'][i]['row']*rowBase + startRow+4, 'col': data[key]['precedentAct'][i]['col']*colBase + startCol+2 };
      var endPos = {'row': row, 'col': col+2 };
      var text = '';
      var count = 0;
      var partition ='';
      //go straight
      if(startPos.col === endPos.col){
        diagram.getRange(startPos.row, startPos.col, endPos.row-startPos.row, 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      };
      //turn right
      if(startPos.col < endPos.col){
        //from the corner
        diagram.getRange(startPos.row, startPos.col+2, 2, 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        diagram.getRange(startPos.row+2, startPos.col+2, endPos.row-startPos.row-2, endPos.col-startPos.col-2).setBorder(true, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        //from the center
        //diagram.getRange(startPos.row, startPos.col, 2, 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        //diagram.getRange(startPos.row+2, startPos.col, endPos.row-startPos.row-2, endPos.col-startPos.col).setBorder(true, null, null, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      };
      //turn left
      if(startPos.col > endPos.col){
        //from the corner
        diagram.getRange(startPos.row, startPos.col-2, 2, 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        diagram.getRange(startPos.row+2, endPos.col, endPos.row-startPos.row-2, startPos.col-endPos.col-2).setBorder(true, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        //from the center
        //diagram.getRange(startPos.row, startPos.col, 2, 1).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
        //diagram.getRange(startPos.row+2, endPos.col, endPos.row-startPos.row-2, startPos.col-endPos.col).setBorder(true, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      };
      //make a text
      for(var j = 0, len2 = data[key]['relationship'].length; j < len2; j++){
        var isPositive = '';
        if(parseFloat(data[key]['L'][j]) > 0){
          isPositive = '+';
        };
        partition = count > 0 ? '  /  ' : '';
        text += partition + '(ID_' + data[key]['precedentid'][j] + ') R: ' + data[key]['relationship'][j] + ', L: ' + isPositive + data[key]['L'][j];
        count += 1;
      };
      diagram.getRange(endPos.row-1, endPos.col, 1, 1).setValue(text);
    };
  };
};