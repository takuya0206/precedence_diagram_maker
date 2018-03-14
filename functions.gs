function getSpreadSheet(){
  if(getSpreadSheet.ss){return getSpreadSheet.ss; };
  getSpreadSheet.ss = SpreadsheetApp.getActive();
  return getSpreadSheet.ss;
};


function getDiagramSheet(){
  var ss = getSpreadSheet();
  if(getDiagramSheet.d_sheet){return getDiagramSheet.d_sheet; };
  getDiagramSheet.d_sheet = ss.getSheetByName('diagram');
  return getDiagramSheet.d_sheet;
};


function getListSheet(){
  var ss = getSpreadSheet();
  if(getListSheet.w_sheet){return getListSheet.w_sheet; };
  getListSheet.w_sheet = ss.getSheetByName('list');
  return getListSheet.w_sheet;
};


function askEnabled(){
  var lang = Session.getActiveUserLocale();
  var title = 'Precedence Diagram Method';
  var msg = lang === 'ja' ? 'Precedence Diagram Methodが有効になりました。もしアドオンのメニューに「プレシデンス・ダイアグラムの作成」が表示されていない場合は一度リロードをお願いします。' : 'Precedence Diagram Method has been enabled. Just in case that the menu: "Create Precedence Diagram" does not appear, please reload this spreadsheet.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
};


function showSidebar() {
  Logger.log('showSidebar start');
  var html = HtmlService.createHtmlOutputFromFile('Page')
  .setTitle('Precedence Diagram Method')
  .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
};


function init(){
  Logger.log('init start');
  var memo = PropertiesService.getDocumentProperties();
  var lang = memo.getProperty('lang');
  var ss = getSpreadSheet();
  var list = getListSheet();
  var rowNum = list.getMaxRows();
  var listItems = lang === 'ja' ? [['ID', 'アクティビティ一覧','期間','先行ID', '関係性', 'リード / ラグ'],
  ['id', 'activity', 'duration', 'precedentAct', 'relationship', 'L']
  ] :
  [['ID', 'Activity List', 'Duration', 'Precedent ID', 'Relationship', 'Lead / Lag'],
  ['id', 'activity', 'duration', 'precedentAct', 'relationship', 'L']
  ];
  var listItemsLength = listItems[0].length;
  var range = list.getRange(1, 1, 2, listItemsLength);
  var firstRow = list.getRange(1, 1, 1, listItemsLength);
  range.setValues(listItems);
  firstRow.setBackground('#f3f3f3');
  list.getRange('A:A').setBackground('#f3f3f3');
  list.getRange('C:F').setHorizontalAlignment('center').setNumberFormat('@');
  list.setFrozenRows(1);
  list.hideRows(2);
  list.setColumnWidth(1, 35);
  list.setColumnWidth(2, 180);
  //note
  var text_precedentAct = lang === 'ja' ?
  '先行するアクティビティのIDを入力ください。もし複数ある場合はコンマ区切りで入力すること （例: 1,2）。また、一番最初のアクティビティの時は、0と入力し、「関係性」「リード/ラグ」の欄は空白にしてください。' : 
  'Enter the ID of precedent activitiy. If there are more than one, enter them in comma separated style. e.g. 1,2. When you enter the first activity, enter 0 and the Relationship and the Lead / Lag column should remain blank. ';
  var text_relationship = lang === 'ja' ?
  'FS (*Finish to Start), SS (*Start to Start), SF (*Start to Finish), FF (*Finish to Finish) のいずれかを入力ください。もし複数ある場合はコンマ区切りで入力すること （例: FS,SS）。':
  'Enter FS (*Finish to Start), SS (*Start to Start), SF (*Start to Finish) or FF (*Finish to Finish). If there are more than one, enter them in comma separated style. e.g. FS,SS.';
  var text_L = lang === 'ja' ?
  'リードタイムのときは負の数を、ラグタイムをのときは正の数を入力ください。もし複数ある場合はコンマ区切りで入力すること （例: 1,-2）':
  'When you enter lead time, use negative number. When you enter lag time, use positive number. If there are more than one, enter them in comma separated style. e.g. 1,-2.';
  list.getRange(1, 4, 1, 3).setNotes([[text_precedentAct, text_relationship, text_L]]);
  //dataValidation
  var rule_PosInt = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).build();
  list.getRange(3, 3, rowNum-3+1, 1).setNumberFormat(0).setDataValidation(rule_PosInt);

  showSidebar()
};


function isNum(val){
  val = String(val).replace(/^[ 　]+|[ 　]+$/g, '');
  if(val.length == 0){
    return false;
  };
  if(isNaN(val) || !isFinite(val)){
    return false;
  };
  return true;
};


function initForDiagram(){
  var diagram = getDiagramSheet();
  var colNum = diagram.getMaxColumns();
  diagram.setColumnWidth(1, 75);
  diagram.deleteColumns(2, colNum-2+1);
  diagram.insertColumnsAfter(1, 200);
};


function makeBox(row, col, id, title, duration, ES, EF, LS, LF){
  Logger.log('makeBox start');
  var diagram = getDiagramSheet();
  var range = diagram.getRange(row, col, 4, 4);
  var firstRow = diagram.getRange(row, col, 1, 4);
  var thirdRow = diagram.getRange(row+2, col, 1, 4);
  var contents = [['ID', id, 'Duration', duration],['ES', ES, 'EF', EF], [title,'','',''],['LS', LS, 'LF', LF]];
  var color = '';
  //show critical path in red color
  if(ES === LS && EF === LF){
    color = '#dc5b5b'; //red
  } else {
    color = '#AEAEAE'; //gray
  };
  range.setValues(contents);
  range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
  range.setHorizontalAlignment("center");
  firstRow.setBackground(color);
  firstRow.setFontColor('#ffffff');
  thirdRow.merge().setFontSize(14);
};


function box(pos, id, title, duration, precedentAct, precedentid, relationship, L, ES, EF, done) {
  this.pos = pos,
  this.id = id,
  this.title = title,
  this.duration = duration,
  this.precedentAct = precedentAct,
  this.precedentid = precedentid,
  this.relationship = relationship,
  this.L = L,
  this.ES = ES,
  this.EF = EF,
  this.done = done
};


function analyzeList(data){
  var acts = {};
  var calculatedActs = {};
  var indexOfPreAct = data[1].indexOf('precedentAct');
  var indexOfDuration = data[1].indexOf('duration');
  var indexOfRelationship = data[1].indexOf('relationship');
  var indexOfL = data[1].indexOf('L');
  var len = data.length;
  var isValid = true;
  var ary = [];
  var id = null;
  var col = null;
  var row = 0;
  var lastId, isMatch, isAvailable, precedentAct, emptyRow;

  //make ary for multiple precedentActs and relationship
  for (var i = 2; i < len; i++){
    //remove half-width space
    data[i][indexOfPreAct] = data[i][indexOfPreAct].toString().replace(/\s+/g,'').split(',');
    data[i][indexOfRelationship] = data[i][indexOfRelationship].toString().replace(/\s+/g,'').split(',');
    data[i][indexOfL] = data[i][indexOfL].toString().replace(/\s+/g,'').split(',');
    if(data[i][0] == ''){
      emptyRow = i;
      break;
    };
  };
  //delete values after the empty row
  if(emptyRow){
    var list = getListSheet();
    data.splice(emptyRow, len-emptyRow);
    len = data.length;
    list.getRange(emptyRow, 1, 1, data[0].length).setBorder(null, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);//show the border
  };
  //check validation
  checkLength(data, {'PreAct': indexOfPreAct, 'relationship': indexOfRelationship, 'L' : indexOfL});
  checkUnusedId(data, indexOfPreAct);
  //find the beginning of a project
  for(var i = 2; i < len; i++){
    if(data[i][indexOfPreAct][0] == 0){
      acts[data[i][0]] = new box({'row': 0, 'col': 0}, data[i][0], data[i][1], data[i][indexOfDuration], false, false, false, false, 0, parseInt(data[i][indexOfDuration]), false);
    };
  };
  while(true){//iteration will be end by break
    id = null;
    col = null;
    row = 0
    //find acts which are not done
    for(key in acts){
      if(acts[key]['done'] === false){
        acts[key]['done'] = true;
        row = acts[key]['pos']['row']+1;
        id = key;
        lastId = key;
        break;
      };
    };
    Logger.log('Current Activity: ' + id);
    //break for while
    if(!id){
      if(len-2 !== Object.keys(acts).length){
        Browser.msgBox('Procedent acts are not properly set.');
        break;
      } else {
        Logger.log('Successfully done.');
        break;
      };
    };
    //crawl the list
    for(var j = 2; j < len; j++){
      isMatch = false;
      isAvailable = true;
      var len2 = data[j][indexOfPreAct].length;
      //find precedent acts to the current Activity
      for(var k = 0; k < len2; k++){
        if(id == data[j][indexOfPreAct][k]){
          isMatch = true;
        };
      };
      //check whether there are already necessary acts
      if(isMatch){
        for(var k = 0; k < len2; k++){
          if(data[j][indexOfPreAct][k] in acts == false){
            isAvailable = false;
          };
        };
      };
      if(isMatch && isAvailable){
        Logger.log('Following Activity: ' + data[j][0]);
        //decide the proper ES/EF
        var preES = null;
        var preEF = null;
        for(var k = 0; k < len2; k++){
          var ES = 0;
          var EF = 0;
          var L = parseFloat(data[j][indexOfL][k]);
          var duration = parseFloat(data[j][indexOfDuration]);
          var relationship = data[j][indexOfRelationship][k];
          //calculate ES and EF based on the relationship and the lead/lag time
          if (relationship === 'FS'){
            ES = parseFloat(acts[data[j][indexOfPreAct][k]]['EF']) + L;
            EF = ES + duration;
          } else if (relationship === 'FF'){
            EF = parseFloat(acts[data[j][indexOfPreAct][k]]['EF']) + L;
            ES = EF - duration;
          } else if (relationship === 'SF'){
            EF = parseFloat(acts[data[j][indexOfPreAct][k]]['ES']) + L;
            ES = EF - duration;
          } else if (relationship === 'SS'){
            ES = parseFloat(acts[data[j][indexOfPreAct][k]]['ES']) + L;
            EF = ES + duration;
          };
          //compare LF with existing one and input proper one
          if(!preES){
            preES = ES;
            preEF = EF;
          } else {
            if(preES <= ES){
              preES = ES;
              preEF = EF;
            };
          };
        };
        //round two decimal fraction
        preES = Math.round(preES * 100) / 100;
        preEF = Math.round(preEF * 100) / 100;
        //adjust positions
        var positions = [];
        var adjustedCol = 0;
        var smallestRow = 0;
        var isObstacle = false;
        var isOverlap = false;
        //pick up the position of the precident acts
        for(var k = 0, len2 = data[j][indexOfPreAct].length; k < len2; k++){
          positions.push(acts[data[j][indexOfPreAct][k]]['pos']);
        };
        //the col of the current Activity should be under the previous one
        positions.sort(function(a, b){return b['col'] - a['col'];}); //right-aligned
        if(positions.length > 1 && positions[0]['col'] === positions[1]['col']){
          Logger.log('move to the next col if present acts are in the same col');
          adjustedCol = positions[0]['col'] + 1; //move to the next col if present acts are in the same col
        } else {
          Logger.log('under the previous one');
          adjustedCol = positions[0]['col'];
        };
        //if there are acts between the rows, move the position to the next col
        positions.sort(function(a, b){return parseInt(a['row']) - parseInt(b['row']);});
        smallestRow = positions[0]['row'];
        if (row - smallestRow > 1){
          for(var key in acts){
            if(data[j][0] != acts[key]['id'] && acts[key]['pos']['col'] === adjustedCol && acts[key]['pos']['row'] < row && acts[key]['pos']['row'] > smallestRow){
              isObstacle = true
            };
          };
          if(isObstacle) {
            Logger.log('isObstacle: true')
            adjustedCol += 1;
          };
        };
        //if there are more than two in the same cycle, col should follow the previous one
        if(col !== null){
           col += 1;
        } else {
          col = adjustedCol;
        };
        //if the position is overlaped with another, move the whole col to the next
        Object.keys(acts).forEach(function(key){
         if(acts[key]['pos']['row'] === row && acts[key]['pos']['col'] === col && acts[key]['id'] !== data[j][0]){
           Logger.log('Overlap: '+ acts[key]['id']);
           isOverlap = true;
         };
        });
        if(isOverlap){
          for(key in acts){
            if(acts[key]['pos']['col'] >= col && acts[key]['pos']['row'] < row){
              acts[key]['pos']['col'] += 1;
            };
          };
          col += 1;
        };
        //input
        Logger.log('row: ' + row);
        Logger.log('col: ' + col);
        Logger.log('ES: ' + preES);
        Logger.log('EF: ' + preEF);
        acts[data[j][0]] = new box({'row': row, 'col': col}, data[j][0], data[j][1], duration, positions, data[j][indexOfPreAct], data[j][indexOfRelationship], data[j][indexOfL], preES, preEF, false);
      };
    };
  };
  //backforward scheduling
  acts[lastId]['LF'] = acts[lastId]['EF'];
  acts[lastId]['LS'] = acts[lastId]['LF'] - acts[lastId]['duration'];
  calculatedActs[lastId] = acts[lastId];

  while(true){//iteration will be end by break
    id = null;
    ary = [];
    for(key in calculatedActs){
      if(calculatedActs[key]['done'] == true){
        ary.push(calculatedActs[key]);
      };
    };
    //break when all activities are done
    if(ary.length === 0){break;};
    //calculate them from the end
    ary.sort(function(a, b){return b['pos']['row']-a['pos']['row'];});
    precedentAct = ary[0]['precedentid'];
    id = ary[0]['id'];
    calculatedActs[id]['done'] = false;

    Logger.log('Current Activity: ' + id);
    Logger.log('Following Activity: ' + precedentAct);

    for (var i = 0, len2 = precedentAct.length; i < len2; i++){
      //calculate ES and EF based on the relationship and the lead/lag time
      var LS = 0;
      var LF = 0;
      var duration = parseFloat(acts[precedentAct[i]]['duration']);
      var L = parseFloat(acts[id]['L'][i]);
      var relationship = acts[id]['relationship'][i];
      if (relationship === 'FS'){
        LF = acts[id]['LS'] - L;
        LS = LF - duration;
      } else if (relationship === 'FF'){
        LF = acts[id]['LF'] - L;
        LS = LF - duration;
      } else if (relationship === 'SF'){
        LS = acts[id]['LF'] - L;
        LF = LS + duration;
      } else if (relationship === 'SS'){
        LS = acts[id]['LS'] -L;
        LF = LS + duration;
      };
      //compare LF with existing one and input proper one
      if(!acts[precedentAct[i]]['LF']){
        acts[precedentAct[i]]['LF'] = Math.round(LF * 100) / 100;
        acts[precedentAct[i]]['LS'] = Math.round(LS * 100) / 100;
      } else {
        if(acts[precedentAct[i]]['LF'] >= LF){
          acts[precedentAct[i]]['LF']  = Math.round(LF * 100) / 100;
          acts[precedentAct[i]]['LS'] = Math.round(LS * 100) / 100;
        };
      };
      //input
      Logger.log('LS: ' + LS);
      Logger.log('LF: ' + LF);
      calculatedActs[precedentAct[i]] = acts[precedentAct[i]];
    };
  };
  return calculatedActs;
};


function checkLength(data, indexs){
  var list = getListSheet();
  var isValid = true;
  for(var i = 2, len = data.length; i < len; i++){
    var isSame = false;
    for(var key in indexs){
      //skip the roop if the beginning of the diagram
      if(data[i][indexs.PreAct] == 0){
        break;
      };
      //if a cell is empty throw false
      if(data[i][indexs[key]][0] == ''){
        list.getRange(i+1, indexs[key]+1).setBackground('red');
        isValid = false;
      }
      //compare length
      if(isSame && isSame !== data[i][indexs[key]].length){
        list.getRange(i+1, indexs[key]+1).setBackground('red');
        isValid = false;
      };
      isSame = data[i][indexs[key]].length;
    };
  };
  if(isValid){
    return true;
  } else {
    throw new Error('There is invaild input. Please check red cells');
  };
};


function checkUnusedId(data, index){
  var id = [];
  var precedentId = [];
  var unused = [];
  for(var i = 2, len = data.length; i < len; i++){
    //input id
    id.push(data[i][0]);
    //input precedent id
    for(var j = 0, len2 = data[i][index].length; j < len2; j++){
      precedentId.push(parseInt(data[i][index][j]));
    };
  };
  //compare id with precedent id
  for(var i = 0, len = id.length; i < len; i++){
    var isThere = precedentId.indexOf(id[i]);
    if(isThere < 0){
      unused.push(id[i])
    };
  };
  if(unused.length > 1){
    throw new Error('Error: ID ' + unused + ' are not used.');
  } else {
    return true;
  };
};


function makeSampleProject(){
  Logger.log('start makeSampleProject');
  var list = getListSheet();
  var memo = PropertiesService.getDocumentProperties();
  var lang = memo.getProperty('lang');
  var isComfirmed = true;
  var data = lang === 'ja' ?
  [[1, '要件定義', 3, 0, '', ''], [2, '設計', 5, 1, 'FS', 0], [3, 'デザイン', 3, 2, 'FS', 0], [4, 'コーディング', 5, 3, 'FS', 0], [5, 'サーバー環境構築', 3, 2, 'FS', 0], [6, '公開', 1, '4,5', 'FS,FS', '0,0']]:
  [[1, 'Requirement Definition', 3, 0, '', ''], [2, 'Basic Design', 5, 1, 'FS', 0], [3, 'UI Design', 3, 2, 'FS', 0], [4, 'Coding', 5, 3, 'FS', 0], [5, 'Building Server', 3, 2, 'FS', 0], [6, 'Release', 1, '4,5', 'FS,FS', '0,0']];
  var title_text = lang === 'ja' ? 'サンプルWebサイトの作成' : 'Creating Sample Website';
  var msg = lang === 'ja' ? '既にListシートが存在します。これまでの内容を消して、新たに作成を行いますか？' : 'The List sheet already exists. Will you delete the existing one and create new one?';
  if(!list){
    ss.insertSheet('list', 1);
    list = getListSheet();
  } else {
    isComfirmed = Browser.msgBox(msg, Browser.Buttons.YES_NO);
    if(isComfirmed === 'yes'){
      list.clear();
    } else {
      isComfirmed = false;
    };
  };
  if(isComfirmed){
    list.getRange(3,1,data.length, data[0].length).setValues(data);
    init();
    createDiagram(title_text);
  };
};

