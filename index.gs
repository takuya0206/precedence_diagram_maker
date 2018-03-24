/** Library
 * Moment.js  = key : MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48
 */
/**
 * @OnlyCurrentDoc
 */


Logger.log('Google Apps Script on...');

function onInstall(e) {
  onOpen(e);
}


function onOpen(e) {
  Logger.log('AuthMode: ' + e.authMode);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if(e && e.authMode == 'NONE'){
    menu.addItem('Getting Started', 'askEnabled');
  } else {
    var memo = PropertiesService.getDocumentProperties();
    var lang = Session.getActiveUserLocale();
    memo.setProperty('lang', lang);
    var createPD_text = lang === 'ja' ? 'プレシデンス・ダイアグラムの作成' : 'Create Precedence Diagram';
    var sidebar_text = lang === 'ja' ? 'サイドバーの表示' : 'Show Sidebar';
    menu.addItem(createPD_text, 'makeSampleProject');
    menu.addItem(sidebar_text, 'showSidebar');
  };
  menu.addToUi();
};


function onEdit(e) {
  Logger.log('onEdit start');
  var ss = getSpreadSheet()
  if (ss.getActiveSheet().getName() === 'list') {
    Logger.log('for the list sheet');
    var list = getListSheet();
    var editedRow = e.range.getRow();
    var editedColumn = e.range.getColumn();
    var editedLastRow = e.range.getLastRow();
    var editedLastColumn = e.range.getLastColumn();
    var lastRowOfContents = list.getLastRow();
    var lastColOfContents = list.getLastColumn();
    var baseRange = list.getRange(1, 1, lastRowOfContents, lastColOfContents);
    var baseData = baseRange.getValues();
    var selectedItem = baseData[1][editedColumn-1];

    //nothing happens if you edit the first and second row
    if(editedRow === 1 || editedRow === 2){
      return;
    };

    if(selectedItem === 'activity'){
      //when you enter a new activity, set an ID in ascending order
      if(e.value && !e.oldValue){
        var ary = [];
        var count = 0;
        for(var i = 2, len = baseData.length; i < len; i++){// 2 means starting after items
          var isUsed = false;
          count += 1;
          for(var j = 2; j < len; j++){
            if(count === baseData[j][0]){
              isUsed = true;
              break;
            };
          };
          if(!isUsed){ary.push(count);};
        };
        list.getRange(editedRow, 1, 1, 1).setValue(ary[0]);
      };
      //when you delete an activity, ID will be deleted
      if(e.range.isBlank()){
        list.getRange(editedRow, 1, editedLastRow-editedRow+1, 1).setValue('');
      };
    };
    if(selectedItem === 'precedentAct'){
      //when you set the beginning of activities, gray out relationship and lead/lag time
      var indexOfRelationship = baseData[1].indexOf('relationship');
      var range = list.getRange(editedRow,indexOfRelationship+1, 1, 2);
      if(e.value == '0'){
        range.setBackground('gray');
      };
      if(e.oldValue && e.oldValue == '0'){
        range.setBackground('');
      };
    };
    if(typeof e.value != 'object' && (selectedItem === 'precedentAct' || selectedItem === 'relationship' || selectedItem === 'L')){
      var ary = e.value.toString().replace(/\s+/g,'').split(',');
      var memo = PropertiesService.getDocumentProperties();
      var lang = memo.getProperty('lang');
      var text;
      e.range.clearNote().setBackground('');
      Logger.log(ary);
      Logger.log(selectedItem);
      //check whether precedentActs you enter exist
      if(selectedItem === 'precedentAct'){
        var indexOfId = baseData[1].indexOf('id');
        var targets = [];
        if(e.value && e.value != '0'){
          for(var i = 2, len = baseData.length; i < len; i++){
            if(i+1 != editedRow){
              targets.push(baseData[i][indexOfId].toString());
            };
          };
          for(var i = 0, len = ary.length; i < len; i++){
            if(targets.indexOf(ary[i]) < 0){
              text = lang === 'ja' ? '該当する先行タスクがありません' : 'The precedent ID you entered does not exist';
              e.range.setNote(text).setBackground('red');
            };
          };
        };
      };
      //check whether relationship you enter are FS, SS, SF or FF
      if(selectedItem === 'relationship'){
        var indexOfRelationship = baseData[1].indexOf('relationship');
        var val;
        for(var i = 0, len = ary.length; i < len; i++){
          val = ary[i];
          if(val != 'FS' && val != 'SS' && val != 'SF' && val != 'FF'){
            text = lang === 'ja' ? '値がFS, SS, SF, FFのいずれかではありません' : 'The value has to be FS, SS, SF, or FF.';
            e.range.setNote(text).setBackground('red');
          };
        };
      };
      //check whether lead/lag you enter is integer
      if(selectedItem === 'L'){
        var indexOfL = baseData[1].indexOf('L');
        for(var i = 0, len = ary.length; i < len; i++){
          if(!isNum(ary[i])){
            text = lang === 'ja' ? '値が数字ではありません' : 'The value has to be number.';
            e.range.setNote(text).setBackground('red');
          };
        };
      };
    };
  };
};