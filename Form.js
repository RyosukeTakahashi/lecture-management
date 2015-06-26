//メインシートにおいて、科目リスト、生徒リストに変更があった場合それをフォームに反映する。
//Trigger:フォームを開く度に実行
function contentsList() {

  //シートをURLで開く
  var ssSchedule = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/19Z9nYpSRL-hECQPLMS50i_THJ2kUMYgo1IBojMpY-mI/edit#gid=1004063562');
  //一番目のシート取得
  var sheetContents = ssSchedule.getSheets()[1];
  //生徒リストを取得（2次元配列[][]）
  var studentList = sheetContents.getRange("1:1").getValues();
  //科目リスト取得
  var contentsList = sheetContents.getRange("A:A").getValues();

  //Formを開き、アイテム一覧取得
  var form = FormApp.openByUrl('https://docs.google.com/forms/d/1-V5yEgPHIN04oZIQzT5GPk_T6lJPMDdnXu3n_aQYXqw/edit');
  var items = form.getItems();
  
  //生徒一覧の修正があった場合、フォームを変更する。
  var itemstudent = items[2];
  var studentMCItem = itemstudent.asMultipleChoiceItem();
  var revisedFormstudentList =[];
  //生徒一覧を格納（getValues()で得られるのは、2次元配列のため1次元配列に入れなおす）
  for(var j=1;j<sheetContents.getLastColumn();j++){
    revisedFormstudentList.push(studentList[0][j]);
  }
  //新たに選択肢を設定する
  studentMCItem.setChoiceValues(formstudentList);
  studentMCItem.showOtherOption(true);
  
  //講義一覧の修正があった場合、フォームを変更する。
  var itemContents = items[3];
  var contentsListItem = itemContents.asListItem();
  var revisedContentsList = [];
  //生徒一覧を格納（getValues()で得られるのは、2次元配列のため1次元配列に入れなおす）
  for(var k=4;k<sheetContents.getLastRow()-4;k++){
    revisedContentsList.push(contentsList[k]);
  }  
  //新たに選択肢を設定する
  contentsListItem.setChoiceValues(revisedContentsList);
}
