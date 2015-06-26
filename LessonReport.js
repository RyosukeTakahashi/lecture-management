//gitHub確認
//開いたら実行される。
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [
               {name: 'メール配信用サイドバー表示', functionName: 'showSidebar'}
              ];
  ss.addMenu('追加機能一覧', menus);
}
 
//予定を日付、時刻でソートするメソッド
function formToSchedule(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var queryCell = sheetSchedule.getRange("K1")
  var lastRow = sheetSchedule.getLastRow();
  
  //Formから送信されたデータが記入されるRangeがあるが、そのRangeはソートなどが仕様上できない。
  //そのため、一度データを隣接したRangeに移動し、Queryを使ってソートする。
  //そしてソートされた内容を、Formから送信されたデータが記入されるRangeに上書きする。

  //QueryRangeの範囲をformRangeに合わせ、日付（B列）と時刻（C列）でソートする
  Utilities.sleep(100)
  queryCell.setFormula("=QUERY($B$1:$J$" + lastRow + ", \"Order By B,C\")");
  Utilities.sleep(3000)
  
  //ソートされた隣接しているRangeを、Formから送信されたデータが記入されるRangeに上書きする。
  var queryRange = sheetSchedule.getRange(2,11,lastRow-1,9);
  var formRange = sheetSchedule.getRange(2,2,lastRow-1,9);
  formRange.setValues(queryRange.getValues());
  Utilities.sleep(1000)
  formFormat();
}

//日付とのフォーマットの修正
function formFormat(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetResponse =ssSchedule.getSheets()[0];
  var dateFormats = []

  for(var k=1;k<=50;k++){
    dateFormats.push(["mm\"/\"dd"])
  }
  sheetResponse.getRange(1,2,50,1).setNumberFormats(dateFormats)
  Utilities.sleep(2000)
}


//メール用出力の形成
//@param i:行数
function createOutput(i){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var values = sheetSchedule.getRange(i,3,1,8).getValues();//Formから受信するデータ記入されるRange
  
  //時間に記入がなければ"?"を記入
  for(j=0;j<=5;j++){
    if(values[0][j]==""){
        values[0][j]="?";
      }
    }    
  var time =values[0][0];
  var student =values[0][1];
  var by_with = "by"
  var contents =values[0][2];
  var lecturer =values[0][3];
  var place =values[0][4];
  var preparation = " 準備:"+values[0][5];
  
  //時間が"1230"のような形式で ":"がないため、それを挿入する
  if(/:/.test(time)==false&&time!="?"){
    if(typeof time == "number"){
      time = String(time)
    }
    //時刻が3桁で記入されていた場合、4桁表示にする。
    if(time.length==3){
      time = ("0000"+time).slice( -4 ); 
    }
    hour=time.substr(0,2)
    minute=time.substr(2,2)
    time = hour +':'+minute
  }
  
  //何かしら授業以外の行事の場合、いくつかの項目を空白にする
  if(/サッカ/.test(contents)||/バレー/.test(contents)||/歌/.test(contents)){
    student = ""
    lecturer = ""
    preparation =""
    by_with = ""
  }
  if(/対話/.test(contents)||/面談/.test(contents)){
    by_with = "with"
    preparation =""
  }
  
  //備考が空白でなければ備考を記入する
  if(values[0][6]!==""){
    var remark ="\n    備考："+values[0][6];
  }else{
    var remark ="";
  }
  var output = "●"+time+"~ "+student+" 『"+contents+"』\n"+"    "+by_with+lecturer+"@"+place+preparation+remark+"\n\n";
  Logger.log(values);
  Logger.log(output)

  return output;
}


//「●昨日」「●今日」 などをふくめてメール用にフォーマット整形
function reportSchedule(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var TimeFormat =[sheetSchedule.getRange(1,3).getNumberFormat()]
  var TimeFormats =[]
  var lastRow = sheetSchedule.getLastRow();
  var queryCell = sheetSchedule.getRange("K1")
  var queryRange = sheetSchedule.getRange(2,11,lastRow-1,9);
  var formRange = sheetSchedule.getRange(2,2,lastRow-1,9);
  
  //QueryRangeの範囲をformRangeに合わせ、日付（B列）と時刻（C列）でソートする
  Utilities.sleep(100)
  queryCell.setFormula("=QUERY($B$1:$J$" + lastRow + ", \"Order By B,C\")");
  Utilities.sleep(3000)
  
  formRange.setValues(queryRange.getValues());
  
  var i=2;
  var j=2;
  var report = "【昨日】\n";
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#ff0000"){
    report+=createOutput(i);
    i++;
  }
  report+="変更/報告ある場合は、シートをを編集してください。\nhttps://goo.gl/dnNxYh\n\n【今日】\n";
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#ffff00"){
    report+=createOutput(i);
    i++;
  }
  report+="\n【明日】\n";
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#00ff00"){
    report+=createOutput(i);
    i++;
  }
  report+="\n以上\n";
  
  return report;  
}



//メール送信用サイドバーを表示。
function showSidebar() {
  colorizeBackGround();

  var userProp = PropertiesService.getUserProperties();
  
  var d = new Date();
  var date = Utilities.formatDate( d, 'JST', 'MM/dd');
  var honbun = reportSchedule();
  
  var prevMailForm = {
    subject : date + "前後の予定" || "",
    body : honbun ||""  };
  
  //sidebarをテンプレートして扱う
  var sidebarHtml = HtmlService.createTemplateFromFile("sidebar");
  
  sidebarHtml.prevMailForm = prevMailForm;
  //SpreadsheetApp上で表示
  SpreadsheetApp.getUi().showSidebar(sidebarHtml.evaluate().setTitle("メール配信"));
  
}

/**
 * メール送信します。(サイドバーから呼び出されます。)
 * @param {object} mailForm メールの内容 subject:タイトル, body:本文
 * @return {object} 送信結果
 */
function sendMail(mailForm) {

  //件名、本文が空欄でないかのチェック。
  validateMailForm_(mailForm);
  
  //メール配信
  MailApp.sendEmail("ramuniku@gmail.com", mailForm.subject, mailForm.body);
  
  //連続で送るとエラーになるので少し待たせます。
  Utilities.sleep(100);    
  
  //完了した旨をSpreadsheetに表示します。
  SpreadsheetApp.getUi().alert("メール送信が完了しました");
  return {message: "メール送信が完了しました"};
}

/**
 * メールの内容をチェックします。
 * @param {object} mailForm メールの内容 subject:タイトル, body:本文
 */
//関数名の最後に_(アンダースコア)が付いているとその関数はプライベート関数になります。
//プライベート関数は上部の関数呼び出しSelectBoxに表示されない、google.script.runで呼び出せないなどの特徴があります。
function validateMailForm_(mailForm) {

  if (mailForm.subject == ""){
    throw new Error("メールタイトルは必須です。");
  }
  
  if (mailForm.body == ""){
    throw new Error("メール本文は必須です。");
  }
}

//背景の色付けと過去の記録の削除
function colorizeBackGround() {
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  
  //昨日以前の列を削除する
  while(sheetSchedule.getRange(2,2).getBackground() == "#b7e1cd"){
    sheetSchedule.deleteRow(2);
  }
  
  
  //灰色と白に背景色を塗る
  for(var i=2;i<=50;i++){
    if(i%2==0){      
      sheetSchedule.getRange(i, 1, 1, 9).setBackgroundColor('#ebebeb');    
    }else{
      sheetSchedule.getRange(i, 1, 1, 9).setBackgroundColor('#FFFFFF');    
    }
  }
}


//実施されたContentsを記録する。
function recordBs(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var sheetContents = ssSchedule.getSheets()[1];
  
  colorizeBackGround();
  
  var contentsList = sheetContents.getRange("A:A").getValues();
  Logger.log(contentsList)
  var contentsRow;
  var studentList = sheetContents.getRange("1:1").getValues();
  var nameCol;
  var d = new Date();
  var date = Utilities.formatDate( d, 'JST', 'MM/dd');
  var i = 2;
  var yesterdayCount = 0;
  
  //赤のセル（昨日実施分）をカウント
  while(sheetSchedule.getRange(i,2).getBackground()=="#ff0000"){
    yesterdayCount++;
    i++;
  }
  
  if(yesterdayCount!=0){
    var rangeYesterday = sheetSchedule.getRange(2, 1, yesterdayCount, 9);
    Logger.log(rangeYesterday.getLastRow())
    for(var i=2;i<=rangeYesterday.getLastRow();i++){
    
      //ノートに記入する事項
      var lec = sheetSchedule.getRange(i,6).getValue();
      var preparation = sheetSchedule.getRange(i,7).getValue();
      var report = sheetSchedule.getRange(i,10).getValue();
      var noteToAdd = date+"\n講師:"+lec+"\n準備:"+preparation+"\n報告:\n"+report+"\n----------\n";
      
      //記入する列を取得、または新規NCの場合、列を足す。
      var indexOfNC = studentList[0].indexOf(sheetSchedule.getRange(i, 4).getValue())
      if(indexOfNC != -1){
        nameCol = indexOfNC + 1;
      }else{
        sheetContents.insertColumnAfter(sheetContents.getLastColumn());//新しい名前の場合、行を足す
        nameCol = sheetContents.getLastColumn()+1;
        sheetContents.getRange(1, nameCol).setValue(sheetSchedule.getRange(i, 4).getValue())
      }
    
      //記入する行を取得し、ノート記入。
      for(var contentsRow=4;contentsRow<=sheetContents.getLastRow();contentsRow++){
        if(sheetSchedule.getRange(i,5).isBlank()){//内容未記入ならBreak
          break;
        }else if(sheetSchedule.getRange(i,5).getValue()==contentsList[contentsRow]){
          target = sheetContents.getRange(contentsRow+1,nameCol);
          target.setValue(date);
          var note = target.getNote();
          note += noteToAdd;
          target.setNote(note);
          break;
        }
      }
    }
  }
}
