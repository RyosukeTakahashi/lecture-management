//開いたら実行される。
function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menus = [
               {name: 'メール配信用サイドバー表示', functionName: 'showSidebar'}
              ];
  ss.addMenu('追加機能一覧', menus);
}
 //昨日実施された
//予定を日付、時刻でソートするメソッド
function sortSchedule(){
  var ssSchedule = SpreadsheetApp.getActive(); //アクティブなスプレッドシート全体（ss）を読み込む
  var sheetSchedule = ssSchedule.getSheets()[0];//[0]番目のシート単体を読む
  var queryCell = sheetSchedule.getRange("K1")//Queryを書くセル
  var lastRow = sheetSchedule.getLastRow();//入力があるセルの内の最後の行の行数を取得
  講義数のカウント
  //Formから送信されたデータが記入されるRangeがあるが、そのRangeはソートなどが仕様上できない。
  //そのため、一度データを隣接したRangeに移動し、Queryを使ってソートする。
  //そしてソートされた内容を、Formから送信されたデータが記入されるRangeに上書きする。

  //ソートする範囲をlastRowに従って変えて、日付（B列）と時刻（C列）でソートする。
  Utilities.sleep(100)
  queryCell.setFormula("=QUERY($B$1:$J$" + lastRow + ", \"Order By B,C\")");
  Utilities.sleep(3000)
  
  //ソートされた隣接しているRangeを、Formから送信されたデータが記入されるRangeに上書きする。
  var queryRange = sheetSchedule.getRange(2,11,lastRow-1,9);
  var formRange = sheetSchedule.getRange(2,2,lastRow-1,9);
  formRange.setValues(queryRange.getValues());
  Utilities.sleep(1000)
  formFormat();//日付のフォーマット修正
}

//日付のフォーマットの修正
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


//一つの予定データから、配信するメール用文章を形成する。他のメソッドこれを全ての予定データに施し、結合する。
//@param i:行数
//@return 一つの予定の文字列
function createSingleOutput(i){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var values = sheetSchedule.getRange(i,3,1,8).getValues();//Formから受信するデータが記入されるRangeの値を取得
  
  //時間に記入がなければ"?"を記入
  for(j=0;j<=5;j++){
    if(values[0][j]==""){
        values[0][j]="?";
      }
    }    
  //可読性のため、取得した値に名前をつける
  var time =values[0][0];
  var student =values[0][1];
  var by_with = "by" //内容に応じて by or with にする。
  var contents =values[0][2];
  var lecturer =values[0][3];
  var preparation = " 準備:"+values[0][4];
  var place =values[0][5];
  
  //時間が"1230"のような形式でユーザーに記入させており ":"がないため、それを挿入する
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

  //それぞれの文字列を結合する。
  //以下のようになる。（例）

  // ●17:00~ こゆみ 『面談』
  //    with高橋@オフィス
  //    備考：高橋初対面
  var output = "●"+time+"~ "+student+" 『"+contents+"』\n"+"    "+by_with+lecturer+"@"+place+preparation+remark+"\n\n";

  return output;
}


// 上記メソッド、CreateSingleOutput(i)をループさせ、配信用メールの全文を作る。
//@return shedule メールで送信するスケジュールの文字列
function generateSchedule(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var TimeFormat =[sheetSchedule.getRange(1,3).getNumberFormat()]
  var TimeFormats =[]
  var lastRow = sheetSchedule.getLastRow();
  var queryCell = sheetSchedule.getRange("K1")
  var queryRange = sheetSchedule.getRange(2,11,lastRow-1,9);
  var formRange = sheetSchedule.getRange(2,2,lastRow-1,9);
  
  //ソートする範囲をlastRowに従って変えて、日付（B列）と時刻（C列）でソートする。
  Utilities.sleep(100)
  queryCell.setFormula("=QUERY($B$1:$J$" + lastRow + ", \"Order By B,C\")");
  Utilities.sleep(3000)
  
  formRange.setValues(queryRange.getValues());
  
  var i=2;
  var j=2;
  var schedule = "【昨日】\n";
  //背景色が赤→昨日の予定
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#ff0000"){
    //予定を整形した文字列を加えていく
    schedule+=createSingleOutput(i);
    i++;
  }
  //背景色が黄色→本日の予定
  schedule+="変更/報告ある場合は、シートをを編集してください。\nhttps://goo.gl/dnNxYh\n\n【今日】\n";
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#ffff00"){
    schedule+=createSingleOutput(i);
    i++;
  }
  schedule+="\n【明日】\n";
  //背景色が黄色→明日の予定
  while(sheetSchedule.getRange(i,2).getBackgroundColor()=="#00ff00"){
    schedule+=createSingleOutput(i);
    i++;
  }
  schedule+="\n以上\n";
  
  return schedule;  
}



//メール送信用サイドバーを表示。
function showSidebar() {
  colorizeBackGround();

  var userProp = PropertiesService.getUserProperties();
  
  var d = new Date();
  var date = Utilities.formatDate( d, 'JST', 'MM/dd');
  var mailBody = generateSchedule();
  
  var prevMailForm = {
    subject : date + "前後の予定" || "",
    body : mailBody ||""  };
  
  //sidebarをテンプレートして扱う
  var sidebarHtml = HtmlService.createTemplateFromFile("sidebar");
  
  sidebarHtml.prevMailForm = prevMailForm;
  //SpreadsheetApp上で表示
  SpreadsheetApp.getUi().showSidebar(sidebarHtml.evaluate().setTitle("メール配信"));
  
}

//メール送信する。(サイドバーのボタン呼び出される
//@param {object} mailForm メールの内容 subject:タイトル, body:本文
//@return {object} 送信結果

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


// メールの内容をチェックする。
// @param {object} mailForm メールの内容 subject:タイトル, body:本文
//注：関数名の最後にアンダーバーをつけることで、単体で実行するFunction一覧から除外させる
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
  
  //昨日以前の列を削除する（日付の背景色が薄緑のもの）
  while(sheetSchedule.getRange(2,2).getBackground() == "#b7e1cd"){
    sheetSchedule.deleteRow(2);
  }
  
  //灰色と白に背景色を交互に塗る
  for(var i=2;i<=50;i++){
    if(i%2==0){      
      sheetSchedule.getRange(i, 1, 1, 9).setBackgroundColor('#ebebeb');    
    }else{
      sheetSchedule.getRange(i, 1, 1, 9).setBackgroundColor('#FFFFFF');    
    }
  }
}


//昨日実施されたContentsを記録する。
function recordBs(){
  var ssSchedule = SpreadsheetApp.getActive();
  var sheetSchedule = ssSchedule.getSheets()[0];
  var sheetReports = ssSchedule.getSheets()[1];
  
  colorizeBackGround();
  
  //講義科目のリストを取得
  var contentsList = sheetReports.getRange("A:A").getValues();
  //科目名の行数
  var contentsRow;
  //生徒の名前のリストを取得
  var studentList = sheetReports.getRange("1:1").getValues();
  //生徒名の列数
  var nameCol;
  //今日の日付
  var d = new Date();
  //MM/ddの形に日付をフォーマットする
  var date = Utilities.formatDate( d, 'JST', 'MM/dd');
  //2行目から科目リストが始まるのでi=2
  var i = 2;
  //昨日実施された講義数のカウント
  var yesterdayCount = 0;
  
  //赤のセル（昨日実施分）をカウント
  while(sheetSchedule.getRange(i,2).getBackground()=="#ff0000"){
    yesterdayCount++;
    i++;
  }
  
  //yesterdayCountをもとに、昨日実施された予定データの範囲を取得
  if(yesterdayCount!=0){
    var rangeYesterday = sheetSchedule.getRange(2, 1, yesterdayCount, 9);
    for(var i=2;i<=rangeYesterday.getLastRow();i++){
    
      //ノートに記入する事項
      var lec = sheetSchedule.getRange(i,6).getValue();//講師
      var preparation = sheetSchedule.getRange(i,7).getValue();//準備
      var report = sheetSchedule.getRange(i,10).getValue();//フィードバックや報告
      //セルのノートに追加する文字列
      var noteToAdd = date+"\n講師:"+lec+"\n準備:"+preparation+"\n報告:\n"+report+"\n----------\n";
      
      //記入する列を取得するために、生徒の名前を生徒リストの何番目かを探す
      var indexOfNC = studentList[0].indexOf(sheetSchedule.getRange(i, 4).getValue())
      //新しい名前の場合、列を足す
      if(indexOfNC != -1){
        nameCol = indexOfNC + 1;//列数の関係で1を足す
      }else{
        sheetReports.insertColumnAfter(sheetReports.getLastColumn());
        nameCol = sheetReports.getLastColumn()+1;
        sheetReports.getRange(1, nameCol).setValue(sheetSchedule.getRange(i, 4).getValue())
      }
    
      //記入する行を取得し、ノート記入。
      for(var contentsRow=4;contentsRow<=sheetReports.getLastRow();contentsRow++){
        if(sheetSchedule.getRange(i,5).isBlank()){//内容未記入ならBreak
          break;
        }else if(sheetSchedule.getRange(i,5).getValue()==contentsList[contentsRow]){
          target = sheetReports.getRange(contentsRow+1,nameCol);
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
