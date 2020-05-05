/* 指定月のカレンダーからイベントを取得する */
function getCalendar() {

  // フォルダ、ファイル関係の定義
  var targetFolderIds = [" #フォルダのIDを入力 "];
  var targetFolder;
  var folderName;
  var objFiles;
  var objFile;
  var fileName;

  // スプレッドシート関係の定義
  var ss;
  var key;
  var sheets;
  var sheetId;


  for (var i = 0; i < targetFolderIds.length; i++) {
    // Idから対象フォルダの取得
    targetFolder = DriveApp.getFolderById(targetFolderIds[i]);
    folderName = targetFolder.getName();

    // 対象フォルダ以下のSpreadsheetを取得
    objFiles = targetFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (objFiles.hasNext()) {
      objFile = objFiles.next();
      fileName = objFile.getName();

      // Spreadsheetのオープン
      ss = SpreadsheetApp.openByUrl(objFile.getUrl());
      key = ss.getId();
      sheets = ss.getSheets();
    }
  }
  
  var now=new Date;
  
  //特定の日時で指定する場合はこのように記載
  //var startDate=new Date('2020/05/04 00:00:00');
  
  var startDate=new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1); //取得開始日
  var endDate=now; 　//取得終了日
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(" #シート名1 ");
  
  var no=1; //No

  var myCal=CalendarApp.getCalendarById(' #カレンダーID(メールアドレス)を入力 '); //特定のIDのカレンダーを取得
  var myEvents=myCal.getEvents(startDate,endDate); //カレンダーのイベントを取得
  
  /* イベントの数だけ繰り返してシートに記録 */
  myEvents.forEach(function(evt){
    sheet.appendRow(
      [
        no, //No
        evt.getTitle(), //イベントタイトル
        evt,
        evt.getStartTime(), //イベントの開始時刻
        evt.getEndTime(), //イベントの終了時刻
        "=INDIRECT(\"RC[-1]\",FALSE)-INDIRECT(\"RC[-2]\",FALSE)" //所要時間を計算
      ]
    );
    no++;
  });

  /* 列分解 */
　var lastRow = sheet.getLastRow();
  for(i=2;i <= lastRow;i++){
    var x = sheet.getRange(i,2);
    var y = sheet.getRange(i,11);  //　使っていないセルを取得
    var z = sheet.getRange(2,3,lastRow-1,1);

    z.clearContent();
    y.setValue(x.getValue());
    x.clearContent();  
    strformula = "=split(K" + i + ",\"/\")";
    x.setFormula(strformula);
  }  

  /* 所要時間の[ss]表示　*/
  var secondTime = '[ss]';
  for(i=2;i <= lastRow;i++){
　　　　　  var numberRange=sheet.getRange(i,6,lastRow);
 　　　　　 numberRange.setNumberFormat( secondTime );
  }

  /* 週番号の追加 */
  var x=sheet.getRange(1,7);
  var y=sheet.getRange(1,8);
  var z=sheet.getRange(1,9);
  var w=sheet.getRange(1,10);
  x.setValue('週番号');
  y.setValue('日にち');
  z.setValue('所要時間(hh:mm)');
  w.setValue('曜日');
  var NumFormats = '0'; 
  for(i=2;i <= lastRow;i++){
    var x=sheet.getRange(i,7);
    var y=sheet.getRange(i,8);
    var z=sheet.getRange(i,9);
    var w=sheet.getRange(i,10);
    weeknum = "=WEEKNUM(D" + i + ")";
    x.setFormula(weeknum);
    x.setNumberFormat( NumFormats );
    daynum = "=day(D" + i + ")";
    y.setFormula(daynum);
    hhmm = "=TEXT(F" + i + ",\"hh:mm\")";
    z.setFormula(hhmm);
    week = "=TEXT(D" + i + ",\"ddd\")";
    w.setFormula(week); 
  }
  
  copyDataToNewSheet();
}

/* シート sheet_raw に、値のペースト　*/
function copyDataToNewSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(" #シート名1 ");
  var sheet2 = spreadsheet.getSheetByName(" #シート名2 ");
  sheet.getRange("A1:K400000").copyValuesToRange(sheet2,1,11,1,400000);
}

/* シートを初期化する関数 */
function initialize() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(" #シート名1 ");
  var range = sheet.getRange(1,11,1,400000);
  //range.clear();
  
  //列名を入力
  var range = sheet.getRange("A1").setValue("No");
  var range = sheet.getRange("B1").setValue("カテゴリ");
  var range = sheet.getRange("C1").setValue("内容");
  var range = sheet.getRange("D1").setValue("開始時刻");
  var range = sheet.getRange("E1").setValue("終了時刻");
  var range = sheet.getRange("F1").setValue("所要時間");
  
  getCalendar();
}
