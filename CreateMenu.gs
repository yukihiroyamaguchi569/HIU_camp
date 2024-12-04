function onOpen() {
  // スプレッドシートを開いた時に実行される関数
  const ui = SpreadsheetApp.getUi();
  
  // メニューを作成
  ui.createMenu('HIU合宿')
    .addItem('名札を作成', 'createNameTags')
    .addToUi();
}