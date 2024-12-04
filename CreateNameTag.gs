function createNameTags() {
  // スプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("名札用");
  
  // データ範囲を取得
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // スプレッドシートを取得
  const presentation = SlidesApp.openById("1TY54vCikZLtv2hXiJL0koxZ8brE5mpetZMWtQEqEfKA");
  
  // タイトルスライドを削除
  presentation.getSlides()[0].remove();
  
  // 各参加者のスライドを作成
  for (let i = 1; i < values.length; i++) {
    const [name,,,,,,,gpt,yado1,keidro,yado2, enma] = values[i];
    
    // 新しいスライドを追加
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // 名前を追加
    const nameShape = slide.insertTextBox(name);
    const nameText = nameShape.getText();
    nameText.getTextStyle().setFontSize(35).setBold(true);
    nameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    nameShape.setTop(40);
    nameShape.setLeft(30);
    nameShape.setWidth(200);
    
    // gptを追加
    const gptShape = slide.insertTextBox(gpt);
    const gptText = gptShape.getText();
    gptText.getTextStyle().setFontSize(24);
    // infoText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    gptShape.setTop(100);
    gptShape.setLeft(50);
    gptShape.setWidth(40);

        // yado1を追加
    const yado1Shape = slide.insertTextBox(yado1);
    const yado1Text = yado1Shape.getText();
    yado1Text.getTextStyle().setFontSize(24);
    // infoText.getParagraphStyle().setAlignment(SlidesApp.ParagraphAlignment.CENTER);
    yado1Shape.setTop(100);
    yado1Shape.setLeft(80);
    yado1Shape.setWidth(40);

        // keidroを追加
    const keidroShape = slide.insertTextBox(keidro);
    const keidroText = keidroShape.getText();
    keidroText.getTextStyle().setFontSize(24);
    // infoText.getParagraphStyle().setAlignment(SlidesApp.ParagraphAlignment.CENTER);
    keidroShape.setTop(100);
    keidroShape.setLeft(110);
    keidroShape.setWidth(40);

        // yado2を追加
    const yado2Shape = slide.insertTextBox(yado2);
    const yado2Text = yado2Shape.getText();
    yado2Text.getTextStyle().setFontSize(24);
    // infoText.getParagraphStyle().setAlignment(SlidesApp.ParagraphAlignment.CENTER);
    yado2Shape.setTop(100);
    yado2Shape.setLeft(140);
    yado2Shape.setWidth(40);

        // enmaを追加
    const enmaShape = slide.insertTextBox(enma);
    const enmaText = enmaShape.getText();
    enmaText.getTextStyle().setFontSize(24);
    // infoText.getParagraphStyle().setAlignment(SlidesApp.ParagraphAlignment.CENTER);
    enmaShape.setTop(100);
    enmaShape.setLeft(170);
    enmaShape.setWidth(40);
  }
  
}

