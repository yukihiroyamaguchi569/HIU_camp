function createNameTags() {
  // スプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("名札用");
  
  // データ範囲を取得
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  // スプレッドシートを取得
  const presentation = SlidesApp.openById("1nYrlLEk9AiI67VCe1PKIK14tT0LAtPbd3YQ65znCYZw");

   // 2枚目以降のスライドを削除
  const slides = presentation.getSlides();
  for (let i = slides.length - 1; i > 0; i--) {
    slides[i].remove();
  }
  
  // 各参加者のスライドを作成
  for (let i = 1; i < values.length; i++) {
    const [name,,,,,,,,gpt,yado1,keidro,yado2, enma] = values[i];

      // 値が空の場合は空文字列を設定
    const safeGpt = gpt || '';
    const safeYado1 = yado1 || '';
    const safeKeidro = keidro || '';
    const safeYado2 = yado2 || '';
    const safeEnma = enma || '';
    
    // 新しいスライドを追加
    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    
    // 名前を追加
    if (name) {
      const nameShape = slide.insertTextBox(name);
      const nameText = nameShape.getText();
      nameText.getTextStyle().setFontSize(35).setBold(true);
      nameText.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      nameShape.setTop(40);
      nameShape.setLeft(30);
      nameShape.setWidth(200);
    }

   
    // gptを追加
    if (safeGpt) {
      const gptShape = slide.insertTextBox(safeGpt);
      const gptText = gptShape.getText();
      gptText.getTextStyle().setFontSize(24);
      gptShape.setTop(100);
      gptShape.setLeft(30);
      gptShape.setWidth(50);
    }

    // yado1を追加
    if (safeYado1) {
      const yado1Shape = slide.insertTextBox(safeYado1);
      const yado1Text = yado1Shape.getText();
      yado1Text.getTextStyle().setFontSize(24);
      yado1Shape.setTop(100);
      yado1Shape.setLeft(60);
      yado1Shape.setWidth(50);
    }

    // keidroを追加
    if (safeKeidro) {
      const keidroShape = slide.insertTextBox(safeKeidro);
      const keidroText = keidroShape.getText();
      keidroText.getTextStyle().setFontSize(24);
      keidroShape.setTop(100);
      keidroShape.setLeft(90);
      keidroShape.setWidth(50);
    }

    // yado2を追加
    if (safeYado2) {
      const yado2Shape = slide.insertTextBox(safeYado2);
      const yado2Text = yado2Shape.getText();
      yado2Text.getTextStyle().setFontSize(24);
      yado2Shape.setTop(100);
      yado2Shape.setLeft(120);
      yado2Shape.setWidth(40);
    }

    // enmaを追加
    if (safeEnma) {
      const enmaShape = slide.insertTextBox(safeEnma);
      const enmaText = enmaShape.getText();
      enmaText.getTextStyle().setFontSize(24);
      enmaShape.setTop(100);
      enmaShape.setLeft(150);
      enmaShape.setWidth(40);
    }
  }

    // 最後にタイトルスライド（1枚目）を削除
  presentation.getSlides()[0].remove();
}





