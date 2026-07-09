function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  
  // 1. 誤作動防止：データ入力シート（シート1）のA列・5行目以降のみ対象
  if (sheet.getName() === "設定" || range.getColumn() !== 1 || range.getRow() < 5) {
    return;
  }
  
  const rawText = range.getValue().toString();
  if (rawText.trim() === "") {
    return;
  }
  
  // 2. 「設定」シートから想定項目数を動的に読み込む
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("設定");
  
  // B1セルから設定値（今回は4）を取得
  const EXPECTED_COUNT = Number(configSheet.getRange("B1").getValue());
  
  // 3. データの揺れ対策
  const normalizedText = rawText.replace(/\r\n|\r/g, "\n").trim();
  const dataLines = normalizedText.split("\n");
  
  const editedRow = range.getRow();
  
  // 4. 現在データがある右端（lastCol）を調べて、そこまで全消しする
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 2) {
    const clearColCount = lastCol - 1; 
    sheet.getRange(editedRow, 2, 1, clearColCount).clearContent();
  }
  
  // 5. バリデーションチェック
  if (dataLines.length !== EXPECTED_COUNT) {
    const errorRange = sheet.getRange(editedRow, 2);
    errorRange.setValue(`【エラー】項目数が違います（現在: ${dataLines.length} / 想定: ${EXPECTED_COUNT}）`);
    errorRange.setBackground("#ff0000"); // ★エラー時は薄い赤にする！
    return;
  }
  
  // 6. データ整形
  const outputValues = [];
  for (let i = 0; i < dataLines.length; i++) {
    outputValues.push(dataLines[i].trim());
  }
  
  // 7. 書き込み（正常に書き込めたら背景を白に戻す）
  const successRange = sheet.getRange(editedRow, 2, 1, EXPECTED_COUNT);
  successRange.setBackground("#ffffff"); // ★背景を白（通常）にリセット！
  successRange.setValues([outputValues]);
}
