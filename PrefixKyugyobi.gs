/**
 * 「SuperSento_Raw」シートの「休業日」列（E列）の各セルの先頭に
 * 「休業日」という文字を追加する。
 * ヘッダー行（1行目）は対象外。
 */
function addKyugyobiPrefix() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SuperSento_Raw');
  if (!sheet) {
    throw new Error('シート「SuperSento_Raw」が見つかりません。');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データ行がありません。');
    return;
  }

  // ヘッダーから「休業日」列の位置を特定
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = header.indexOf('休業日') + 1; // 1始まりの列番号
  if (col === 0) {
    throw new Error('ヘッダーに「休業日」列が見つかりません。');
  }

  const range = sheet.getRange(2, col, lastRow - 1, 1);
  const values = range.getValues();

  const newValues = values.map(([v]) => {
    if (v === '' || v === null) return [v];
    const text = String(v);
    // 既に「休業日」が付いている場合は二重に付けない
    if (text.startsWith('休業日')) return [text];
    return ['休業日' + text];
  });

  range.setValues(newValues);
  Logger.log(`${newValues.length}件のセルを更新しました。`);
}
