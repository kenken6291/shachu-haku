/**
 * 全国スーパー銭湯一覧 → Spots(category: onsen) 登録スクリプト
 * -----------------------------------------------------------
 * 事前準備:
 *   1. SuperSento_Raw.csv を、このスプレッドシートに「SuperSento_Raw」という
 *      名前のシートとしてインポートしておく（ヘッダー行あり: 都道府県, 住所, 名称, 休業日, 平日大人料金, 詳細ページURL）
 *   2. このファイルを、Spots管理に使っている既存のApps Scriptプロジェクトに
 *      新しい.gsファイルとして追加する
 *
 * 使い方:
 *   - importSuperSentoToSpots() を実行 → 1回の実行で最大 BATCH_SIZE 件を処理
 *   - 2566件あるため1回では終わらない。続きから自動再開するので、
 *     完了するまで手動で数回実行するか、installAutoTrigger() で自動化する
 *   - 進捗は ScriptProperties に保存されるため、実行を中断・再開しても
 *     重複登録されない
 *   - ジオコーディングに失敗した住所は「GeocodeErrors」シートに記録される
 */

const RAW_SHEET_NAME = 'SuperSento_Raw';
const SPOTS_SHEET_NAME = 'Spots';
const ERROR_SHEET_NAME = 'GeocodeErrors';
const PROGRESS_KEY = 'superSentoImportProgress';

// 1回の実行で処理する件数。GASの実行時間上限(6分)とジオコーディングの
// レート制限を考慮して調整可能。
const BATCH_SIZE = 80;

// 既存Spotsデータの登録者に合わせる（他の手動登録スポットと同じ値）
const AUTHOR_ID = 'u-4dfa6aa4';
const AUTHOR_NICKNAME = 'ケン';

function importSuperSentoToSpots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);
  if (!rawSheet) {
    throw new Error(`シート「${RAW_SHEET_NAME}」が見つかりません。先にCSVをインポートしてください。`);
  }
  const spotsSheet = ss.getSheetByName(SPOTS_SHEET_NAME);
  if (!spotsSheet) {
    throw new Error(`シート「${SPOTS_SHEET_NAME}」が見つかりません。`);
  }
  let errorSheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!errorSheet) {
    errorSheet = ss.insertSheet(ERROR_SHEET_NAME);
    errorSheet.appendRow(['都道府県', '住所', '名称', 'エラー内容']);
  }

  const data = rawSheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1); // [都道府県, 住所, 名称, 休業日, 平日大人料金, 詳細ページURL]

  const props = PropertiesService.getScriptProperties();
  let startIndex = Number(props.getProperty(PROGRESS_KEY) || '0');

  if (startIndex >= rows.length) {
    Logger.log('すべての行が処理済みです。');
    return { done: true, processed: 0, total: rows.length };
  }

  const geocoder = Maps.newGeocoder().setRegion('jp');
  const endIndex = Math.min(startIndex + BATCH_SIZE, rows.length);
  const newSpotRows = [];

  for (let i = startIndex; i < endIndex; i++) {
    const [pref, addr, name, closed, fee, url] = rows[i];
    if (!name || !addr) continue;

    const fullAddress = `${pref || ''}${addr || ''}`;
    let lat = '';
    let lng = '';

    try {
      const result = geocoder.geocode(fullAddress);
      if (result.status === 'OK' && result.results.length > 0) {
        lat = result.results[0].geometry.location.lat;
        lng = result.results[0].geometry.location.lng;
      } else {
        errorSheet.appendRow([pref, addr, name, `geocode status: ${result.status}`]);
      }
    } catch (e) {
      errorSheet.appendRow([pref, addr, name, `例外: ${e.message}`]);
    }

    const now = new Date().toISOString();
    const feeText = (typeof fee === 'number') ? `${fee.toLocaleString('ja-JP')}円` : (fee || '');
    const description = [addr, url].filter(Boolean).join(' / ');

    newSpotRows.push([
      'm-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), // id
      name,                                                      // name
      'onsen',                                                   // category
      lat,                                                        // lat
      lng,                                                        // lng
      description,                                                // description
      '{"bath":true}',                                            // facilities（デフォルト値、要確認）
      closed || '',                                               // openHours ※休業日の値を格納（指示通り）
      feeText,                                                    // fee
      AUTHOR_ID,                                                  // authorId
      AUTHOR_NICKNAME,                                            // authorNickname
      now,                                                        // createdAt
      now,                                                        // updatedAt
      'unclear',                                                  // sleepStatus
      '[]',                                                       // photos
      '',                                                          // tel
    ]);

    // ジオコーディングAPIのレート制限対策
    Utilities.sleep(250);
  }

  if (newSpotRows.length > 0) {
    spotsSheet.getRange(spotsSheet.getLastRow() + 1, 1, newSpotRows.length, newSpotRows[0].length)
      .setValues(newSpotRows);
  }

  props.setProperty(PROGRESS_KEY, String(endIndex));

  const done = endIndex >= rows.length;
  Logger.log(`処理済み: ${endIndex}/${rows.length}件${done ? '（完了）' : ''}`);

  if (done) {
    removeAutoTrigger();
  }

  return { done, processed: endIndex - startIndex, total: rows.length, nextIndex: endIndex };
}

/** 進捗をリセットして最初からやり直したいとき用 */
function resetSuperSentoProgress() {
  PropertiesService.getScriptProperties().deleteProperty(PROGRESS_KEY);
  Logger.log('進捗をリセットしました。');
}

/** 完了まで自動的に数分おきに続きを実行するトリガーを設置する */
function installAutoTrigger() {
  removeAutoTrigger();
  ScriptApp.newTrigger('importSuperSentoToSpots')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('自動実行トリガーを設置しました（5分おき）。完了すると自動的に削除されます。');
}

function removeAutoTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'importSuperSentoToSpots') {
      ScriptApp.deleteTrigger(t);
    }
  });
}
