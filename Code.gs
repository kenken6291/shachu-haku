/**
 * くるま旅マップ - 会員システム バックエンド (Google Apps Script)
 * ------------------------------------------------------------
 * シート構成:
 *   Users     : id, email, passwordHash, salt, nickname, mustChangePassword, agreedTerms, createdAt, updatedAt
 *   Spots     : id, name, category, lat, lng, description, facilities(JSON文字列), openHours, fee, authorId, authorNickname, createdAt, updatedAt, sleepStatus, photos(JSON文字列・DriveファイルID配列), tel
 *   Comments  : id, spotId, text, authorId, authorNickname, createdAt, updatedAt, rating(0〜5)
 *   Favorites : id, userId, spotId, createdAt
 *   Visits    : id, userId, spotId, visitedDate(YYYY-MM-DD), note, createdAt, updatedAt
 *
 * 初回セットアップ:
 *   1. このプロジェクトに紐づくスプレッドシートを用意
 *   2. スクリプトエディタで setupSheets() を一度実行（シートと見出し行を自動作成）
 *   3. [プロジェクトの設定] → [スクリプト プロパティ] に以下を追加
 *        PEPPER            : パスワードハッシュ用の適当な長いランダム文字列
 *        PHOTOS_FOLDER_ID  : 写真保存用Googleドライブフォルダの ID（フォルダURLの末尾の文字列）
 *        GEMINI_API_KEY    : 音声入力の文字起こし・整形に使うGemini APIキー
 *                            （モデルは画面で選択。使えない時は汎用の最新モデルへ自動切替。testGeminiModels() で確認可）
 *   4. [デプロイ] → [新しいデプロイ] → ウェブアプリ
 *        実行するユーザー: 自分
 *        アクセスできるユーザー: 全員
 *   5. 発行されたURLを index.html の CONFIG.gasUrl に設定
 *   6. 初回実行時にGmail送信・Googleドライブの権限承認が必要（ダイアログの指示に従う）
 *   7. 既にデプロイ済みのスプレッドシートを更新する場合は、
 *      migrateAddSleepStatusColumn() / migrateAddPhotosColumn() / migrateAddRatingColumn() / migrateAddTelColumn() /
 *      migrateAddFavoritesSheet() / migrateAddVisitsSheet() / migrateFixVisitDatesToText() を
 *      それぞれ一度だけ手動実行してから、デプロイを「編集」して新バージョンを作成してください。
 * ------------------------------------------------------------
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_USERS = 'Users';
const SHEET_SPOTS = 'Spots';
const SHEET_COMMENTS = 'Comments';
const SHEET_FAVORITES = 'Favorites';
const SHEET_VISITS = 'Visits';

const PROPS = PropertiesService.getScriptProperties();
const PEPPER = PROPS.getProperty('PEPPER') || 'CHANGE_ME_PEPPER_IN_SCRIPT_PROPERTIES';
const SESSION_TTL_SEC = 60 * 60 * 24 * 7; // セッション有効期限: 7日間
const MAX_PHOTOS_PER_SPOT = 3;
const MAX_PHOTO_BYTES = 4 * 1024 * 1024; // 4MB（クライアント側で圧縮済みの想定に対するサーバー側の安全弁）

/* ============================== エントリポイント ============================== */

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ ok: false, error: 'invalid_request' });
  }
  try {
    switch (body.action) {
      case 'register':        return respond(register(body));
      case 'login':            return respond(login(body));
      case 'me':                return respond(me(body));
      case 'changePassword':  return respond(changePassword(body));
      case 'updateNickname':  return respond(updateNickname(body));
      case 'forgotPassword':  return respond(forgotPassword(body));
      case 'listSpots':        return respond(listSpots());
      case 'submitSpot':      return respond(submitSpot(body));
      case 'updateSpot':      return respond(updateSpot(body));
      case 'deleteSpot':      return respond(deleteSpot(body));
      case 'listComments':    return respond(listComments(body));
      case 'addComment':      return respond(addComment(body));
      case 'updateComment':   return respond(updateComment(body));
      case 'deleteComment':   return respond(deleteComment(body));
      case 'listSpotStats':   return respond(listSpotStats());
      case 'toggleFavorite':  return respond(toggleFavorite(body));
      case 'listFavorites':   return respond(listFavorites(body));
      case 'listVisits':      return respond(listVisits(body));
      case 'addVisit':        return respond(addVisit(body));
      case 'updateVisit':     return respond(updateVisit(body));
      case 'deleteVisit':     return respond(deleteVisit(body));
      case 'uploadPhoto':     return respond(uploadPhoto(body));
      case 'deletePhoto':     return respond(deletePhoto(body));
      case 'transcribeVoice': return respond(transcribeVoice(body));
      case 'geocodeAddress':  return respond(geocodeAddress(body));
      default:                  return respond({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return respond({ ok: false, error: 'server_error', message: String(err) });
  }
}

function doGet(e) {
  // 読み取り専用の一覧取得（プリフライト不要でスマホからも軽く叩けるように）
  try {
    const action = e.parameter.action;
    if (action === 'listSpots') return respond(listSpots());
    if (action === 'listComments') return respond(listComments(e.parameter));
    if (action === 'listSpotStats') return respond(listSpotStats());
    return respond({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return respond({ ok: false, error: 'server_error', message: String(err) });
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================== ユーティリティ ============================== */

function sheet(name) {
  const sh = SS.getSheetByName(name);
  if (!sh) throw new Error('sheet_not_found:' + name);
  return sh;
}

function nowIso() { return new Date().toISOString(); }

function genId(prefix) { return prefix + '-' + Utilities.getUuid().split('-')[0]; }

function genTempPassword() {
  const chars = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; // 紛らわしい文字(0,O,1,I)を除外
  let out = '';
  for (let i = 0; i < 8; i++) out += chars[Math.floor(Math.random() * chars.length)];
  return out;
}

function genSalt() { return Utilities.getUuid(); }

function hashPassword(password, salt) {
  const raw = salt + password + PEPPER;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function findRowByValue(sh, colIndex, value) {
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][colIndex] === value) return { row: r + 1, data: data[r] };
  }
  return null;
}

function safeJson(str, fallback) {
  try { return JSON.parse(str); } catch (e) { return fallback; }
}

function sendMail(to, subject, body) {
  MailApp.sendEmail(to, subject, body);
}

/* ============================== セッション ============================== */

function createSession(userId) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('session_' + token, userId, SESSION_TTL_SEC);
  return token;
}

function getUserIdFromToken(token) {
  if (!token) return null;
  return CacheService.getScriptCache().get('session_' + token);
}

function requireAuth(token) {
  const userId = getUserIdFromToken(token);
  if (!userId) throw new Error('unauthorized');
  return userId;
}

/* ============================== 住所→緯度経度（ジオコーディング） ============================== */
// フロント側で国土地理院APIが見つけられなかった場合（施設名など）のフォールバック。
// API利用量を抑えるためログイン会員のみ利用可。
function geocodeAddress(body) {
  requireAuth(body.token);
  const address = String(body.address || '').trim().slice(0, 120);
  if (!address) return { ok: false, error: 'empty_address' };

  const cache = CacheService.getScriptCache();
  const key = 'geo_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, address, Utilities.Charset.UTF_8));
  const cached = cache.get(key);
  if (cached) return { ok: true, results: JSON.parse(cached) };

  const res = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(address);
  if (!res || res.status !== 'OK') return { ok: true, results: [] };
  const results = (res.results || []).slice(0, 8).map(r => ({
    lat: r.geometry.location.lat,
    lng: r.geometry.location.lng,
    label: String(r.formatted_address || '').replace(/^日本、\s*/, ''),
  }));
  cache.put(key, JSON.stringify(results), 60 * 60 * 6);
  return { ok: true, results };
}

/* ============================== Users ============================== */
// 列: id, email, passwordHash, salt, nickname, mustChangePassword, agreedTerms, createdAt, updatedAt
const U = { id: 0, email: 1, passwordHash: 2, salt: 3, nickname: 4, mustChangePassword: 5, agreedTerms: 6, createdAt: 7, updatedAt: 8 };

function register(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const nickname = String(body.nickname || '').trim();
  const agreedTerms = !!body.agreedTerms;

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, error: 'invalid_email' };
  if (!nickname) return { ok: false, error: 'invalid_nickname' };
  if (!agreedTerms) return { ok: false, error: 'terms_not_agreed' };

  const sh = sheet(SHEET_USERS);
  if (findRowByValue(sh, U.email, email)) return { ok: false, error: 'email_taken' };

  const tempPassword = genTempPassword();
  const salt = genSalt();
  const hash = hashPassword(tempPassword, salt);
  const id = genId('u');
  sh.appendRow([id, email, hash, salt, nickname, true, true, nowIso(), nowIso()]);

  sendMail(email, '【くるま旅マップ】仮パスワードのお知らせ',
    nickname + ' 様\n\n' +
    'くるま旅マップへの会員登録ありがとうございます。\n' +
    '下記の仮パスワードでログインし、初回ログイン後、必ず新しいパスワードに変更してください。\n\n' +
    'メールアドレス: ' + email + '\n' +
    '仮パスワード: ' + tempPassword + '\n\n' +
    '※このメールに心当たりがない場合は破棄してください。');

  return { ok: true };
}

function login(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.email, email);
  if (!found) return { ok: false, error: 'invalid_credentials' };
  const row = found.data;
  if (hashPassword(password, row[U.salt]) !== row[U.passwordHash]) return { ok: false, error: 'invalid_credentials' };

  const token = createSession(row[U.id]);
  return { ok: true, token, userId: row[U.id], nickname: row[U.nickname], mustChangePassword: !!row[U.mustChangePassword] };
}

function me(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.id, userId);
  if (!found) return { ok: false, error: 'not_found' };
  return { ok: true, userId, nickname: found.data[U.nickname], email: found.data[U.email], mustChangePassword: !!found.data[U.mustChangePassword] };
}

function changePassword(body) {
  const userId = requireAuth(body.token);
  const newPassword = String(body.newPassword || '');
  if (newPassword.length < 8) return { ok: false, error: 'password_too_short' };

  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.id, userId);
  if (!found) return { ok: false, error: 'not_found' };

  const salt = genSalt();
  const hash = hashPassword(newPassword, salt);
  sh.getRange(found.row, U.passwordHash + 1).setValue(hash);
  sh.getRange(found.row, U.salt + 1).setValue(salt);
  sh.getRange(found.row, U.mustChangePassword + 1).setValue(false);
  sh.getRange(found.row, U.updatedAt + 1).setValue(nowIso());
  return { ok: true };
}

function updateNickname(body) {
  const userId = requireAuth(body.token);
  const nickname = String(body.nickname || '').trim();
  if (!nickname) return { ok: false, error: 'invalid_nickname' };

  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.id, userId);
  if (!found) return { ok: false, error: 'not_found' };
  sh.getRange(found.row, U.nickname + 1).setValue(nickname);
  sh.getRange(found.row, U.updatedAt + 1).setValue(nowIso());

  updateAuthorNicknameEverywhere(userId, nickname);
  return { ok: true, nickname };
}

function updateAuthorNicknameEverywhere(userId, nickname) {
  [SHEET_SPOTS, SHEET_COMMENTS].forEach(name => {
    const sh = sheet(name);
    const data = sh.getDataRange().getValues();
    const header = data[0];
    const authorIdCol = header.indexOf('authorId');
    const authorNicknameCol = header.indexOf('authorNickname');
    if (authorIdCol === -1 || authorNicknameCol === -1) return;
    for (let r = 1; r < data.length; r++) {
      if (data[r][authorIdCol] === userId) sh.getRange(r + 1, authorNicknameCol + 1).setValue(nickname);
    }
  });
}

function forgotPassword(body) {
  const email = String(body.email || '').trim().toLowerCase();
  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.email, email);
  if (!found) return { ok: true }; // メールの登録有無を外部に漏らさない

  const tempPassword = genTempPassword();
  const salt = genSalt();
  const hash = hashPassword(tempPassword, salt);
  sh.getRange(found.row, U.passwordHash + 1).setValue(hash);
  sh.getRange(found.row, U.salt + 1).setValue(salt);
  sh.getRange(found.row, U.mustChangePassword + 1).setValue(true);
  sh.getRange(found.row, U.updatedAt + 1).setValue(nowIso());

  sendMail(email, '【くるま旅マップ】仮パスワード再発行のお知らせ',
    '仮パスワードを再発行しました。ログイン後、新しいパスワードに変更してください。\n\n' +
    '仮パスワード: ' + tempPassword + '\n\n' +
    '※心当たりがない場合はこのメールを破棄してください。');
  return { ok: true };
}

/* ============================== Spots（会員投稿） ============================== */
// 列: id, name, category, lat, lng, description, facilities, openHours, fee, authorId, authorNickname, createdAt, updatedAt, sleepStatus, photos, tel
// ※ sleepStatus / photos / tel は既存シートへの後付け列のため末尾に追加（既存データの列位置を壊さないため）
const S = { id: 0, name: 1, category: 2, lat: 3, lng: 4, description: 5, facilities: 6, openHours: 7, fee: 8, authorId: 9, authorNickname: 10, createdAt: 11, updatedAt: 12, sleepStatus: 13, photos: 14, tel: 15 };
const SLEEP_STATUS_VALUES = ['ok', 'ng', 'unclear'];
function normalizeSleepStatus(v) { return SLEEP_STATUS_VALUES.includes(v) ? v : 'unclear'; }

function listSpots() {
  const sh = sheet(SHEET_SPOTS);
  const data = sh.getDataRange().getValues();
  const spots = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[S.id]) continue;
    spots.push({
      id: row[S.id], name: row[S.name], category: row[S.category],
      lat: row[S.lat], lng: row[S.lng], description: row[S.description],
      facilities: safeJson(row[S.facilities], {}),
      openHours: row[S.openHours], fee: row[S.fee],
      authorId: row[S.authorId], authorNickname: row[S.authorNickname],
      createdAt: row[S.createdAt], updatedAt: row[S.updatedAt],
      sleepStatus: normalizeSleepStatus(row[S.sleepStatus]),
      photos: safeJson(row[S.photos], []),
      tel: row[S.tel] == null ? '' : String(row[S.tel]),
      source: 'member',
    });
  }
  return { ok: true, spots };
}

function submitSpot(body) {
  const userId = requireAuth(body.token);
  const nickname = getUserNickname(userId);
  const d = body.spot || {};
  if (!d.name || typeof d.lat !== 'number' || typeof d.lng !== 'number') return { ok: false, error: 'invalid_spot' };

  const sh = sheet(SHEET_SPOTS);
  const id = genId('m');
  sh.appendRow([
    id, d.name, d.category || 'parking', d.lat, d.lng, d.description || '',
    JSON.stringify(d.facilities || {}), d.openHours || '', d.fee || '',
    userId, nickname, nowIso(), nowIso(),
    normalizeSleepStatus(d.sleepStatus),
    '[]', // photos は投稿後に uploadPhoto で追加する
    d.tel || '',
  ]);
  return { ok: true, id };
}

function updateSpot(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_SPOTS);
  const found = findRowByValue(sh, S.id, body.spotId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[S.authorId] !== userId) return { ok: false, error: 'forbidden' };

  const d = body.spot || {};
  if (d.name !== undefined) sh.getRange(found.row, S.name + 1).setValue(d.name);
  if (d.category !== undefined) sh.getRange(found.row, S.category + 1).setValue(d.category);
  if (d.lat !== undefined) sh.getRange(found.row, S.lat + 1).setValue(d.lat);
  if (d.lng !== undefined) sh.getRange(found.row, S.lng + 1).setValue(d.lng);
  if (d.description !== undefined) sh.getRange(found.row, S.description + 1).setValue(d.description);
  if (d.facilities !== undefined) sh.getRange(found.row, S.facilities + 1).setValue(JSON.stringify(d.facilities));
  if (d.openHours !== undefined) sh.getRange(found.row, S.openHours + 1).setValue(d.openHours);
  if (d.fee !== undefined) sh.getRange(found.row, S.fee + 1).setValue(d.fee);
  if (d.sleepStatus !== undefined) sh.getRange(found.row, S.sleepStatus + 1).setValue(normalizeSleepStatus(d.sleepStatus));
  if (d.tel !== undefined) sh.getRange(found.row, S.tel + 1).setValue(d.tel);
  sh.getRange(found.row, S.updatedAt + 1).setValue(nowIso());
  return { ok: true };
}

function deleteSpot(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_SPOTS);
  const found = findRowByValue(sh, S.id, body.spotId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[S.authorId] !== userId) return { ok: false, error: 'forbidden' };
  sh.deleteRow(found.row);
  deleteCommentsBySpot(body.spotId);
  deleteVisitsBySpot(body.spotId);
  return { ok: true };
}

function getUserNickname(userId) {
  const sh = sheet(SHEET_USERS);
  const found = findRowByValue(sh, U.id, userId);
  return found ? found.data[U.nickname] : '';
}

/* ============================== Spots：写真（Googleドライブ保存） ============================== */

function uploadPhoto(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_SPOTS);
  const found = findRowByValue(sh, S.id, body.spotId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[S.authorId] !== userId) return { ok: false, error: 'forbidden' };

  const photos = safeJson(found.data[S.photos], []);
  if (photos.length >= MAX_PHOTOS_PER_SPOT) return { ok: false, error: 'photo_limit' };

  const folderId = PROPS.getProperty('PHOTOS_FOLDER_ID');
  if (!folderId) return { ok: false, error: 'photos_not_configured' };

  const dataBase64 = body.dataBase64;
  if (!dataBase64) return { ok: false, error: 'invalid_photo' };

  let bytes;
  try { bytes = Utilities.base64Decode(dataBase64); } catch (e) { return { ok: false, error: 'invalid_photo' }; }
  if (bytes.length > MAX_PHOTO_BYTES) return { ok: false, error: 'photo_too_large' };

  const mimeType = body.mimeType || 'image/jpeg';
  const filename = 'spot_' + body.spotId + '_' + Utilities.getUuid().split('-')[0] + '.jpg';
  const blob = Utilities.newBlob(bytes, mimeType, filename);

  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  photos.push(file.getId());
  sh.getRange(found.row, S.photos + 1).setValue(JSON.stringify(photos));
  sh.getRange(found.row, S.updatedAt + 1).setValue(nowIso());
  return { ok: true, fileId: file.getId(), photos };
}

function deletePhoto(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_SPOTS);
  const found = findRowByValue(sh, S.id, body.spotId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[S.authorId] !== userId) return { ok: false, error: 'forbidden' };

  const photos = safeJson(found.data[S.photos], []).filter(id => id !== body.fileId);
  sh.getRange(found.row, S.photos + 1).setValue(JSON.stringify(photos));
  sh.getRange(found.row, S.updatedAt + 1).setValue(nowIso());
  try { DriveApp.getFileById(body.fileId).setTrashed(true); } catch (e) { /* すでに削除済みなどは無視 */ }
  return { ok: true, photos };
}

/* ============================== Favorites（お気に入り） ============================== */
// 列: id, userId, spotId, createdAt
const F = { id: 0, userId: 1, spotId: 2, createdAt: 3 };

function listFavorites(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_FAVORITES);
  const data = sh.getDataRange().getValues();
  const spotIds = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][F.userId] === userId) spotIds.push(data[r][F.spotId]);
  }
  return { ok: true, spotIds };
}

function toggleFavorite(body) {
  const userId = requireAuth(body.token);
  const spotId = body.spotId;
  if (!spotId) return { ok: false, error: 'invalid_spot' };
  const sh = sheet(SHEET_FAVORITES);
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][F.userId] === userId && data[r][F.spotId] === spotId) {
      sh.deleteRow(r + 1);
      return { ok: true, favorited: false };
    }
  }
  sh.appendRow([genId('f'), userId, spotId, nowIso()]);
  return { ok: true, favorited: true };
}

/* ============================== Visits（行った記録） ============================== */
// 列: id, userId, spotId, visitedDate, note, createdAt, updatedAt
const V = { id: 0, userId: 1, spotId: 2, visitedDate: 3, note: 4, createdAt: 5, updatedAt: 6 };
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
function normalizeVisitedDate(v) {
  const s = String(v || '').trim();
  return DATE_RE.test(s) ? s : '';
}
// シートのセル書式が「自動」だと "2026-09-22" のような文字列がGoogle Sheets側で
// 日付型（Dateオブジェクト）に自動変換されてしまい、読み込み時に "2026-09-21T15:00:00.000Z"
// のようなISOタイムスタンプになってしまう（JSTの日付が前日UTCに変換されるため）。
// 読み込み時にDateオブジェクトを検出したら "yyyy-MM-dd" 文字列に戻す。
function visitedDateCellToString(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return normalizeVisitedDate(v) || String(v || '');
}
// 書き込み時は該当セルをあらかじめ「書式なしテキスト」にして自動変換を防ぐ
function writeVisitedDateCell(sh, row, value) {
  const cell = sh.getRange(row, V.visitedDate + 1);
  cell.setNumberFormat('@');
  cell.setValue(value);
}

function listVisits(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_VISITS);
  const data = sh.getDataRange().getValues();
  const visits = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[V.userId] === userId) {
      visits.push({
        id: row[V.id], spotId: row[V.spotId], visitedDate: visitedDateCellToString(row[V.visitedDate]),
        note: row[V.note], createdAt: row[V.createdAt], updatedAt: row[V.updatedAt],
      });
    }
  }
  return { ok: true, visits };
}

function addVisit(body) {
  const userId = requireAuth(body.token);
  const spotId = body.spotId;
  if (!spotId) return { ok: false, error: 'invalid_spot' };
  const visitedDate = normalizeVisitedDate(body.visitedDate);
  if (!visitedDate) return { ok: false, error: 'invalid_date' };
  const note = String(body.note || '').trim();

  const sh = sheet(SHEET_VISITS);
  const id = genId('v');
  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 7).setValues([[id, userId, spotId, '', note, nowIso(), nowIso()]]);
  writeVisitedDateCell(sh, row, visitedDate); // 日付列だけテキスト書式で別途書き込み（自動変換防止）
  return { ok: true, visit: { id, spotId, visitedDate, note } };
}

function updateVisit(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_VISITS);
  const found = findRowByValue(sh, V.id, body.visitId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[V.userId] !== userId) return { ok: false, error: 'forbidden' };

  if (body.visitedDate !== undefined) {
    const visitedDate = normalizeVisitedDate(body.visitedDate);
    if (!visitedDate) return { ok: false, error: 'invalid_date' };
    writeVisitedDateCell(sh, found.row, visitedDate);
  }
  if (body.note !== undefined) sh.getRange(found.row, V.note + 1).setValue(String(body.note || '').trim());
  sh.getRange(found.row, V.updatedAt + 1).setValue(nowIso());
  return { ok: true };
}

function deleteVisit(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_VISITS);
  const found = findRowByValue(sh, V.id, body.visitId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[V.userId] !== userId) return { ok: false, error: 'forbidden' };
  sh.deleteRow(found.row);
  return { ok: true };
}

function deleteVisitsBySpot(spotId) {
  const sh = sheet(SHEET_VISITS);
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][V.spotId] === spotId) sh.deleteRow(r + 1);
  }
}

/* ============================== Comments ============================== */
// 列: id, spotId, text, authorId, authorNickname, createdAt, updatedAt, rating
// ※ rating は既存シートへの後付け列のため末尾に追加（0 = 評価なし, 1〜5 = 星の数）
const C = { id: 0, spotId: 1, text: 2, authorId: 3, authorNickname: 4, createdAt: 5, updatedAt: 6, rating: 7 };
function normalizeRating(v) {
  const n = Number(v);
  return Number.isInteger(n) && n >= 0 && n <= 5 ? n : 0;
}

function listComments(body) {
  const spotId = body.spotId;
  const sh = sheet(SHEET_COMMENTS);
  const data = sh.getDataRange().getValues();
  const comments = [];
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (row[C.spotId] === spotId) {
      comments.push({
        id: row[C.id], spotId: row[C.spotId], text: row[C.text],
        authorId: row[C.authorId], authorNickname: row[C.authorNickname],
        createdAt: row[C.createdAt], updatedAt: row[C.updatedAt],
        rating: normalizeRating(row[C.rating]),
      });
    }
  }
  comments.sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
  return { ok: true, comments };
}

function addComment(body) {
  const userId = requireAuth(body.token);
  const nickname = getUserNickname(userId);
  const text = String(body.text || '').trim();
  const spotId = body.spotId;
  if (!text) return { ok: false, error: 'empty_comment' };
  if (!spotId) return { ok: false, error: 'invalid_spot' };

  const sh = sheet(SHEET_COMMENTS);
  const id = genId('c');
  sh.appendRow([id, spotId, text, userId, nickname, nowIso(), nowIso(), normalizeRating(body.rating)]);
  return { ok: true, id };
}

function updateComment(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_COMMENTS);
  const found = findRowByValue(sh, C.id, body.commentId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[C.authorId] !== userId) return { ok: false, error: 'forbidden' };

  const text = String(body.text || '').trim();
  if (!text) return { ok: false, error: 'empty_comment' };
  sh.getRange(found.row, C.text + 1).setValue(text);
  if (body.rating !== undefined) sh.getRange(found.row, C.rating + 1).setValue(normalizeRating(body.rating));
  sh.getRange(found.row, C.updatedAt + 1).setValue(nowIso());
  return { ok: true };
}

function deleteComment(body) {
  const userId = requireAuth(body.token);
  const sh = sheet(SHEET_COMMENTS);
  const found = findRowByValue(sh, C.id, body.commentId);
  if (!found) return { ok: false, error: 'not_found' };
  if (found.data[C.authorId] !== userId) return { ok: false, error: 'forbidden' };
  sh.deleteRow(found.row);
  return { ok: true };
}

function deleteCommentsBySpot(spotId) {
  const sh = sheet(SHEET_COMMENTS);
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][C.spotId] === spotId) sh.deleteRow(r + 1);
  }
}

/* ⑩ スポットごとの平均評価（Commentsのratingを集計） */
function listSpotStats() {
  const sh = sheet(SHEET_COMMENTS);
  const data = sh.getDataRange().getValues();
  const agg = {}; // spotId -> { sum, count }
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const rating = normalizeRating(row[C.rating]);
    if (rating < 1) continue;
    const spotId = row[C.spotId];
    if (!agg[spotId]) agg[spotId] = { sum: 0, count: 0 };
    agg[spotId].sum += rating;
    agg[spotId].count += 1;
  }
  const stats = {};
  Object.keys(agg).forEach(id => {
    stats[id] = { avg: Math.round((agg[id].sum / agg[id].count) * 10) / 10, count: agg[id].count };
  });
  return { ok: true, stats };
}

/* ============================== 初期セットアップ ============================== */

function setupSheets() {
  createSheetIfMissing(SHEET_USERS, ['id', 'email', 'passwordHash', 'salt', 'nickname', 'mustChangePassword', 'agreedTerms', 'createdAt', 'updatedAt']);
  createSheetIfMissing(SHEET_SPOTS, ['id', 'name', 'category', 'lat', 'lng', 'description', 'facilities', 'openHours', 'fee', 'authorId', 'authorNickname', 'createdAt', 'updatedAt', 'sleepStatus', 'photos', 'tel']);
  createSheetIfMissing(SHEET_COMMENTS, ['id', 'spotId', 'text', 'authorId', 'authorNickname', 'createdAt', 'updatedAt', 'rating']);
  createSheetIfMissing(SHEET_FAVORITES, ['id', 'userId', 'spotId', 'createdAt']);
  createSheetIfMissing(SHEET_VISITS, ['id', 'userId', 'spotId', 'visitedDate', 'note', 'createdAt', 'updatedAt']);
}

/**
 * 【既存デプロイ向けマイグレーション】
 * 車中泊の可否(sleepStatus)機能を追加した際、Spotsシートが既に存在する場合は
 * setupSheets() では列が増えないため、この関数を一度だけ手動実行してください。
 * （スクリプトエディタの関数選択で migrateAddSleepStatusColumn を選び ▶実行）
 */
function migrateAddSleepStatusColumn() {
  const sh = sheet(SHEET_SPOTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('sleepStatus') !== -1) return; // すでに追加済み
  sh.getRange(1, lastCol + 1).setValue('sleepStatus');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, lastCol + 1, lastRow - 1, 1).setValue('unclear'); // 既存投稿は「要確認」で初期化
  }
}

/**
 * 【既存デプロイ向けマイグレーション】
 * 写真投稿(photos)機能を追加した際、Spotsシートに列を追加します。一度だけ手動実行してください。
 */
function migrateAddPhotosColumn() {
  const sh = sheet(SHEET_SPOTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('photos') !== -1) return; // すでに追加済み
  sh.getRange(1, lastCol + 1).setValue('photos');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, lastCol + 1, lastRow - 1, 1).setValue('[]'); // 既存投稿は写真なしで初期化
  }
}

/**
 * 【既存デプロイ向けマイグレーション】
 * 評価(rating)機能を追加した際、Commentsシートに列を追加します。一度だけ手動実行してください。
 */
function migrateAddRatingColumn() {
  const sh = sheet(SHEET_COMMENTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('rating') !== -1) return; // すでに追加済み
  sh.getRange(1, lastCol + 1).setValue('rating');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, lastCol + 1, lastRow - 1, 1).setValue(0); // 既存コメントは「評価なし」で初期化
  }
}

/**
 * 【既存デプロイ向けマイグレーション】
 * 電話番号(tel)項目を追加した際、Spotsシートに列を追加します。一度だけ手動実行してください。
 */
function migrateAddTelColumn() {
  const sh = sheet(SHEET_SPOTS);
  const lastCol = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (header.indexOf('tel') !== -1) return; // すでに追加済み
  sh.getRange(1, lastCol + 1).setValue('tel');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, lastCol + 1, lastRow - 1, 1).setValue(''); // 既存投稿は電話番号なしで初期化
  }
}

/**
 * 【既存デプロイ向けマイグレーション】
 * お気に入り(favorites)機能を追加した際、Favoritesシートが無ければ作成します。一度だけ手動実行してください。
 */
function migrateAddFavoritesSheet() {
  createSheetIfMissing(SHEET_FAVORITES, ['id', 'userId', 'spotId', 'createdAt']);
}

/**
 * 【既存デプロイ向けマイグレーション】
 * 「行った記録」(Visits)機能を追加した際、Visitsシートが無ければ作成します。一度だけ手動実行してください。
 */
function migrateAddVisitsSheet() {
  createSheetIfMissing(SHEET_VISITS, ['id', 'userId', 'spotId', 'visitedDate', 'note', 'createdAt', 'updatedAt']);
}

/**
 * 【既存デプロイ向け修正】
 * Visitsシートのvisiteddate列がGoogle Sheets側で日付型に自動変換され、
 * ポップアップ上に "2026-09-21T15:00:00.000Z" のようなISO形式で表示されてしまっていた不具合の修正。
 * 該当列をテキスト書式にし、既存の日付型セルを "yyyy-MM-dd" 文字列に書き戻します。一度だけ手動実行してください。
 */
function migrateFixVisitDatesToText() {
  const sh = sheet(SHEET_VISITS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) { Logger.log('データ行がありません。'); return; }
  const range = sh.getRange(2, V.visitedDate + 1, lastRow - 1, 1);
  const fixed = range.getValues().map(([v]) => [visitedDateCellToString(v)]);
  range.setNumberFormat('@');
  range.setValues(fixed);
  Logger.log(`${fixed.length}件の訪問日を文字列形式に修正しました。`);
}

function createSheetIfMissing(name, header) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
}

/* ============================== 音声入力（Gemini APIで文字起こし・整形） ============================== */

function transcribeVoice(body) {
  requireAuth(body.token); // ログイン会員のみ（API利用量制御のため）
  if (!PROPS.getProperty('GEMINI_API_KEY')) return { ok: false, error: 'gemini_not_configured' };

  const audioBase64 = body.audioBase64;
  if (!audioBase64) return { ok: false, error: 'invalid_audio' };
  const mimeType = body.mimeType || 'audio/webm';

  const prompt =
    'あなたは車中泊・キャンプスポット共有サイトの入力補助アシスタントです。' +
    '添付の音声は、投稿者がスポットの説明を口頭で話したものです。' +
    '音声の内容を日本語で文字起こしし、「えーと」「あの」などのフィラーや言い淀みを取り除いたうえで、' +
    'そのままスポットの説明文として使える自然な文章に整えてください。' +
    '出力は整形後の説明文の本文のみとし、前置き・見出し・カギ括弧などの装飾は付けないでください。';

  const contents = [{
    parts: [
      { text: prompt },
      { inline_data: { mime_type: mimeType, data: audioBase64 } },
    ],
  }];

  try {
    // 選択モデル（body.model）で実行。使えない時は汎用の最新モデルへ自動切替
    const r = callGeminiWithModel_(body.model || 'auto', contents);
    return { ok: true, text: r.text, model: r.model };
  } catch (e) {
    return { ok: false, error: 'gemini_error', message: String(e.message || e).slice(0, 300) };
  }
}

/* ============================== Geminiモデル選択＆自動フォールバック ============================== */
// index.html の GEMINI_MODELS と合わせる。
// 選んだモデルが使えない時（404/400の未対応・429混雑・5xx障害）は GEMINI_AUTO_CHAIN の順に切り替える。

/** 画面で選べるモデル（index.html の GEMINI_MODELS と合わせる） */
const GEMINI_ALLOWED_MODELS = [
  'gemini-4-hydrogen',
  'gemini-4-neon',
  'gemini-4-carbon',
  'gemini-4-argon',
  'gemini-3.8-flash',
  'gemini-3.5-flash',
];

/** 自動モード（汎用の最新）の候補。先頭から順に試します */
const GEMINI_AUTO_CHAIN = [
  'gemini-4-hydrogen',
  'gemini-3.8-flash',
  'gemini-3.5-flash',
  'gemini-flash-latest',
];

/** 選択モデル → 試す順番のリストを作る */
function buildGeminiModelChain_(requested) {
  const chain = [];
  if (requested && requested !== 'auto' && GEMINI_ALLOWED_MODELS.indexOf(requested) !== -1) {
    chain.push(requested);
  }
  GEMINI_AUTO_CHAIN.forEach(function (m) {
    if (chain.indexOf(m) === -1) chain.push(m);
  });
  return chain;
}

/**
 * Gemini generateContent をモデル自動切替つきで呼び出す
 * @param {string} requestedModel 'auto' または GEMINI_ALLOWED_MODELS のいずれか
 * @param {Array}  contents       Gemini API の contents 配列
 * @param {Object} [generationConfig]
 * @return {{text: string, model: string}}
 */
function callGeminiWithModel_(requestedModel, contents, generationConfig) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が未設定です');

  const chain = buildGeminiModelChain_(requestedModel);
  const body = { contents: contents };
  if (generationConfig) body.generationConfig = generationConfig;
  const payload = JSON.stringify(body);

  let lastError = '';
  for (let i = 0; i < chain.length; i++) {
    const model = chain[i];
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
      encodeURIComponent(model) + ':generateContent?key=' + apiKey;

    // 同じモデルで一時的な混雑(429/503)なら1回だけ再試行
    for (let attempt = 0; attempt < 2; attempt++) {
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: payload,
        muteHttpExceptions: true,
      });
      const code = res.getResponseCode();

      if (code === 200) {
        const json = JSON.parse(res.getContentText());
        const parts = (((json.candidates || [])[0] || {}).content || {}).parts || [];
        const text = parts.map(function (p) { return p.text || ''; }).join('').trim();
        if (text) return { text: text, model: model };
        lastError = model + ': 空の応答';
        break; // 次のモデルへ
      }

      lastError = model + ': HTTP ' + code + ' ' + res.getContentText().slice(0, 200);
      if ((code === 429 || code === 503) && attempt === 0) {
        Utilities.sleep(1500);
        continue;
      }
      break; // 404/400(未対応モデル)・その他エラーは次のモデルへ
    }
    console.warn('Geminiモデル切替: ' + lastError);
  }
  throw new Error('すべてのGeminiモデルで失敗しました（' + lastError + '）');
}

/** 動作確認用：GASエディタでこの関数を選んで▶実行 */
function testGeminiModels() {
  const contents = [{ parts: [{ text: '「こんにちは」とだけ返してください。' }] }];
  ['auto'].concat(GEMINI_ALLOWED_MODELS).forEach(function (m) {
    try {
      const r = callGeminiWithModel_(m, contents);
      console.log(m + ' → ' + r.model + ' : ' + r.text);
    } catch (e) {
      console.log(m + ' → 失敗: ' + e.message);
    }
  });
}
