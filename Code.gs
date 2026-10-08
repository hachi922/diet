/**
 * ダイエット記録 同期用 Google Apps Script
 * スプレッドシートに紐づけて使います(拡張機能 → Apps Script)。
 * 「weight」「diary」シートは初回アクセス時に自動で作られます。
 */
const KEY = 'うた';   // アプリの設定画面と同じものを入れる
const TZ = 'Asia/Tokyo';
const FOLDER_ID = '1rprwGesWeHmcHjriF5-lO7TnXaIBofKy';                      // 写真の保存先フォルダのID。空のままなら下の名前のフォルダを自動作成
const FOLDER_NAME = 'ダイエット記録写真';   // FOLDER_IDが空のときに、マイドライブに自動作成するフォルダ名(非公開)
const MAX_PHOTO_B64 = 4 * 1024 * 1024;      // 1枚あたりの上限(base64文字数)
const MP3_FILE_ID = '1oawRqVnQClqsJoA3gnkisSod9C1ME0pM';                    // 再生したいMP3のファイルID(Googleドライブ)。空のままならボタンはエラーになります
const MAX_AUDIO_BYTES = 10 * 1024 * 1024;  // MP3の上限サイズ(10MB)

const SCHEMA = {
  weight: {
    cols: ['date', 'weight', 'updatedAt'],
    text: ['date']
  },
  diary: {
    cols: ['date', 'morning', 'lunch', 'dinner', 'snack', 'bowelCount', 'bowelNote', 'period', 'exercise', 'exerciseMin', 'memo', 'updatedAt', 'photos'],
    text: ['date', 'morning', 'lunch', 'dinner', 'snack', 'bowelNote', 'period', 'exercise', 'memo', 'photos']
  }
};

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const s = SCHEMA[name];
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, s.cols.length).setValues([s.cols]);
    sh.setFrozenRows(1);
  } else {
    // 既存シートに新しい列(photos)が無ければ追加する
    const s = SCHEMA[name];
    if (sh.getLastColumn() < s.cols.length) {
      sh.getRange(1, 1, 1, s.cols.length).setValues([s.cols]);
      s.text.forEach(function (c) {
        sh.getRange(2, s.cols.indexOf(c) + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
      });
    }
  }
  return sh;
}

/* ---------- 写真(Googleドライブ) ---------- */
function getFolder_() {
  // フォルダIDが指定されていれば、常にそのフォルダを使う
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 作り直す */ }
  }
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', folder.getId());
  return folder;
}

function inFolder_(file) {
  const fid = getFolder_().getId();
  const it = file.getParents();
  while (it.hasNext()) { if (it.next().getId() === fid) return true; }
  return false;
}

function getPhoto_(id) {
  const file = DriveApp.getFileById(id);
  if (!inFolder_(file)) return { error: 'not found' };
  const blob = file.getBlob();
  return { mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

function uploadPhoto_(name, data) {
  if (!data || data.length > MAX_PHOTO_B64) return { error: '画像が大きすぎます' };
  const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/jpeg', String(name || 'photo.jpg'));
  const file = getFolder_().createFile(blob);
  return { id: file.getId() };
}

function deletePhoto_(id) {
  const file = DriveApp.getFileById(id);
  if (!inFolder_(file)) return { error: 'not found' };
  file.setTrashed(true);
  return { ok: true };
}

/* ---------- MP3(Googleドライブ) ---------- */
function getAudio_(metaOnly) {
  if (!MP3_FILE_ID) return { error: 'MP3が未設定です(Code.gsのMP3_FILE_ID)' };
  const file = DriveApp.getFileById(MP3_FILE_ID);
  const ver = String(file.getLastUpdated().getTime());
  const size = file.getSize();
  if (size > MAX_AUDIO_BYTES) return { error: 'MP3が大きすぎます(10MBまで)' };
  if (metaOnly) return { id: MP3_FILE_ID, ver: ver, size: size };
  const blob = file.getBlob();
  let mime = blob.getContentType();
  if (!mime || mime.indexOf('audio') !== 0) mime = 'audio/mpeg';
  return { id: MP3_FILE_ID, ver: ver, mime: mime, data: Utilities.base64Encode(blob.getBytes()) };
}

/** 初回の権限承認用:Apps Scriptエディタでこの関数を選んで「実行」する */
function authorize_() {
  getFolder_();
}

function normDate_(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
}

function readAll_(name) {
  const s = SCHEMA[name];
  const sh = getSheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, s.cols.length).getValues();
  return vals.map(function (row) {
    const o = {};
    s.cols.forEach(function (c, i) {
      let v = row[i];
      if (c === 'date') v = normDate_(v);
      o[c] = v;
    });
    return o;
  }).filter(function (o) { return o.date; });
}

function upsert_(name, items) {
  if (!items || !items.length) return;
  const s = SCHEMA[name];
  const sh = getSheet_(name);
  const last = sh.getLastRow();
  const idx = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
      idx[normDate_(r[0])] = i + 2;
    });
  }
  const adds = [];
  items.forEach(function (it) {
    if (!it.date) return;
    const row = s.cols.map(function (c) {
      return it[c] === undefined || it[c] === null ? '' : it[c];
    });
    if (idx[it.date]) {
      sh.getRange(idx[it.date], 1, 1, s.cols.length).setValues([row]);
    } else {
      adds.push(row);
    }
  });
  if (adds.length) {
    const start = Math.max(sh.getLastRow(), 1) + 1;
    const need = start + adds.length - 1 - sh.getMaxRows();
    if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
    // 先頭が「=」「+」などでも数式にならないよう、文字列列は書式を「書式なしテキスト」に
    s.text.forEach(function (c) {
      sh.getRange(start, s.cols.indexOf(c) + 1, adds.length, 1).setNumberFormat('@');
    });
    sh.getRange(start, 1, adds.length, s.cols.length).setValues(adds);
  }
}

function doGet(e) {
  if (((e && e.parameter && e.parameter.key) || '') !== KEY) return json_({ error: 'unauthorized' });
  if (e.parameter.action === 'audio') {
    try { return json_(getAudio_(e.parameter.meta === '1')); } catch (err) { return json_({ error: String(err) }); }
  }
  if (e.parameter.action === 'photo') {
    try { return json_(getPhoto_(e.parameter.id)); } catch (err) { return json_({ error: String(err) }); }
  }
  return json_({ weight: readAll_('weight'), diary: readAll_('diary') });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if ((body.key || '') !== KEY) return json_({ error: 'unauthorized' });
    if (body.action === 'uploadPhoto') return json_(uploadPhoto_(body.name, body.data));
    if (body.action === 'deletePhoto') return json_(deletePhoto_(body.id));
    upsert_('weight', body.weight);
    upsert_('diary', body.diary);
    return json_({ ok: true });
  } catch (err) {
    return json_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
