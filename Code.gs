// 校外学習・チェックポイント通過管理 v4（Google Apps Script V8）
const SPREADSHEET_ID = ''; // 既存のスプレッドシートIDを設定
const CHAT_WINDOW = 100;

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action || p.view === 'app') return hostedPage_(p);
  const cb = p.callback || '';
  if (cb && !/^[A-Za-z_$][0-9A-Za-z_$]{0,100}$/.test(cb)) {
    return output_({ok:false, error:'コールバック名が正しくありません'}, '');
  }
  return output_(handle_(p, false), cb);
}

// Googleのログインを通ったウェブアプリ内で、同じ画面を開く。
function hostedPage_(p) {
  let html;
  if (typeof bundledAppHtml_ === 'function') {
    html = bundledAppHtml_();
  } else {
    try {
      html = HtmlService.createHtmlOutputFromFile('Index').getContent();
    } catch (_) {
      html = HtmlService.createHtmlOutputFromFile('index').getContent();
    }
  }
  const event = String(p.event || '').trim().toLowerCase();
  const boot = JSON.stringify({
    appUrl: ScriptApp.getService().getUrl() || '',
    eventId: /^[a-z0-9]{8}$/.test(event) ? event : ''
  }).replace(/</g, '\\u003c');
  const page = html.replace(/<head>/i, () => '<head><script>window.KOUGAI_HOSTED=' + boot + ';</script>');
  return HtmlService.createHtmlOutput(page)
    .setTitle('校外学習・チェックポイント通過管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// HTML Serviceからの呼び出し。HTTP経由と同じ操作の検証・受付番号を使う。
function appRequest(params) {
  if (!params || typeof params !== 'object' || Array.isArray(params)) {
    return {ok:false, code:'INVALID_INPUT', error:'入力を確認してください'};
  }
  if (params.action === 'capabilities') {
    try {
      openWorkbook_().getId();
      return {ok:true, version:4};
    } catch (err) {
      return {ok:false, code:err.clientCode || 'SHEET_UNAVAILABLE',
        error:err.clientCode ? err.message : '行事の保存先に接続できませんでした。'};
    }
  }
  const request = Object.assign({}, params, {v:'4'});
  if (['createEvent','saveConfig','toggle','reset','sendMessage'].includes(request.action) && !request.requestId) {
    return {ok:false, code:'INVALID_INPUT', error:'操作の受付番号がありません。画面を開き直してください'};
  }
  // handle_内の許可リストとイベント・入力の検証を必ず通す。
  return JSON.parse(JSON.stringify(handle_(request, true)));
}

function openWorkbook_() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (_) {
    if (!SPREADSHEET_ID) invalid_('GASのSPREADSHEET_IDに行事の保存先を設定してください。', 'SHEET_NOT_CONFIGURED');
    invalid_('行事の保存先を開けません。管理者にスプレッドシートIDと実行アカウントの権限を確認してください。', 'SHEET_UNAVAILABLE');
  }
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    return output_(handle_(p, true), '');
  } catch (err) {
    return output_({ok:false, error:'リクエストの形式が正しくありません'}, '');
  }
}

function output_(result, callback) {
  const json = JSON.stringify(result);
  // JSONPでも「__proto__」等を含む班名・CP名を通常のデータキーとして保つ。
  const body = callback ? callback + '(JSON.parse(' + JSON.stringify(json) + '));' : json;
  return ContentService.createTextOutput(body)
    .setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function invalid_(message, code) {
  const err = new Error(message);
  err.clientCode = code || 'INVALID_INPUT';
  throw err;
}

function handle_(p, post) {
  try {
    if (!p || typeof p !== 'object' || Array.isArray(p)) invalid_('入力を確認してください');
    const a = p.action;
    const v4 = String(p.v) === '4';
    if (a === 'capabilities') return {ok:true, version:4};
    if (!['createEvent','getState','getReceipt','saveConfig','toggle','reset','sendMessage'].includes(a)) {
      invalid_('この操作には対応していません');
    }
    if (a !== 'createEvent' && !(a === 'getReceipt' && !p.eventId) &&
        (typeof p.eventId !== 'string' || !/^[a-z0-9]{8}$/.test(p.eventId))) {
      invalid_('イベントIDを確認してください');
    }
    if (p.requestId || a === 'getReceipt') checkId_(p.requestId);
    if (!post && a !== 'getReceipt' && a !== 'getState' && (v4 || a === 'sendMessage')) {
      invalid_('更新にはPOSTを使用してください');
    }
    const ss = openWorkbook_();
    if (a === 'getReceipt') {
      const row = receipt_(ss, p.eventId, p.requestId);
      return row ? {ok:true, found:true, result:row.result} : {ok:true, found:false};
    }
    if (a === 'getState') return state_(ss, p.eventId, v4);
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(25000)) return {ok:false, code:'BUSY', error:'混み合っています。もう一度お試しください'};
    try {
      const scope = a === 'createEvent' ? '' : p.eventId;
      if (p.requestId) {
        const previous = receipt_(ss, scope, p.requestId);
        if (previous) {
          if (previous.action !== a) invalid_('受付番号が重複しています');
          return previous.result;
        }
      }
      let result;
      try {
        if (a === 'createEvent') result = create_(ss, p, v4);
        else {
          const sheet = eventSheet_(ss, p.eventId);
          if (a === 'saveConfig') result = saveConfig_(sheet, p, v4);
          if (a === 'toggle') result = toggle_(sheet, p, v4);
          if (a === 'reset') result = reset_(sheet, p);
          if (a === 'sendMessage') result = send_(ss, p);
        }
      } catch (err) {
        if (!err.clientCode) throw err;
        result = {ok:false, code:err.clientCode, error:err.message};
      }
      if (p.requestId) {
        // 表示名・本文を含む結果も、JSONとしてセルに格納する。
        append_(aux_(ss, receiptSheetName_(scope), ['requestId','action','result']),
                [p.requestId, a, JSON.stringify(result)]);
      }
      SpreadsheetApp.flush();
      return result;
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return {ok:false, code:err.clientCode || 'SERVICE_ERROR',
      error:err.clientCode ? err.message : '同期できませんでした。時間をおいてもう一度お試しください'};
  }
}

function checkId_(id) {
  if (typeof id !== 'string' || !/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(id)) {
    invalid_('受付番号が正しくありません');
  }
}

function eventSheet_(ss, id) {
  const sheet = ss.getSheetByName(id);
  if (!sheet || sheet.getRange(1,1,1,1).getValues()[0][0] !== '__meta__') {
    invalid_('イベントが見つかりません', 'EVENT_NOT_FOUND');
  }
  return sheet;
}

function append_(sheet, row) {
  const n = sheet.getLastRow() + 1;
  if (n > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 100);
  sheet.getRange(n,1,1,row.length).setNumberFormat('@').setValues([row]);
}

function aux_(ss, name, header) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (!sheet.getLastRow()) append_(sheet, header);
  return sheet;
}

function findRow_(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const found = sheet.getRange(2,1,sheet.getLastRow()-1,1).createTextFinder(id)
    .matchEntireCell(true).matchCase(true).findNext();
  return found ? found.getRow() : 0;
}

function receiptSheetName_(eventId) {
  return '__receipts_' + (eventId || 'create');
}

function receipt_(ss, eventId, id) {
  const sheet = ss.getSheetByName(receiptSheetName_(eventId));
  const row = findRow_(sheet, id);
  if (!row) return null;
  const values = sheet.getRange(row,2,1,2).getValues()[0];
  return {action:values[0], result:JSON.parse(values[1])};
}

// v4で新しく書くセルは文字列書式を使い、式・特殊な先頭文字は可逆に保護する。
function sheetText_(value) { return /^[=+@'~-]/.test(value) ? '~' + value : value; }
function unsheet_(value) {
  const s = String(value);
  return s.startsWith('~') ? s.slice(1) : s;
}

function label_(value, limit, title, legacyValue) {
  if (legacyValue !== undefined && value === legacyValue && typeof value === 'string' && value.trim()) return value;
  if (typeof value !== 'string' || !value.trim() || value.trim().length > limit) {
    invalid_(title + 'は1～' + limit + '文字で入力してください');
  }
  return value.trim();
}

// 同じ作成リクエストを再送した場合、受付の書き込み前に通信が途切れても同じイベントを発見する。
function candidate_(seed) {
  let a = 2166136261, b = 5381;
  for (let i=0; i<seed.length; i++) {
    a = Math.imul(a ^ seed.charCodeAt(i), 16777619);
    b = Math.imul(b, 33) ^ seed.charCodeAt(i);
  }
  return ((a >>> 0) * 1024 + ((b >>> 0) % 1024)).toString(36).padStart(8, '0').slice(-8);
}

function create_(ss, p, v4) {
  const name = label_(p.eventName, 120, '行事名');
  const initial = Object.assign({}, p);
  if (p.checkpoints === undefined) initial.checkpoints = [];
  if (p.groups === undefined) initial.groups = [];
  const lists = configLists_(initial, v4, {checkpoints:[], groups:[]}, true);
  const version = v4 && (lists.checkpoints !== '[]' || lists.groups !== '[]') ? 1 : 0;
  const seed = p.requestId || Utilities.getUuid();
  const marker = 'v4:' + seed;
  for (let attempt = 0; attempt < 100; attempt++) {
    const id = candidate_(seed + '/' + attempt);
    const old = ss.getSheetByName(id);
    if (old) {
      const meta = old.getRange(1,1,1,4).getValues()[0];
      if (meta[0] === '__meta__' && meta[3] === marker) {
        return {ok:true, eventId:id, eventName:unsheet_(meta[1])};
      }
      continue;
    }
    const sheet = ss.insertSheet(id);
    sheet.getRange(1,1,2,5).setNumberFormat('@').setValues([
      ['__meta__', sheetText_(name), new Date().toISOString(), marker, ''],
      ['__config__', lists.checkpoints, lists.groups, String(version), '']
    ]);
    return {ok:true, eventId:id, eventName:name};
  }
  throw new Error('event id unavailable');
}

function config_(rows) {
  const meta = rows[0], config = rows[1];
  const name = String(meta[3]).startsWith('v4:') ? unsheet_(meta[1]) : String(meta[1]);
  const configured = config[0] === '__config__';
  const checkpoints = configured ? JSON.parse(config[1] || '[]') : [];
  const rawGroups = configured ? JSON.parse(config[2] || '[]') : [];
  const groups = rawGroups.map(g => typeof g === 'string' ? {id:g,name:g} : g);
  const savedBy = String(config[4] || '');
  const legacySaved = savedBy.startsWith('v3:');
  return {meta:meta, name:name, checkpoints:checkpoints, groups:groups, legacySaved:legacySaved,
    version:configured ? Number(config[3]) || 0 : 0, lastRequest:legacySaved ? savedBy.slice(3) : savedBy};
}

function asList_(value) {
  if (typeof value === 'string') {
    try { value = JSON.parse(value); } catch (err) { invalid_('設定の形式を確認してください'); }
  }
  if (!Array.isArray(value)) invalid_('設定の形式を確認してください');
  return value;
}

function configLists_(p, v4, c, allowEmpty) {
  const oldCPs = new Set(c.checkpoints);
  const oldGroups = new Map(c.groups.map(g=>[g.id,g.name]));
  const checkpoints = asList_(p.checkpoints).map(cp =>
    label_(cp,80,'チェックポイント名',oldCPs.has(cp) ? cp : undefined));
  const groups = asList_(p.groups).map(item => {
    const group = !v4 && typeof item === 'string' ? {id:item,name:item} : item;
    if (!group || typeof group.id !== 'string' || !group.id ||
        (!oldGroups.has(group.id) && v4 && !/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(group.id))) {
      invalid_('班の情報を確認してください');
    }
    return {id:group.id, name:label_(group.name,80,'班名',oldGroups.get(group.id))};
  });
  if ((!allowEmpty && (!checkpoints.length || !groups.length)) ||
      checkpoints.length > Math.max(200,c.checkpoints.length) || groups.length > Math.max(400,c.groups.length)) {
    invalid_('チェックポイント数と班数を確認してください');
  }
  if (new Set(checkpoints).size !== checkpoints.length ||
      new Set(groups.map(g=>g.id)).size !== groups.length ||
      new Set(groups.map(g=>g.name)).size !== groups.length) {
    invalid_('同じ班名やチェックポイント名が重複しています');
  }
  const jsonCPs = JSON.stringify(checkpoints);
  const jsonGroups = JSON.stringify(v4 ? groups : groups.map(g=>g.name));
  if (jsonCPs.length > 49000 || jsonGroups.length > 49000) invalid_('設定が大きすぎます。名前や件数を調整してください');
  return {checkpoints:jsonCPs, groups:jsonGroups};
}

function saveConfig_(sheet, p, v4) {
  const c = config_(sheet.getRange(1,1,2,5).getValues());
  if (p.requestId && c.lastRequest === p.requestId) return {ok:true, configVersion:c.version};
  if (v4) {
    if (p.configVersion === null || p.configVersion === undefined || p.configVersion === '' ||
        Number(p.configVersion) !== c.version) {
      invalid_('別の端末で設定が更新されています。最新の設定を確認してください', 'CONFIG_CONFLICT');
    }
  } else if (c.version > 0 && !c.legacySaved) {
    invalid_('この行事の設定は更新版アプリで編集してください', 'CONFIG_CONFLICT');
  }
  const name = label_(p.eventName, 120, '行事名', c.name);
  const lists = configLists_(p, v4, c, false);
  const version = c.version + 1;
  const savedBy = (v4 ? '' : 'v3:') + (p.requestId || '');
  const marker = String(c.meta[3]).startsWith('v4:') ? c.meta[3] : 'v4:';
  // メタデータと設定だけを一括更新。3行目以降の通過記録に触れない。
  sheet.getRange(1,1,2,5).setNumberFormat('@').setValues([
    ['__meta__', sheetText_(name), c.meta[2] instanceof Date ? c.meta[2].toISOString() : c.meta[2], marker, ''],
    ['__config__', lists.checkpoints, lists.groups, String(version), savedBy]
  ]);
  return {ok:true, configVersion:version};
}

function passageRows_(sheet) {
  const n = sheet.getLastRow();
  return n < 3 ? [] : sheet.getRange(3,1,n-2,5).getValues();
}

function readPassages_(rows) {
  let records = Object.create(null);
  for (const row of rows) {
    const tag = row[3];
    if (tag === 'reset') { records = Object.create(null); continue; }
    if (row[0] === '' || row[1] === '') continue;
    const group = tag === 'add' || tag === 'remove' ? unsheet_(row[0]) : String(row[0]);
    const cp = tag === 'add' || tag === 'remove' ? unsheet_(row[1]) : String(row[1]);
    if (tag === 'remove') {
      if (records[group]) delete records[group][cp];
    } else if (row[2] !== '' && row[2] !== null) {
      if (!records[group]) records[group] = Object.create(null);
      records[group][cp] = row[2] instanceof Date
        ? Utilities.formatDate(row[2], 'Asia/Tokyo','HH:mm') : String(row[2]);
    }
  }
  return records;
}

function state_(ss, id, v4) {
  const sheet = ss.getSheetByName(id);
  const rows = sheet ? sheet.getRange(1,1,Math.max(2,sheet.getLastRow()),5).getValues() : [];
  if (!rows.length || rows[0][0] !== '__meta__') invalid_('イベントが見つかりません', 'EVENT_NOT_FOUND');
  const c = config_(rows);
  const records = readPassages_(rows.slice(2));
  const result = {ok:true, eventId:id, eventName:c.name, checkpoints:c.checkpoints,
    groups:c.groups, records:records};
  if (v4) {
    result.configVersion = c.version;
    result.messages = messages_(ss, id);
  } else {
    result.groups = c.groups.map(g=>g.name);
    result.records = Object.create(null);
    c.groups.forEach(g=>{ if (records[g.id]) result.records[g.name] = records[g.id]; });
  }
  return result;
}

// 通過操作は追記式。v3の既存行はそのまま読み、取消・リセットも時系列順に反映する。
// 操作と受付番号を同じ行に残し、応答が途切れた場合の再送を判別する。
function toggle_(sheet, p, v4) {
  const values = sheet.getRange(1,1,Math.max(2,sheet.getLastRow()),5).getValues();
  const c = config_(values), rows = values.slice(2);
  if (p.requestId) {
    const applied = rows.find(row=>row[4] === p.requestId);
    if (applied) return applied[3] === 'remove'
      ? {ok:true, action:'removed'} : {ok:true, action:'added', time:String(applied[2])};
  }
  const group = c.groups.find(g=>(v4 ? g.id : g.name) === p.group);
  if (!group || !c.checkpoints.includes(p.cp)) invalid_('班またはチェックポイントが変更されました。画面を更新してください');
  const op = p.op || '';
  if (op !== 'add' && op !== 'remove' && (v4 || op !== '')) invalid_('通過操作を確認してください');
  const records = readPassages_(rows);
  const old = records[group.id] && records[group.id][p.cp];
  const remove = op === 'remove' || (op === '' && old);
  if (remove && !old) return {ok:true, action:'already_removed'};
  if (!remove && old) return {ok:true, action:'already_added', time:old};
  const time = remove ? '' : Utilities.formatDate(new Date(), 'Asia/Tokyo','HH:mm');
  append_(sheet, [sheetText_(group.id),sheetText_(p.cp),time,remove ? 'remove' : 'add',p.requestId || '']);
  return remove ? {ok:true, action:'removed'} : {ok:true, action:'added',time:time};
}

function reset_(sheet, p) {
  if (!p.requestId || !passageRows_(sheet).some(row=>row[4] === p.requestId)) {
    // 設定未保存のイベントでも2行目は常に設定行として予約する。
    if (sheet.getLastRow() === 1) append_(sheet, ['__config__','[]','[]','0','']);
    append_(sheet, ['','','','reset',p.requestId || '']);
  }
  return {ok:true};
}

function messages_(ss, eventId) {
  const sheet = ss.getSheetByName('__chat_' + eventId);
  const last = sheet ? sheet.getLastRow() : 0;
  if (last < 2) return [];
  const count = Math.min(CHAT_WINDOW, last - 1);
  return sheet.getRange(last-count+1,2,count,1).getValues().map(row=>JSON.parse(row[0]));
}

function send_(ss, p) {
  checkId_(p.id);
  const name = label_(p.name,24,'参加者名');
  const text = label_(p.text,300,'メッセージ');
  const sheet = aux_(ss, '__chat_' + p.eventId, ['messageId','message']);
  const row = findRow_(sheet, p.id);
  if (row) {
    const old = JSON.parse(sheet.getRange(row,2,1,1).getValues()[0][0]);
    if (old.name !== name || old.text !== text) invalid_('送信番号が重複しています。内容を確認してください');
    return {ok:true, message:old};
  }
  const message = {id:p.id, name:name, text:text, at:new Date().toISOString()};
  append_(sheet, [p.id,JSON.stringify(message)]);
  return {ok:true, message:message};
}
