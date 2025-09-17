/**
 * Code.gs — FINAL FIX
 * - 讀取 quizzes 時，自動回填舊前端期望的欄位：choices{} / answer
 * - 寫入 quizzes 時，自動展開 choiceA~D / correct / startTime
 * - 刪除獎品「前端只覆寫 setStorage」也會反映到試算表：rewards 用「整表覆寫」(含刪除)
 * - 其它：增量寫入、60s 快取、錯誤紀錄、歷史裁切、deleteStorage 真刪
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index').setTitle('班級管理系統');
}

function _nowISO() { return new Date().toISOString(); }

function _logError(fn, msg) {
  var sh = _ensureExactSheet('Logs', ['time','function','message']);
  sh.appendRow([_nowISO(), fn, msg]);
}

function _ensureExactSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  var cur = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
  var ok = true;
  for (var i = 0; i < headers.length; i++) {
    if (String(cur[i] || '') !== String(headers[i])) { ok = false; break; }
  }
  if (!ok) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function _getAllObjects(name, headers, jsonFields) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(name);
    if (cached) return JSON.parse(cached);
  } catch (e) { _logError('_getAllObjects-cacheGet', e.toString()); }

  var sh = _ensureExactSheet(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rng = sh.getRange(2, 1, last - 1, headers.length).getValues();
  var arr = rng.map(function (row) {
    var o = {};
    for (var i = 0; i < headers.length; i++) {
      var k = headers[i], v = row[i];
      if (jsonFields && jsonFields.indexOf(k) >= 0 && typeof v === 'string' && v) {
        try { o[k] = JSON.parse(v); } catch (e) { o[k] = v; }
      } else {
        o[k] = v;
      }
    }
    return o;
  });

  // 特別處理：回填 quizzes 舊欄位 (choices/answer)，讓前端不用改也能讀到
  if (name === 'Quizzes') {
    arr.forEach(function(q){
      if (!q.choices) {
        q.choices = {
          A: q.choiceA || '',
          B: q.choiceB || '',
          C: q.choiceC || '',
          D: q.choiceD || ''
        };
      }
      if (!q.answer && q.correct !== undefined) {
        q.answer = q.correct;  // 舊前端顯示 answer，這裡幫你補成 correct
      }
    });
  }

  try { CacheService.getScriptCache().put(name, JSON.stringify(arr), 60); }
  catch (e) { _logError('_getAllObjects-cachePut', e.toString()); }

  return arr;
}

function _valForCell(field, value) {
  if ((field === 'winners' || field === 'affectedStudents') && typeof value !== 'string') {
    try { return JSON.stringify(value || []); } catch (e) { return '[]'; }
  }
  if (value === undefined) return "";
  return value;
}

/** 完全覆寫（含刪除） */
function _replaceAllRows(name, headers, items, jsonFields) {
  var sh = _ensureExactSheet(name, headers);
  sh.clear(); // 會把標題也清掉
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (!items || !items.length) return { replaced: 0 };
  var rows = items.map(function(it){
    return headers.map(function(h){ return _valForCell(h, it[h]); });
  });
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();
  try { CacheService.getScriptCache().remove(name); }
  catch (e) { _logError('_replaceAllRows-cacheRemove', e.toString()); }
  return { replaced: rows.length };
}

/** 增量 upsert */
function _upsertRows(name, headers, items, idField, jsonFields) {
  var sh = _ensureExactSheet(name, headers);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { _logError('_upsertRows-lock', e.toString()); }
  try {
    var existing = _getAllObjects(name, headers, jsonFields);
    var index = {};
    for (var i = 0; i < existing.length; i++) index[String(existing[i][idField])] = i + 2;

    var toAppend = [];
    for (var j = 0; j < (items || []).length; j++) {
      var it = items[j];
      var id = String(it[idField] || '');
      if (!id) { id = String(new Date().getTime()) + '-' + j; it[idField] = id; }
      if (index[id]) {
        var row = index[id];
        var vals = headers.map(function(key){ return _valForCell(key, (it[key] !== undefined ? it[key] : '')); });
        sh.getRange(row, 1, 1, headers.length).setValues([vals]);
      } else {
        var valsNew = headers.map(function(key){ return _valForCell(key, (it[key] !== undefined ? it[key] : '')); });
        toAppend.push(valsNew);
      }
    }
    if (toAppend.length) {
      var BATCH = 200, start = 0;
      while (start < toAppend.length) {
        var part = toAppend.slice(start, start + BATCH);
        sh.getRange(sh.getLastRow() + 1, 1, part.length, headers.length).setValues(part);
        start += BATCH;
      }
    }
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove(name); }
    catch (e) { _logError('_upsertRows-cacheRemove', e.toString()); }
    return { status: 'ok' };
  } catch (e) {
    _logError('_upsertRows', e.toString());
    return { status: 'error', message: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** 僅追加新 id（ScoreHistory/QuizAnswers 用） */
function _appendNewById(name, headers, items, jsonFields) {
  var list = _getAllObjects(name, headers, jsonFields);
  var exist = {}; for (var i = 0; i < list.length; i++) exist[String(list[i].id)] = true;
  var toAdd = [];
  for (var j = 0; j < (items || []).length; j++) { var it = items[j]; if (!exist[String(it.id)]) toAdd.push(it); }
  if (!toAdd.length) return { appended: 0 };

  var rows = toAdd.map(function(item){
    if (!item.id) item.id = String(new Date().getTime());
    return H[ name ].map(function(h){ return _valForCell(h, (item[h] !== undefined ? item[h] : '')); });
  });

  var sh = _ensureExactSheet(name, H[name]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, H[name].length).setValues(rows);
  SpreadsheetApp.flush();
  try { CacheService.getScriptCache().remove(name); }
  catch (e) { _logError('_appendNewById-cacheRemove', e.toString()); }
  _trimHistoryIfNeeded(name, sh);
  return { appended: rows.length };
}

function _trimHistoryIfNeeded(name, sh) {
  var limits = { 'ScoreHistory': 10000, 'QuizAnswers': 5000 };
  if (!(name in limits)) return;
  var max = limits[name], last = sh.getLastRow();
  if (last - 1 > max) {
    var removeCount = (last - 1) - max;
    sh.deleteRows(2, removeCount);
    _logError('_trimHistoryIfNeeded', name + ' trimmed ' + removeCount + ' rows');
  }
}

var H = {
  Students:        ['id','name','groupId','groupName','score'],
  Groups:          ['id','name','score'],
  Rewards:         ['id','name','points','quantity','image','description'],
  ExchangeHistory: ['id','studentId','rewardId','status','date'],
  ScoreHistory:    ['id','type','targetId','targetName','groupName','scoreChange','reason','date','affectedStudents'],
  Announcements:   ['id','title','content','link','date'],
  Quizzes:         ['id','winners','title','question','choiceA','choiceB','choiceC','choiceD','correct','startType','startTime','status'],
  QuizAnswers:     ['id','quizId','studentId','studentName','rank','scoreAwarded','answer','isCorrect','submitTime'],
  Settings:        ['key','value','updatedAt']
};

function getStorage(key) {
  switch (String(key)) {
    case 'students':         return _getAllObjects('Students', H.Students);
    case 'groups':           return _getAllObjects('Groups', H.Groups);
    case 'rewards':          return _getAllObjects('Rewards', H.Rewards);
    case 'exchangeRequests': return _getAllObjects('ExchangeHistory', H.ExchangeHistory);
    case 'scoreHistory':     return _getAllObjects('ScoreHistory', H.ScoreHistory, ['affectedStudents']);
    case 'announcements':    return _getAllObjects('Announcements', H.Announcements);
    case 'quizzes':          return _getAllObjects('Quizzes', H.Quizzes, ['winners']); // 這裡會自動補 choices/answer
    case 'quizAnswers':      return _getAllObjects('QuizAnswers', H.QuizAnswers);
    case 'classTitle':       return _getSetting('classTitle');
    case 'loginAttempts':    return _getSetting('loginAttempts');
    case 'lockoutTime':      return _getSetting('lockoutTime');
    default:                 return _getSetting(String(key));
  }
}

function setStorage(key, value) {
  switch (String(key)) {
    case 'students':         return _upsertRows('Students', H.Students, value || [], 'id');
    case 'groups':           return _upsertRows('Groups', H.Groups, value || [], 'id');

    // ★獎品：整表覆寫（含刪除），解決「前端刪了但 Sheet 沒刪」的問題
    case 'rewards': {
      var items = (value || []).map(function(r){
        return {
          id: r.id, name: r.name, points: r.points, quantity: r.quantity,
          image: r.image || '', description: r.description || ''
        };
      });
      return _replaceAllRows('Rewards', H.Rewards, items);
    }

    case 'announcements':    return _upsertRows('Announcements', H.Announcements, value || [], 'id');

    // ★quizzes：同時支援舊/新欄位；寫入時自動展開 choiceA~D / correct / startTime
    case 'quizzes': {
      var itemsQ = (value || []).map(function(q){
        var out = Object.assign({}, q);
        if (q.choices) {
          out.choiceA = q.choices.A || '';
          out.choiceB = q.choices.B || '';
          out.choiceC = q.choices.C || '';
          out.choiceD = q.choices.D || '';
        } else {
          out.choiceA = out.choiceA || '';
          out.choiceB = out.choiceB || '';
          out.choiceC = out.choiceC || '';
          out.choiceD = out.choiceD || '';
        }
        if (out.correct === undefined && q.answer !== undefined) out.correct = q.answer;
        if (!out.startTime && out.startType === 'immediate') out.startTime = _nowISO();
        if (!out.status) out.status = 'active';
        return out;
      });
      return _upsertRows('Quizzes', H.Quizzes, itemsQ, 'id');
    }

    case 'scoreHistory':     return _appendNewById('ScoreHistory', H.ScoreHistory, value || [], ['affectedStudents']);
    case 'quizAnswers':      return _appendNewById('QuizAnswers', H.QuizAnswers, value || []);
    case 'exchangeRequests': return _upsertRows('ExchangeHistory', H.ExchangeHistory, value || [], 'id');
    default:                 return _setSetting(String(key), value);
  }
}

function _getSetting(key) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rng = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rng.length; i++) if (String(rng[i][0]) === String(key)) return rng[i][1];
  return null;
}

function _setSetting(key, value) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  for (var r = 2; r <= last; r++) {
    if (String(sh.getRange(r, 1).getValue()) === String(key)) {
      sh.getRange(r, 2, 1, 2).setValues([[value, _nowISO()]]);
      return { status: 'updated' };
    }
  }
  sh.appendRow([key, value, _nowISO()]);
  return { status: 'created' };
}

function saveAllData(data) {
  try {
    if (data && data.students !== undefined)         setStorage('students', data.students);
    if (data && data.rewards  !== undefined)         setStorage('rewards',  data.rewards);
    if (data && data.history  !== undefined)         setStorage('scoreHistory', data.history);
    if (data && data.groups   !== undefined)         setStorage('groups',   data.groups);
    if (data && data.exchangeRequests !== undefined) setStorage('exchangeRequests', data.exchangeRequests);
    if (data && data.quizzes  !== undefined)         setStorage('quizzes',  data.quizzes);
    return { success: true };
  } catch (e) {
    _logError('saveAllData', e.toString());
    return { success: false, error: e.toString() };
  }
}

/** 真刪：刪對應 id 的資料列 */
function deleteStorage(key, id) {
  var sh, headers;
  switch (String(key)) {
    case 'students':         sh = _ensureExactSheet('Students', H.Students); headers = H.Students; break;
    case 'groups':           sh = _ensureExactSheet('Groups', H.Groups); headers = H.Groups; break;
    case 'rewards':          sh = _ensureExactSheet('Rewards', H.Rewards); headers = H.Rewards; break;
    case 'announcements':    sh = _ensureExactSheet('Announcements', H.Announcements); headers = H.Announcements; break;
    case 'quizzes':          sh = _ensureExactSheet('Quizzes', H.Quizzes); headers = H.Quizzes; break;
    case 'exchangeRequests': sh = _ensureExactSheet('ExchangeHistory', H.ExchangeHistory); headers = H.ExchangeHistory; break;
    case 'scoreHistory':     sh = _ensureExactSheet('ScoreHistory', H.ScoreHistory); headers = H.ScoreHistory; break;
    case 'quizAnswers':      sh = _ensureExactSheet('QuizAnswers', H.QuizAnswers); headers = H.QuizAnswers; break;
    default:                 return { status: 'error', message: 'Unsupported key: ' + key };
  }
  var last = sh.getLastRow();
  if (last < 2) return { status: 'notfound' };
  var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      try { CacheService.getScriptCache().remove(key); }
      catch (e) { _logError('deleteStorage-cacheRemove', e.toString()); }
      return { status: 'deleted' };
    }
  }
  return { status: 'notfound' };
}
