
/**
 * Apps Script 後端 (Code.gs) — IMPROVED + DELETE
 * 功能：
 *  - 分表制：Students / Groups / Rewards / ScoreHistory / Quizzes / QuizAnswers / ExchangeHistory / Announcements / Settings / Logs
 *  - 增量寫入（只更新變動）
 *  - 快取（CacheService, 60s）
 *  - 錯誤紀錄（Logs 表）
 *  - 歷史表容量管理（ScoreHistory 10000, QuizAnswers 5000）
 *  - deleteStorage(key, id)：真正刪除指定 id 的資料列
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
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    return sh;
  }
  var cur = sh.getRange(1,1,1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
  var ok = true;
  for (var i=0;i<headers.length;i++) if (String(cur[i]||'') !== String(headers[i])) { ok = false; break; }
  if (!ok) {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}

function _getAllObjects(name, headers, jsonFields) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(name);
    if (cached) return JSON.parse(cached);
  } catch(e) {
    _logError('_getAllObjects', e.toString());
  }

  var sh = _ensureExactSheet(name, headers);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rng = sh.getRange(2, 1, last-1, headers.length).getValues();
  var arr = rng.map(function(row){
    var o = {};
    for (var i=0;i<headers.length;i++) {
      var k = headers[i], v = row[i];
      if (jsonFields && jsonFields.indexOf(k) >= 0 && typeof v === 'string' && v) {
        try { o[k] = JSON.parse(v); } catch(e) { o[k] = v; }
      } else o[k] = v;
    }
    return o;
  });

  try {
    CacheService.getScriptCache().put(name, JSON.stringify(arr), 60); // cache 60s
  } catch(e) {
    _logError('_getAllObjects-cachePut', e.toString());
  }
  return arr;
}

function _valForCell(field, value) {
  if ((field === 'winners' || field === 'affectedStudents') && typeof value !== 'string') {
    try { return JSON.stringify(value || []); } catch(e) { return '[]'; }
  }
  return value;
}

function _upsertRows(name, headers, items, idField, jsonFields) {
  var sh = _ensureExactSheet(name, headers);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var existing = _getAllObjects(name, headers, jsonFields);
    var index = {};
    for (var i=0;i<existing.length;i++) index[String(existing[i][idField])] = i+2;

    var toAppend = [];
    for (var j=0;j<(items||[]).length;j++) {
      var it = items[j];
      var id = String(it[idField] || '');
      if (!id) { id = String(Date.now()) + '-' + j; it[idField] = id; }
      if (index[id]) {
        var row = index[id];
        var old = existing.find(function(x){ return String(x[idField])===id; }) || {};
        var changed = false;
        for (var k in it) {
          if (it[k] != old[k]) { changed = true; break; }
        }
        if (changed) {
          var vals = headers.map(function(h){ return _valForCell(h, it[h] !== undefined ? it[h] : ''); });
          sh.getRange(row, 1, 1, headers.length).setValues([vals]);
        }
      } else {
        var valsNew = headers.map(function(h){ return _valForCell(h, it[h] !== undefined ? it[h] : ''); });
        toAppend.push(valsNew);
      }
    }
    if (toAppend.length) {
      var BATCH=200, start=0;
      while (start < toAppend.length) {
        var part = toAppend.slice(start, start+BATCH);
        sh.getRange(sh.getLastRow()+1, 1, part.length, headers.length).setValues(part);
        start += BATCH;
      }
    }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(name);
    return {status:'ok'};
  } catch(e) {
    _logError('_upsertRows', e.toString());
    return {status:'error', message:e.toString()};
  } finally {
    lock.releaseLock();
  }
}

function _appendNewById(name, headers, items, jsonFields) {
  var list = _getAllObjects(name, headers, jsonFields);
  var exist = {};
  for (var i=0;i<list.length;i++) exist[String(list[i].id)] = true;
  var toAdd = [];
  for (var j=0;j<(items||[]).length;j++) {
    var it = items[j];
    if (!exist[String(it.id)]) toAdd.push(it);
  }
  if (!toAdd.length) return {appended:0};
  var rows = toAdd.map(function(it){
    if (!it.id) it.id = String(Date.now());
    return headers.map(function(h){ return _valForCell(h, it[h] !== undefined ? it[h] : ''); });
  });
  var sh = _ensureExactSheet(name, headers);
  sh.getRange(sh.getLastRow()+1, 1, rows.length, headers.length).setValues(rows);
  SpreadsheetApp.flush();
  CacheService.getScriptCache().remove(name);
  _trimHistoryIfNeeded(name, sh);
  return {appended: rows.length};
}

// 表格容量管理
function _trimHistoryIfNeeded(name, sh) {
  var limits = { 'ScoreHistory':10000, 'QuizAnswers':5000 };
  if (!(name in limits)) return;
  var max = limits[name];
  var last = sh.getLastRow();
  if (last-1 > max) {
    var removeCount = (last-1) - max;
    sh.deleteRows(2, removeCount);
    _logError('_trimHistoryIfNeeded', name+' trimmed '+removeCount+' rows');
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
  switch(String(key)) {
    case 'students': return _getAllObjects('Students', H.Students);
    case 'groups': return _getAllObjects('Groups', H.Groups);
    case 'rewards': return _getAllObjects('Rewards', H.Rewards);
    case 'exchangeRequests': return _getAllObjects('ExchangeHistory', H.ExchangeHistory);
    case 'scoreHistory': return _getAllObjects('ScoreHistory', H.ScoreHistory, ['affectedStudents']);
    case 'announcements': return _getAllObjects('Announcements', H.Announcements);
    case 'quizzes': return _getAllObjects('Quizzes', H.Quizzes, ['winners']);
    case 'quizAnswers': return _getAllObjects('QuizAnswers', H.QuizAnswers);
    case 'classTitle': return _getSetting('classTitle');
    case 'loginAttempts': return _getSetting('loginAttempts');
    case 'lockoutTime': return _getSetting('lockoutTime');
    default: return _getSetting(String(key));
  }
}

function setStorage(key, value) {
  switch(String(key)) {
    case 'students': return _upsertRows('Students', H.Students, value||[], 'id');
    case 'groups': return _upsertRows('Groups', H.Groups, value||[], 'id');
    case 'rewards': return _upsertRows('Rewards', H.Rewards, value||[], 'id');
    case 'announcements': return _upsertRows('Announcements', H.Announcements, value||[], 'id');
    case 'quizzes': return _upsertRows('Quizzes', H.Quizzes, value||[], 'id');
    case 'scoreHistory': return _appendNewById('ScoreHistory', H.ScoreHistory, value||[], ['affectedStudents']);
    case 'quizAnswers': return _appendNewById('QuizAnswers', H.QuizAnswers, value||[]);
    case 'exchangeRequests': return _upsertRows('ExchangeHistory', H.ExchangeHistory, value||[], 'id');
    default: return _setSetting(String(key), value);
  }
}

function _getSetting(key) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  if (last < 2) return null;
  var rng = sh.getRange(2,1,last-1,3).getValues();
  for (var i=0;i<rng.length;i++) if (String(rng[i][0])===String(key)) return rng[i][1];
  return null;
}

function _setSetting(key, value) {
  var sh = _ensureExactSheet('Settings', H.Settings);
  var last = sh.getLastRow();
  for (var r=2;r<=last;r++) {
    if (String(sh.getRange(r,1).getValue()) === String(key)) {
      sh.getRange(r,2,1,2).setValues([[value, _nowISO()]]);
      return {status:'updated'};
    }
  }
  sh.appendRow([key,value,_nowISO()]);
  return {status:'created'};
}

function saveAllData(data) {
  try {
    if (data && data.students !== undefined) setStorage('students', data.students);
    if (data && data.rewards !== undefined) setStorage('rewards', data.rewards);
    if (data && data.history !== undefined) setStorage('scoreHistory', data.history);
    if (data && data.groups !== undefined) setStorage('groups', data.groups);
    if (data && data.exchangeRequests !== undefined) setStorage('exchangeRequests', data.exchangeRequests);
    if (data && data.quizzes !== undefined) setStorage('quizzes', data.quizzes);
    return {success:true};
  } catch(e) {
    _logError('saveAllData', e.toString());
    return {success:false, error:e.toString()};
  }
}

/** 刪除指定 key 的 id 資料列（真刪） */
function deleteStorage(key, id) {
  var sh, headers;
  switch(String(key)) {
    case 'students':        sh = _ensureExactSheet('Students', H.Students); headers = H.Students; break;
    case 'groups':          sh = _ensureExactSheet('Groups', H.Groups); headers = H.Groups; break;
    case 'rewards':         sh = _ensureExactSheet('Rewards', H.Rewards); headers = H.Rewards; break;
    case 'announcements':   sh = _ensureExactSheet('Announcements', H.Announcements); headers = H.Announcements; break;
    case 'quizzes':         sh = _ensureExactSheet('Quizzes', H.Quizzes); headers = H.Quizzes; break;
    case 'exchangeRequests':sh = _ensureExactSheet('ExchangeHistory', H.ExchangeHistory); headers = H.ExchangeHistory; break;
    case 'scoreHistory':    sh = _ensureExactSheet('ScoreHistory', H.ScoreHistory); headers = H.ScoreHistory; break;
    case 'quizAnswers':     sh = _ensureExactSheet('QuizAnswers', H.QuizAnswers); headers = H.QuizAnswers; break;
    default: return {status:'error', message:'Unsupported key: '+key};
  }

  var last = sh.getLastRow();
  if (last < 2) return {status:'notfound'};

  var values = sh.getRange(2,1,last-1,headers.length).getValues();
  for (var i=0;i<values.length;i++) {
    if (String(values[i][0]) === String(id)) { // 第一欄為 id
      sh.deleteRow(i+2);
      CacheService.getScriptCache().remove(key);
      return {status:'deleted'};
    }
  }
  return {status:'notfound'};
}
