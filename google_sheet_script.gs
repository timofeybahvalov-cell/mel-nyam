/**
 * «Мел-Ням» — онлайн-таблица рекордов и уровни игроков.
 *
 * Этот файл целиком вставляется в редактор Apps Script, открытый из Google Таблицы
 * (меню «Расширения» → «Apps Script»), и публикуется как веб-приложение
 * («Начать развёртывание» → «Новое развёртывание» → тип «Веб-приложение»,
 *  «Запуск от имени: Я», «У кого есть доступ: Все»).
 * Листы «records» и «levels» скрипт создаёт сам при первом обращении.
 *
 * Модерация — прямо в таблице:
 *   • чтобы убрать чужой рекорд, удали его строку на листе «records»;
 *   • чтобы скрыть уровень игрока, впиши любой символ в столбец «hidden» на листе «levels».
 */
var SALT = 'mel-nyam/v2/4eKywka';   // тот же, что в игре (src/g_online.inc)
var TOP_N = 100;                    // сколько рекордов хранится на один уровень
var PAGE = 8;                       // уровней игроков на одной странице списка
var LEVELS_PER_DAY = 20;            // сколько уровней один игрок может опубликовать за сутки
var MAX_LEVELS = 5000;              // всего уровней игроков
var EXAM_FROM = 54;                 // встроенные уровни с этим номером и дальше — экзамены (10 звёзд)

// Корни слов, которые нельзя писать в нике и в названии уровня (ник заменяется на «Игрок», уровень не принимается).
var BAD = ['хуй', 'хуе', 'хуи', 'хуя', 'хую', 'пизд', 'ебат', 'ебан', 'ебуч', 'ебля', 'ебло', 'ебись', 'заеб', 'уеб', 'наеб', 'выеб', 'бляд', 'блят', 'сука', 'сучк',
           'мудак', 'мудил', 'гандон', 'гондон', 'пидор', 'пидар', 'педик', 'шлюх', 'говн', 'залуп', 'дроч', 'нигер', 'ниггер',
           'fuck', 'shit', 'bitch', 'cunt', 'nigg', 'pidor', 'pizd', 'blya', 'suka', 'ebat', 'whore', 'porn'];

function doGet(e) {
  var out;
  try { out = handleGet_((e && e.parameter) || {}); } catch (err) { out = 'err\tserver: ' + err; }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  var out;
  try { out = handlePost_(String((e && e.postData && e.postData.contents) || '')); } catch (err) { out = 'err\tserver: ' + err; }
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.TEXT);
}

function handleGet_(p) {
  var a = String(p.a || '');
  if (a === 'top') return top_(String(p.l || ''));
  if (a === 'rep') return rep_(String(p.l || ''), String(p.p || ''));
  if (a === 'list') return list_(String(p.s || 'new'), parseInt(p.o, 10) || 0);
  if (a === 'lvl') return lvl_(parseInt(p.id, 10) || 0);
  if (a === 'ping' || a === '') return 'ok\tmel-nyam\t2';
  return 'err\tunknown';
}

function handlePost_(body) {
  var f = body.replace(/[\r\n]+$/, '').split('\t');
  if (f[0] === 'sub') return sub_(f);
  if (f[0] === 'pub') return pub_(f);
  return 'err\tunknown';
}

/* ---------------------------------------------------------------- листы */
function sheet_(name, header, textCols) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    for (var i = 0; i < textCols.length; i++) sh.getRange(textCols[i]).setNumberFormat('@');   // текст остаётся текстом («007», «1-2»)
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
function recSheet_() { return sheet_('records', ['level', 'uid', 'nick', 'stars', 'ticks', 'replay', 'date'], ['A:C', 'F:F']); }
function lvlSheet_() { return sheet_('levels', ['id', 'uid', 'author', 'name', 'stars', 'bg', 'plays', 'date', 'hidden', 'data', 'replay'], ['B:D', 'I:K']); }
function rows_(sh, cols) {
  var n = sh.getLastRow() - 1;
  return n < 1 ? [] : sh.getRange(2, 1, n, cols).getValues();
}

/* ---------------------------------------------------------------- подпись и проверки */
function utf8_(s) {
  var out = [];
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c >= 0xD800 && c < 0xDC00 && i + 1 < s.length) { c = 0x10000 + ((c - 0xD800) << 10) + (s.charCodeAt(++i) - 0xDC00); }
    if (c < 0x80) out.push(c);
    else if (c < 0x800) out.push(0xC0 | (c >> 6), 0x80 | (c & 63));
    else if (c < 0x10000) out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
    else out.push(0xF0 | (c >> 18), 0x80 | ((c >> 12) & 63), 0x80 | ((c >> 6) & 63), 0x80 | (c & 63));
  }
  return out;
}
function fnv_(h, s) {
  var b = utf8_(s);
  for (var i = 0; i < b.length; i++) {
    h = (h ^ b[i]) >>> 0;
    h = (((h & 0xFFFF) * 16777619) + ((((h >>> 16) * 16777619) & 0xFFFF) * 65536)) % 4294967296;      // h * 16777619 по модулю 2^32
  }
  return h;
}
function sig_(parts) {               // всегда пять полей, как в игре
  var h = fnv_(2166136261, SALT);
  for (var i = 0; i < 5; i++) { h = fnv_(h, '|'); h = fnv_(h, parts[i] || ''); }
  return ('0000000' + h.toString(16)).slice(-8);
}
function pid_(uid, level) { return sig_(['p', String(uid), level, '', '']); }      // открытый номер записи: по нему нельзя узнать uid
function isBad_(s) {
  var t = String(s).toLowerCase().replace(/ё/g, 'е').replace(/[_\-\s.!?,]/g, '');
  t = t.replace(/0/g, 'о').replace(/3/g, 'з').replace(/1/g, 'и').replace(/4/g, 'ч').replace(/6/g, 'б');
  for (var i = 0; i < BAD.length; i++) if (t.indexOf(BAD[i]) >= 0) return true;
  return false;
}
function cleanNick_(s) {
  s = String(s || '');
  if (!/^[0-9A-Za-zА-Яа-яЁё_\-]{2,12}$/.test(s) || isBad_(s)) return 'Игрок';
  return s;
}
function cleanName_(s) {
  s = String(s || '').replace(/\s+/g, ' ').replace(/^ | $/g, '');
  if (!/^[0-9A-Za-zА-Яа-яЁё_\- !?.,]{2,20}$/.test(s) || isBad_(s)) return null;
  return s;
}
function replayInfo_(rep) {          // «тики_звёзды_шестнадцатеричные действия»
  var m = /^(\d{1,6})_(\d{1,2})_((?:[0-9a-f]{8}){0,1500})$/.exec(String(rep));
  if (!m) return null;
  var ticks = parseInt(m[1], 10), stars = parseInt(m[2], 10), last = 0;
  if (ticks < 1 || ticks > 131071 || stars > 10) return null;
  for (var i = 0; i < m[3].length; i += 8) {
    var t = parseInt(m[3].substr(i, 8), 16) >>> 15;
    if (t < last || t > ticks) return null;
    last = t;
  }
  return { ticks: ticks, stars: stars };
}
function better_(s1, t1, s2, t2) { return s1 > s2 || (s1 === s2 && t1 < t2); }
function levelOk_(level) { return /^(b\d\d-[0-9a-f]{4}|u\d{1,9})$/.test(level); }
function findLevelRow_(sh, id) {     // номер строки уровня игрока на листе или 0
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var ids = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) if (Number(ids[i][0]) === id) return i + 2;
  return 0;
}

/* ---------------------------------------------------------------- рекорды */
function levelRecs_(all, level) {
  var out = [];
  for (var i = 0; i < all.length; i++) {
    if (String(all[i][0]) !== level) continue;
    out.push({ row: i + 2, uid: String(all[i][1]), nick: String(all[i][2]), stars: Number(all[i][3]) || 0, ticks: Number(all[i][4]) || 0 });
  }
  out.sort(function (a, b) { return b.stars - a.stars || a.ticks - b.ticks || a.row - b.row; });
  return out;
}
function top_(level) {
  if (!levelOk_(level)) return 'err\tlevel';
  var cache = CacheService.getScriptCache(), key = 'top:' + level, hit = cache.get(key);
  if (hit) return hit;
  var r = levelRecs_(rows_(recSheet_(), 5), level), n = Math.min(r.length, TOP_N), out = ['ok\t' + n];
  for (var i = 0; i < n; i++) out.push((i + 1) + '\t' + cleanNick_(r[i].nick) + '\t' + r[i].stars + '\t' + r[i].ticks + '\t' + pid_(r[i].uid, level));
  var text = out.join('\n');
  cache.put(key, text, 60);
  return text;
}
function rep_(level, p) {
  if (!levelOk_(level) || !/^[0-9a-f]{8}$/.test(p)) return 'err\tformat';
  var sh = recSheet_(), r = levelRecs_(rows_(sh, 5), level);
  for (var i = 0; i < r.length; i++) {
    if (pid_(r[i].uid, level) !== p) continue;
    return 'ok\t' + cleanNick_(r[i].nick) + '\t' + String(sh.getRange(r[i].row, 6).getValue());
  }
  return 'err\tnot found';
}
function sub_(f) {                   // sub, уровень, uid, ник, звёзды, тики, запись, подпись
  if (f.length < 8) return 'err\tformat';
  var level = f[1], uid = f[2], nick = cleanNick_(f[3]), stars = parseInt(f[4], 10), ticks = parseInt(f[5], 10), rep = f[6];
  if (!levelOk_(level) || !/^\d{1,10}$/.test(uid)) return 'err\tformat';
  if (sig_([level, uid, f[3], rep, '']) !== f[7]) return 'err\tsig';
  var info = replayInfo_(rep);
  if (!info || info.ticks !== ticks || info.stars !== stars) return 'err\treplay';
  var maxStars = 3;
  if (level.charAt(0) === 'b') { if (parseInt(level.substr(1, 2), 10) >= EXAM_FROM) maxStars = 10; }
  else {
    var ls = lvlSheet_(), lr = findLevelRow_(ls, parseInt(level.substr(1), 10));
    if (!lr) return 'err\tno level';
    maxStars = Number(ls.getRange(lr, 5).getValue()) || 0;
  }
  if (stars > maxStars) return 'err\tstars';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recSheet_(), r = levelRecs_(rows_(sh, 5), level), mine = null, now = new Date(), i;
    for (i = 0; i < r.length; i++) if (r[i].uid === uid) { mine = r[i]; break; }
    if (mine) {
      if (better_(stars, ticks, mine.stars, mine.ticks)) sh.getRange(mine.row, 3, 1, 5).setValues([[nick, stars, ticks, rep, now]]);
      else { if (mine.nick !== nick) sh.getRange(mine.row, 3).setValue(nick); stars = mine.stars; ticks = mine.ticks; }
    } else if (r.length < TOP_N) {
      sh.appendRow([level, uid, nick, stars, ticks, rep, now]);
    } else {
      var worst = r[r.length - 1];
      if (!better_(stars, ticks, worst.stars, worst.ticks)) return 'ok\t0';          // не попал в первую сотню
      sh.getRange(worst.row, 1, 1, 7).setValues([[level, uid, nick, stars, ticks, rep, now]]);
    }
    CacheService.getScriptCache().remove('top:' + level);
    var rank = 1;
    for (i = 0; i < r.length; i++) if (r[i].uid !== uid && better_(r[i].stars, r[i].ticks, stars, ticks)) rank++;
    return 'ok\t' + rank;
  } finally { lock.releaseLock(); }
}

/* ---------------------------------------------------------------- уровни игроков */
function list_(sort, off) {
  sort = sort === 'top' ? 'top' : 'new';
  if (off < 0) off = 0;
  var cache = CacheService.getScriptCache(), key = 'list:' + sort + ':' + off, hit = cache.get(key);
  if (hit) return hit;
  var all = rows_(lvlSheet_(), 9), r = [];
  for (var i = 0; i < all.length; i++) {
    if (String(all[i][8]) !== '') continue;                                           // скрыт владельцем таблицы
    r.push({ id: Number(all[i][0]), author: cleanNick_(all[i][2]), name: String(all[i][3]), stars: Number(all[i][4]) || 0, bg: Number(all[i][5]) || 0, plays: Number(all[i][6]) || 0 });
  }
  if (sort === 'top') r.sort(function (a, b) { return b.plays - a.plays || b.id - a.id; });
  else r.sort(function (a, b) { return b.id - a.id; });
  var out = ['ok\t' + r.length];
  for (var k = off; k < r.length && k < off + PAGE; k++) out.push(r[k].id + '\t' + r[k].name + '\t' + r[k].author + '\t' + r[k].stars + '\t' + r[k].plays + '\t' + r[k].bg);
  var text = out.join('\n');
  cache.put(key, text, 30);
  return text;
}
function lvl_(id) {
  if (id < 1) return 'err\tformat';
  var sh = lvlSheet_(), row = findLevelRow_(sh, id);
  if (!row) return 'err\tnot found';
  var v = sh.getRange(row, 1, 1, 11).getValues()[0];
  if (String(v[8]) !== '') return 'err\tnot found';
  sh.getRange(row, 7).setValue((Number(v[6]) || 0) + 1);                              // счётчик «сыграно»
  return 'ok\t' + id + '\t' + String(v[3]) + '\t' + cleanNick_(v[2]) + '\t' + String(v[9]) + '\t' + String(v[10]);
}
function pub_(f) {                   // pub, uid, ник, название, уровень, запись прохождения, подпись
  if (f.length < 7) return 'err\tformat';
  var uid = f[1], nick = cleanNick_(f[2]), name = cleanName_(f[3]), data = f[4], rep = f[5];
  if (!/^\d{1,10}$/.test(uid)) return 'err\tformat';
  if (sig_(['pub', uid, f[3], data, rep]) !== f[6]) return 'err\tsig';
  if (name === null) return 'err\tname';
  if (data.length > 4000 || !/^1;\d{1,2}(;\d{1,2}(,-?\d{1,6}){2,7}){2,48}$/.test(data)) return 'err\tdata';
  var info = replayInfo_(rep);
  if (!info) return 'err\treplay';
  var objs = data.split(';').slice(2), count = {}, i;
  for (i = 0; i < objs.length; i++) { var t = parseInt(objs[i], 10); count[t] = (count[t] || 0) + 1; }
  if (count[1] !== 1 || count[8] !== 1 || (count[4] || 0) > 10 || info.stars > (count[4] || 0)) return 'err\tdata';
  var bg = parseInt(data.split(';')[1], 10) % 10;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = lvlSheet_(), all = rows_(sh, 9), now = new Date(), today = 0, maxId = 0;
    if (all.length >= MAX_LEVELS) return 'err\tfull';
    for (i = 0; i < all.length; i++) {
      maxId = Math.max(maxId, Number(all[i][0]) || 0);
      if (String(all[i][1]) === uid && all[i][7] && now.getTime() - new Date(all[i][7]).getTime() < 86400000) today++;
    }
    if (today >= LEVELS_PER_DAY) return 'err\tlimit';
    var id = maxId + 1;
    sh.appendRow([id, uid, nick, name, count[4] || 0, bg, 0, now, '', data, rep]);
    recSheet_().appendRow(['u' + id, uid, nick, info.stars, info.ticks, rep, now]);   // автор — первый в таблице своего уровня
    var cache = CacheService.getScriptCache();
    cache.remove('list:new:0'); cache.remove('list:top:0');
    return 'ok\t' + id;
  } finally { lock.releaseLock(); }
}
