/**
 * «Мел-Ням» — онлайн-таблица рекордов и уровни игроков.
 *
 * Этот файл целиком вставляется в редактор Apps Script, открытый из Google Таблицы
 * (меню «Расширения» → «Apps Script»), и публикуется как веб-приложение
 * («Начать развёртывание» → «Новое развёртывание» → тип «Веб-приложение»,
 *  «Запуск от имени: Я», «У кого есть доступ: Все»).
 * Листы «records», «levels» и «users» скрипт создаёт сам при первом обращении.
 * Как всё это сделать по шагам (и как перейти с версии 0.3 на 0.4) — в файле ONLINE_SETUP.md.
 *
 * С версии 0.4 здесь же: аккаунты игроков (ник + пароль), список «кто в сети» и команды админки.
 *   • Пароли игроков в таблицу не попадают: игра присылает только их SHA-256, а на листе «users»
 *     хранится хеш от этого хеша с «солью». Узнать пароль по таблице нельзя.
 *   • Админка (MelNyamAdmin.exe) работает только с паролем админки. Его создаёт функция
 *     melAdminPassword(): выбери её в редакторе Apps Script и нажми «Выполнить» — пароль появится
 *     в журнале выполнения. В таблице и в этом файле пароля нет, поэтому файл можно публиковать.
 *
 * Модерация: программой MelNyamAdmin.exe или прямо в таблице:
 *   • чтобы убрать чужой рекорд, удали его строку на листе «records»;
 *   • чтобы скрыть уровень игрока, впиши любой символ в столбец «hidden» на листе «levels»;
 *   • чтобы заблокировать игрока, впиши любой символ в столбец «banned» на листе «users».
 */
var SALT = 'mel-nyam/v2/4eKywka';   // тот же, что в игре (src/g_online.inc)
var PROTO = 4;                      // что умеет этот скрипт: 2 — рекорды и уровни (V0.2, V0.3), 4 — ещё и аккаунты (V0.4)
var MAX_TOKENS = 5;                 // на скольких устройствах можно быть в аккаунте одновременно
var REGS_PER_DAY = 300;             // сколько аккаунтов можно создать за сутки (защита от ботов)
var ONLINE_SEC = 150;               // игрок считается «в сети», если его игра откликалась не раньше, чем столько секунд назад
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
  if (a === 'hb') return hb_(p);
  if (a === 'mine') return mine_(String(p.u || ''), String(p.t || ''));
  if (a === 'time') return 'ok\t' + Math.floor(Date.now() / 1000);
  if (a === 'ping' || a === '') return 'ok\tmel-nyam\t' + PROTO;
  return 'err\tunknown';
}

function handlePost_(body) {
  var f = body.replace(/[\r\n]+$/, '').split('\t');
  if (f[0] === 'sub') return sub_(f, false);       // V0.2, V0.3: без аккаунта
  if (f[0] === 'sub2') return sub_(f, true);       // V0.4: от имени аккаунта
  if (f[0] === 'pub') return pub_(f, false);
  if (f[0] === 'pub2') return pub_(f, true);
  if (f[0] === 'reg') return reg_(f);
  if (f[0] === 'login') return login_(f);
  if (f[0] === 'ren') return ren_(f);
  if (f[0] === 'pw') return pw_(f);
  if (f[0] === 'out') return out_(f);
  if (f[0] === 'adm') return adm_(f);
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
function userSheet_() {
  return sheet_('users', ['uid', 'nick', 'key', 'salt', 'hash', 'tokens', 'created', 'seen', 'where', 'ver', 'platform', 'banned', 'note'], ['A:F', 'I:I', 'L:M']);
}
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
  var cache = CacheService.getScriptCache(), key = kTop_(level), hit = cache.get(key);
  if (hit) return hit;
  var U = usersSafe_(), all = levelRecs_(rows_(recSheet_(), 5), level), r = [], i;
  for (i = 0; i < all.length && r.length < TOP_N; i++) if (!isBanned_(U, all[i].uid)) r.push(all[i]);       // заблокированных в таблице не видно
  var n = r.length, out = ['ok\t' + n];
  for (i = 0; i < n; i++) out.push((i + 1) + '\t' + showNick_(U, r[i].uid, r[i].nick) + '\t' + r[i].stars + '\t' + r[i].ticks + '\t' + pid_(r[i].uid, level));
  var text = out.join('\n');
  cache.put(key, text, 60);
  return text;
}
function rep_(level, p) {
  if (!levelOk_(level) || !/^[0-9a-f]{8}$/.test(p)) return 'err\tformat';
  var sh = recSheet_(), r = levelRecs_(rows_(sh, 5), level), U = usersSafe_();
  for (var i = 0; i < r.length; i++) {
    if (pid_(r[i].uid, level) !== p) continue;
    if (isBanned_(U, r[i].uid)) break;
    return 'ok\t' + showNick_(U, r[i].uid, r[i].nick) + '\t' + String(sh.getRange(r[i].row, 6).getValue());
  }
  return 'err\tnot found';
}
function sub_(f, v4) {               // sub, уровень, uid, ник, звёзды, тики, запись, подпись   |   sub2: вместо ника — ключ аккаунта
  if (f.length < 8) return 'err\tformat';
  var level = f[1], uid = f[2], nick, stars = parseInt(f[4], 10), ticks = parseInt(f[5], 10), rep = f[6];
  if (!levelOk_(level) || !/^\d{1,10}$/.test(uid)) return 'err\tformat';
  if ((v4 ? sig_(['sub2', level, uid, rep, '']) : sig_([level, uid, f[3], rep, ''])) !== f[7]) return 'err\tsig';
  var U = usersSafe_();
  if (isBanned_(U, uid)) return 'err\tban';
  if (v4) { var acc = auth_(U, uid, f[3]); if (!acc) return 'err\tauth'; nick = acc.nick; }
  else { if (U.byUid[uid] && U.byUid[uid].hash) return 'err\tauth'; nick = legacyNick_(U, f[3]); }      // у аккаунта есть пароль: старая игра от его имени писать не может
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
    CacheService.getScriptCache().remove(kTop_(level));
    var rank = 1;
    for (i = 0; i < r.length; i++) if (r[i].uid !== uid && better_(r[i].stars, r[i].ticks, stars, ticks)) rank++;
    return 'ok\t' + rank;
  } finally { lock.releaseLock(); }
}

/* ---------------------------------------------------------------- уровни игроков */
function list_(sort, off) {
  sort = sort === 'top' ? 'top' : 'new';
  if (off < 0) off = 0;
  var cache = CacheService.getScriptCache(), key = kList_(sort, off), hit = cache.get(key);
  if (hit) return hit;
  var all = rows_(lvlSheet_(), 9), r = [], U = usersSafe_();
  for (var i = 0; i < all.length; i++) {
    if (String(all[i][8]) !== '' || isBanned_(U, String(all[i][1]))) continue;         // скрыт владельцем таблицы или автор заблокирован
    r.push({ id: Number(all[i][0]), author: showNick_(U, String(all[i][1]), all[i][2]), name: String(all[i][3]), stars: Number(all[i][4]) || 0, bg: Number(all[i][5]) || 0, plays: Number(all[i][6]) || 0 });
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
  var v = sh.getRange(row, 1, 1, 11).getValues()[0], U = usersSafe_();
  if (String(v[8]) !== '' || isBanned_(U, String(v[1]))) return 'err\tnot found';
  sh.getRange(row, 7).setValue((Number(v[6]) || 0) + 1);                              // счётчик «сыграно»
  return 'ok\t' + id + '\t' + String(v[3]) + '\t' + showNick_(U, String(v[1]), v[2]) + '\t' + String(v[9]) + '\t' + String(v[10]);
}
function pub_(f, v4) {               // pub, uid, ник, название, уровень, запись прохождения, подпись   |   pub2: вместо ника — ключ аккаунта
  if (f.length < 7) return 'err\tformat';
  var uid = f[1], nick, name = cleanName_(f[3]), data = f[4], rep = f[5];
  if (!/^\d{1,10}$/.test(uid)) return 'err\tformat';
  if (sig_([v4 ? 'pub2' : 'pub', uid, f[3], data, rep]) !== f[6]) return 'err\tsig';
  var U = usersSafe_();
  if (isBanned_(U, uid)) return 'err\tban';
  if (v4) { var acc = auth_(U, uid, f[2]); if (!acc) return 'err\tauth'; nick = acc.nick; }
  else { if (U.byUid[uid] && U.byUid[uid].hash) return 'err\tauth'; nick = legacyNick_(U, f[2]); }
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
    cache.remove(kList_('new', 0)); cache.remove(kList_('top', 0));
    return 'ok\t' + id;
  } finally { lock.releaseLock(); }
}

/* ---------------------------------------------------------------- V0.4: общие мелочи */
function hex_(bytes) {               // байты Apps Script — числа от -128 до 127
  var o = '';
  for (var i = 0; i < bytes.length; i++) { var v = (bytes[i] + 256) % 256; o += (v < 16 ? '0' : '') + v.toString(16); }
  return o;
}
function sha_(s) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)); }
function hmac_(msg, key) { return hex_(Utilities.computeHmacSha256Signature(String(msg), String(key), Utilities.Charset.UTF_8)); }
function rnd_() { return sha_(Utilities.getUuid() + ':' + Utilities.getUuid() + ':' + Date.now() + ':' + Math.random()).substr(0, 32); }
function safeEq_(a, b) {             // сравнение без «подсказок» по времени
  a = String(a); b = String(b);
  var d = a.length ^ b.length;
  for (var i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i % (b.length || 1));
  return d === 0;
}
function ts_(d) { var t = d ? new Date(d).getTime() : NaN; return isNaN(t) ? 0 : Math.floor(t / 1000); }
function nickKey_(s) { return String(s).toLowerCase().replace(/ё/g, 'е'); }       // «Вася» и «вася» — один и тот же ник
function tokHash_(token) { return sha_('t:' + token).substr(0, 32); }              // сами ключи устройств в таблице не лежат
function gen_() { return CacheService.getScriptCache().get('gen') || '0'; }        // «поколение» кеша: меняется, когда меняются ники и блокировки
function bump_() { CacheService.getScriptCache().put('gen', String(Date.now()), 21600); }
function kTop_(level) { return 'top:' + gen_() + ':' + level; }
function kList_(sort, off) { return 'list:' + gen_() + ':' + sort + ':' + off; }

/* ---------------------------------------------------------------- V0.4: игроки */
function users_(sh) {
  var v = rows_(sh, 13), U = { byUid: {}, byKey: {}, list: [] };
  for (var i = 0; i < v.length; i++) {
    var u = { row: i + 2, uid: String(v[i][0]), nick: String(v[i][1]), key: String(v[i][2]), salt: String(v[i][3]), hash: String(v[i][4]),
              tokens: String(v[i][5]) ? String(v[i][5]).split(',') : [], created: v[i][6], seen: v[i][7], where: String(v[i][8]),
              ver: Number(v[i][9]) || 0, pf: Number(v[i][10]) || 0, banned: String(v[i][11]) !== '', note: String(v[i][12]) };
    if (!u.uid) continue;
    U.byUid[u.uid] = u; U.list.push(u);
    if (u.hash && u.key) U.byKey[u.key] = u;
  }
  return U;
}
function usersSafe_() {              // для старых команд: если с листом «users» что-то не так, рекорды всё равно работают
  try { return users_(userSheet_()); } catch (e) { return { byUid: {}, byKey: {}, list: [] }; }
}
function isBanned_(U, uid) { var u = U.byUid[String(uid)]; return !!(u && u.banned); }
function legacyNick_(U, nick) {      // ник игрока без аккаунта: чужой ник с паролем занять нельзя
  var n = cleanNick_(nick);
  return U.byKey[nickKey_(n)] ? 'Игрок' : n;
}
function showNick_(U, uid, rowNick) {
  var u = U.byUid[String(uid)];
  return u && u.hash && u.nick ? cleanNick_(u.nick) : legacyNick_(U, rowNick);
}
function auth_(U, uid, token) {      // игрок с таким номером и таким ключом устройства или null
  var u = U.byUid[String(uid)];
  if (!u || !u.hash || !token || u.banned) return null;
  return u.tokens.indexOf(tokHash_(token)) >= 0 ? u : null;
}
function nickOk_(nick) { return cleanNick_(nick) === nick && nickKey_(nick) !== 'игрок'; }

function reg_(f) {                   // reg, uid, ник, хеш пароля, версия игры, подпись
  if (f.length < 6) return 'err\tformat';
  var uid = f[1], nick = f[2], ph = f[3];
  if (!/^[1-9]\d{0,9}$/.test(uid) || !/^[0-9a-f]{64}$/.test(ph)) return 'err\tformat';
  if (sig_(['reg', uid, nick, ph, '']) !== f[5]) return 'err\tsig';
  if (!nickOk_(nick)) return 'err\tbadnick';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = userSheet_(), U = users_(sh), u = U.byUid[uid], key = nickKey_(nick), now = new Date(), cache = CacheService.getScriptCache();
    if (u && u.banned) return 'err\tban';
    if (u && u.hash) return 'err\texists';
    if (U.byKey[key]) return 'err\ttaken';
    var day = 'regs:' + Math.floor(now.getTime() / 86400000), made = parseInt(cache.get(day), 10) || 0;
    if (made >= REGS_PER_DAY) return 'err\tlimit';
    cache.put(day, String(made + 1), 21600);
    var salt = rnd_().substr(0, 16), token = rnd_();
    var row = [uid, nick, key, salt, sha_(salt + ':' + ph), tokHash_(token), now, now, '', parseInt(f[4], 10) || 0, '', '', ''];
    if (u) sh.getRange(u.row, 1, 1, 13).setValues([row]); else sh.appendRow(row);
    bump_();
    return 'ok\t' + uid + '\t' + nick + '\t' + token;
  } finally { lock.releaseLock(); }
}
function login_(f) {                 // login, ник, хеш пароля, версия игры, подпись
  if (f.length < 5) return 'err\tformat';
  var nick = f[1], ph = f[2];
  if (!/^[0-9a-f]{64}$/.test(ph)) return 'err\tformat';
  if (sig_(['login', nick, ph, '', '']) !== f[4]) return 'err\tsig';
  var key = nickKey_(nick), cache = CacheService.getScriptCache(), fk = 'lf:' + sha_(key).substr(0, 24), fails = parseInt(cache.get(fk), 10) || 0;
  if (fails >= 6) return 'err\twait';                                                 // подбор пароля: шесть промахов — и 10 минут тишины
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = userSheet_(), U = users_(sh), u = U.byKey[key];
    if (!u || !safeEq_(sha_(u.salt + ':' + ph), u.hash)) { cache.put(fk, String(fails + 1), 600); return 'err\tauth'; }
    if (u.banned) return 'err\tban';
    var token = rnd_(), toks = u.tokens.concat([tokHash_(token)]);
    while (toks.length > MAX_TOKENS) toks.shift();
    sh.getRange(u.row, 6).setValue(toks.join(','));
    sh.getRange(u.row, 8).setValue(new Date());
    cache.remove(fk);
    return 'ok\t' + u.uid + '\t' + u.nick + '\t' + token;
  } finally { lock.releaseLock(); }
}
function ren_(f) {                   // ren, uid, ключ, новый ник, подпись
  if (f.length < 5) return 'err\tformat';
  var uid = f[1], nick = f[3];
  if (sig_(['ren', uid, nick, '', '']) !== f[4]) return 'err\tsig';
  if (!nickOk_(nick)) return 'err\tbadnick';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = userSheet_(), U = users_(sh), u = auth_(U, uid, f[2]), key = nickKey_(nick);
    if (!u) return isBanned_(U, uid) ? 'err\tban' : 'err\tauth';
    if (U.byKey[key] && U.byKey[key].uid !== u.uid) return 'err\ttaken';
    sh.getRange(u.row, 2, 1, 2).setValues([[nick, key]]);
    bump_();
    return 'ok\t' + nick;
  } finally { lock.releaseLock(); }
}
function pw_(f) {                    // pw, uid, ключ, хеш старого пароля, хеш нового, подпись
  if (f.length < 6) return 'err\tformat';
  var uid = f[1], oldph = f[3], newph = f[4];
  if (!/^[0-9a-f]{64}$/.test(newph)) return 'err\tformat';
  if (sig_(['pw', uid, oldph, newph, '']) !== f[5]) return 'err\tsig';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = userSheet_(), U = users_(sh), u = auth_(U, uid, f[2]);
    if (!u || !safeEq_(sha_(u.salt + ':' + oldph), u.hash)) return 'err\tauth';
    var salt = rnd_().substr(0, 16), token = rnd_();
    sh.getRange(u.row, 4, 1, 3).setValues([[salt, sha_(salt + ':' + newph), tokHash_(token)]]);      // остальные устройства из аккаунта выходят
    return 'ok\t' + token;
  } finally { lock.releaseLock(); }
}
function out_(f) {                   // out, uid, ключ, подпись: это устройство вышло из аккаунта
  if (f.length < 4) return 'err\tformat';
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = userSheet_(), U = users_(sh), u = auth_(U, f[1], f[2]);
    if (!u) return 'err\tauth';
    var h = tokHash_(f[2]), toks = [];
    for (var i = 0; i < u.tokens.length; i++) if (u.tokens[i] !== h) toks.push(u.tokens[i]);
    sh.getRange(u.row, 6).setValue(toks.join(','));
    return 'ok';
  } finally { lock.releaseLock(); }
}
function mine_(uid, token) {         // прогресс игрока по встроенным уровням: уровень, звёзды, тики
  var U = usersSafe_();
  if (!auth_(U, uid, token)) return isBanned_(U, uid) ? 'err\tban' : 'err\tauth';
  var all = rows_(recSheet_(), 5), out = [];
  for (var i = 0; i < all.length; i++) {
    var level = String(all[i][0]);
    if (String(all[i][1]) !== uid || !/^b\d\d-[0-9a-f]{4}$/.test(level)) continue;
    out.push(level + '\t' + (Number(all[i][3]) || 0) + '\t' + (Number(all[i][4]) || 0));
  }
  return 'ok\t' + out.length + (out.length ? '\n' + out.join('\n') : '');
}

/* ---------------------------------------------------------------- V0.4: кто в сети */
function hb_(p) {                    // a=hb, u — номер игрока, t — ключ (если есть аккаунт), w — где он сейчас, v — версия игры, p — платформа
  var uid = String(p.u || ''), w = String(p.w || '').replace(/[^0-9a-z]/g, '').substr(0, 12), v = parseInt(p.v, 10) || 0, pf = parseInt(p.p, 10) || 0, t = String(p.t || '');
  if (!/^[1-9]\d{0,9}$/.test(uid)) return 'err\tformat';
  var cache = CacheService.getScriptCache(), U = usersSafe_(), u = U.byUid[uid], nick = '', now = Math.floor(Date.now() / 1000);
  if (u && u.banned) return 'err\tban';
  if (t) { if (!auth_(U, uid, t)) return 'err\tauth'; nick = u.nick; }
  cache.put('on:' + uid, [now, nick, w, v, pf].join('|'), 300);
  var idx = cache.get('online') || '', list = idx ? idx.split(',') : [];
  if (list.indexOf(uid) < 0) { list.push(uid); if (list.length > 4000) list = list.slice(-4000); cache.put('online', list.join(','), 21600); }
  if (nick && !cache.get('ls:' + uid)) {                                              // на листе «users» время обновляется не чаще раза в 5 минут
    cache.put('ls:' + uid, '1', 300);
    try {
      var sh = userSheet_();
      if (String(sh.getRange(u.row, 1).getValue()) === uid) sh.getRange(u.row, 8, 1, 4).setValues([[new Date(), w, v, pf]]);
    } catch (e) {}
  }
  return 'ok\t' + now;
}
function onlineNow_() {
  var cache = CacheService.getScriptCache(), idx = cache.get('online') || '', list = idx ? idx.split(',') : [], out = [], keep = [], now = Math.floor(Date.now() / 1000);
  var keys = [];
  for (var k = 0; k < list.length; k++) keys.push('on:' + list[k]);
  var got = keys.length ? cache.getAll(keys) : {};
  for (var i = 0; i < list.length; i++) {
    var s = got['on:' + list[i]];
    if (!s) continue;
    keep.push(list[i]);
    var f = s.split('|'), age = now - (parseInt(f[0], 10) || 0);
    if (age <= ONLINE_SEC) out.push({ uid: list[i], nick: f[1] || '', where: f[2] || '', ver: f[3] || '0', pf: f[4] || '0', age: age });
  }
  if (keep.length !== list.length) cache.put('online', keep.join(','), 21600);
  return out;
}

/* ---------------------------------------------------------------- V0.4: админка */
/* Каждая команда подписана паролем админки (HMAC-SHA256), с временем и одноразовым числом: чужую команду нельзя ни подделать, ни повторить.
   Сам пароль по сети не ходит и в таблице не лежит: в свойствах скрипта хранится только его SHA-256. */
function melAdminPassword() {        // ЗАПУСТИ ЭТУ ФУНКЦИЮ В РЕДАКТОРЕ: «Выполнить» → новый пароль админки появится в журнале выполнения
  var abc = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789', p = '', h = sha_(Utilities.getUuid() + Utilities.getUuid() + Date.now() + Math.random());
  for (var i = 0; i < 20; i++) p += abc.charAt(parseInt(h.substr(i * 3, 3), 16) % abc.length) + (i % 5 === 4 && i < 19 ? '-' : '');
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', sha_('mel-nyam/admin/v1:' + p));
  Logger.log('Пароль админки (запиши и никому не показывай): ' + p);
  Logger.log('Старый пароль, если он был, больше не работает. Вводи новый в программе MelNyamAdmin.exe.');
  return p;
}
function adm_(f) {                   // adm, время, одноразовое число, команда, аргумент 1, аргумент 2, подпись
  if (f.length < 7) return 'err\tformat';
  var props = PropertiesService.getScriptProperties(), key = props.getProperty('ADMIN_KEY') || '';
  if (!key) return 'err\tnoadmin';
  var ts = parseInt(f[1], 10) || 0, nonce = f[2], now = Math.floor(Date.now() / 1000), cache = CacheService.getScriptCache();
  if (!/^[0-9a-f]{16,64}$/.test(nonce) || !safeEq_(hmac_(['adm', f[1], nonce, f[3], f[4], f[5]].join('|'), key), f[6])) {
    props.setProperty('ADMIN_FAILS', String((parseInt(props.getProperty('ADMIN_FAILS'), 10) || 0) + 1));      // сколько раз в админку стучались с неверным паролем
    Utilities.sleep(1000);
    return 'err\tdenied';
  }
  if (Math.abs(now - ts) > 600) return 'err\ttime\t' + now;
  if (cache.get('an:' + nonce)) return 'err\treplay';
  cache.put('an:' + nonce, '1', 1500);
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return admRun_(f[3], f[4], f[5], props); } finally { lock.releaseLock(); }
}
function admRun_(cmd, a1, a2, props) {
  var sh = userSheet_(), U = users_(sh), u = U.byUid[a1], i, out;
  if (cmd === 'hello') {
    var fails = parseInt(props.getProperty('ADMIN_FAILS'), 10) || 0;
    props.setProperty('ADMIN_FAILS', '0');
    return 'ok\t' + PROTO + '\t' + fails + '\t' + U.list.length + '\t' + Math.max(0, recSheet_().getLastRow() - 1) + '\t' + Math.max(0, lvlSheet_().getLastRow() - 1) + '\t' + onlineNow_().length;
  }
  if (cmd === 'online') {
    var on = onlineNow_(); out = ['ok\t' + on.length];
    for (i = 0; i < on.length; i++) {
      var ou = U.byUid[on[i].uid];
      out.push([on[i].uid, on[i].nick || (ou && ou.nick) || '', on[i].where, on[i].ver, on[i].pf, on[i].age, ou && ou.hash ? 1 : 0].join('\t'));
    }
    return out.join('\n');
  }
  if (cmd === 'users') {             // все, кого знает таблица: и с аккаунтом, и без (по рекордам и уровням)
    var rs = recSheet_(), recs = rows_(rs, 3), lv = rows_(lvlSheet_(), 8), m = {}, order = [], e;
    var rd = recs.length ? rs.getRange(2, 7, recs.length, 1).getValues() : [];             // когда поставлен каждый рекорд
    var get = function (uid) { if (!m[uid]) { m[uid] = { uid: uid, nick: '', reg: 0, ban: 0, created: 0, seen: 0, where: '', ver: 0, pf: 0, recs: 0, lvls: 0, note: '' }; order.push(uid); } return m[uid]; };
    for (i = 0; i < U.list.length; i++) {
      var x = U.list[i]; e = get(x.uid);
      e.nick = x.nick; e.reg = x.hash ? 1 : 0; e.ban = x.banned ? 1 : 0; e.created = ts_(x.created); e.seen = ts_(x.seen); e.where = x.where; e.ver = x.ver; e.pf = x.pf; e.note = x.note;
    }
    // «был в игре»: последний отклик игры, а у тех, кто играл до V0.4, — день последнего рекорда или уровня
    for (i = 0; i < recs.length; i++) { e = get(String(recs[i][1])); e.recs++; if (!e.nick) e.nick = String(recs[i][2]); e.seen = Math.max(e.seen, ts_(rd[i][0])); }
    for (i = 0; i < lv.length; i++) { e = get(String(lv[i][1])); e.lvls++; if (!e.nick) e.nick = String(lv[i][2]); e.seen = Math.max(e.seen, ts_(lv[i][7])); }
    out = ['ok\t' + order.length];
    for (i = 0; i < order.length; i++) { e = m[order[i]]; out.push([e.uid, e.nick, e.reg, e.ban, e.created, e.seen, e.where, e.ver, e.pf, e.recs, e.lvls, String(e.note).replace(/[\t\n\r]/g, ' ')].join('\t')); }
    return out.join('\n');
  }
  if (cmd === 'levels') {
    var all = rows_(lvlSheet_(), 9); out = ['ok\t' + all.length];
    for (i = 0; i < all.length; i++) out.push([Number(all[i][0]), String(all[i][1]), showNick_(U, String(all[i][1]), all[i][2]), String(all[i][3]), Number(all[i][4]) || 0, Number(all[i][6]) || 0, String(all[i][8]) !== '' ? 1 : 0, ts_(all[i][7])].join('\t'));
    return out.join('\n');
  }
  if (cmd === 'hide') {              // скрыть уровень игрока (a2 = 1) или показать снова (a2 = 0)
    var ls = lvlSheet_(), lr = findLevelRow_(ls, parseInt(a1, 10) || 0);
    if (!lr) return 'err\tnot found';
    ls.getRange(lr, 9).setValue(a2 === '1' ? 'x' : '');
    bump_();
    return 'ok';
  }
  if (!/^[1-9]\d{0,9}$/.test(a1)) return 'err\tformat';
  if (cmd === 'ban' || cmd === 'unban') {
    var mark = cmd === 'ban' ? 'x' : '', note = cmd === 'ban' ? String(a2).substr(0, 200) : '';
    if (u) sh.getRange(u.row, 12, 1, 2).setValues([[mark, note]]);
    else if (cmd === 'ban') sh.appendRow([a1, '', '', '', '', '', new Date(), '', '', 0, '', mark, note]);       // игрок без аккаунта: строка только ради блокировки
    CacheService.getScriptCache().remove('on:' + a1);
    bump_();
    return 'ok';
  }
  if (cmd === 'del') {               // удалить игрока совсем: аккаунт, рекорды, уровни
    var n = delRows_(recSheet_(), 2, a1) + delRows_(lvlSheet_(), 2, a1);
    if (u) sh.deleteRow(u.row);
    CacheService.getScriptCache().remove('on:' + a1);
    bump_();
    return 'ok\t' + n + '\t' + (u ? 1 : 0);
  }
  if (cmd === 'resetpw') {           // новый пароль игроку (a2 — SHA-256 пароля, как его прислала бы игра); со всех устройств он выходит
    if (!u || !u.hash) return 'err\tnot found';
    if (!/^[0-9a-f]{64}$/.test(a2)) return 'err\tformat';
    var salt = rnd_().substr(0, 16);
    sh.getRange(u.row, 4, 1, 3).setValues([[salt, sha_(salt + ':' + a2), '']]);
    return 'ok';
  }
  if (cmd === 'nick') {              // сменить ник игроку
    if (!u || !u.hash) return 'err\tnot found';
    if (!nickOk_(a2)) return 'err\tbadnick';
    if (U.byKey[nickKey_(a2)] && U.byKey[nickKey_(a2)].uid !== u.uid) return 'err\ttaken';
    sh.getRange(u.row, 2, 1, 2).setValues([[a2, nickKey_(a2)]]);
    bump_();
    return 'ok';
  }
  return 'err\tunknown';
}
function delRows_(sh, col, uid) {    // удалить все строки игрока, снизу вверх
  var n = sh.getLastRow() - 1, cnt = 0;
  if (n < 1) return 0;
  var v = sh.getRange(2, col, n, 1).getValues();
  for (var i = n - 1; i >= 0; i--) if (String(v[i][0]) === uid) { sh.deleteRow(i + 2); cnt++; }
  return cnt;
}
