/**
 * ホストクラブ（ホストパロ）用 GAS
 * 新しいスプレッドシートの「拡張機能 → Apps Script」にこのコードを丸ごと貼り付けて、
 * 「デプロイ → 新しいデプロイ → ウェブアプリ（実行ユーザー：自分 / アクセス：全員）」で公開してください。
 * シートは初回アクセス時に自動で作られます。
 * 管理パスワードと合言葉は「設定」シートで変更できます（画面には出ません）。
 */

const TZ = 'Asia/Tokyo';
const SHEETS = {
  hosts:    { name: 'ホスト',   cols: ['id', 'name', 'icon', 'catch', 'passHash', 'token', 'createdAt'] },
  guests:   { name: 'お客様',   cols: ['id', 'name', 'icon', 'points', 'passHash', 'token', 'createdAt', 'lastBonus', 'blogDay', 'blogCount'] },
  gifts:    { name: '貢ぎ物',   cols: ['id', 'name', 'price', 'icon'] },
  tributes: { name: '貢ぎ履歴', cols: ['id', 'at', 'guestId', 'hostId', 'giftId', 'giftName', 'price', 'icon', 'message'] },
  settings: { name: '設定',     cols: ['key', 'value', 'memo'] },
  blog:     { name: 'ブログ',   cols: ['id', 'at', 'guestId', 'title', 'body', 'points'] }
};
const DEFAULT_SETTINGS = [
  ['clubName', 'CLUB NOCTURNE', 'お店の名前（画面のタイトル）'],
  ['initialPoints', 10000, 'お客様が登録したときにもらえるポイント'],
  ['dailyBonus', 0, '1日1回もらえるポイント（0ならボーナスなし）'],
  ['blogPoints', 3000, 'ブログを1記事書くともらえるポイント'],
  ['blogDailyMax', 3, 'ポイントがもらえるのは1日何記事まで'],
  ['blogMinChars', 50, 'ポイントがもらえる最低文字数'],
  ['roomKey', '', '登録に必要な合言葉（空なら誰でも登録できる）'],
  ['adminPassword', 'changeme', '管理画面のパスワード（必ず変えてください）']
];
const DEFAULT_GIFTS = [
  ['g1', 'ドリンク', 300, 'local_bar'],
  ['g2', '指名', 500, 'favorite'],
  ['g3', '花束', 1000, 'local_florist'],
  ['g4', 'ケーキ', 2000, 'cake'],
  ['g5', 'シャンパン', 5000, 'wine_bar'],
  ['g6', 'ブランド時計', 10000, 'watch'],
  ['g7', 'ドンペリ', 20000, 'liquor'],
  ['g8', 'シャンパンタワー', 50000, 'celebration'],
  ['g9', '高級車', 100000, 'directions_car']
];

// ---------- 入口 ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  let body = {};
  try { body = p.data ? JSON.parse(p.data) : p; } catch (err) { body = p; }
  return respond(handle(body), p.callback);
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return respond({ error: '送信データが読めませんでした' }); }
  return respond(handle(body));
}
function respond(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function handle(b) {
  try {
    setup_();
    switch (b.action) {
      case 'readAll':       return readAll_(b);
      case 'register':      return withLock_(() => register_(b));
      case 'login':         return login_(b);
      case 'uploadIcon':    return uploadIcon_(b);
      case 'updateProfile': return withLock_(() => updateProfile_(b));
      case 'claimBonus':    return withLock_(() => claimBonus_(b));
      case 'tribute':       return withLock_(() => tribute_(b));
      case 'postBlog':      return withLock_(() => postBlog_(b));
      case 'deleteBlog':    return withLock_(() => deleteBlog_(b));
      case 'admin':         return withLock_(() => admin_(b));
      default:              return { error: '不明な操作です' };
    }
  } catch (err) {
    return { error: String(err && err.message || err) };
  }
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- シート操作 ----------
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(key) { return ss_().getSheetByName(SHEETS[key].name); }
function setup_() {
  const ss = ss_();
  Object.keys(SHEETS).forEach(key => {
    const def = SHEETS[key];
    const old = ss.getSheetByName(def.name);
    if (old) {
      // あとから増えた列・設定を足す（既存のデータはそのまま）
      const head = old.getRange(1, 1, 1, Math.max(1, old.getLastColumn())).getValues()[0];
      if (head.length < def.cols.length) old.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
      if (key === 'settings') {
        const keys = old.getLastRow() > 1 ? old.getRange(2, 1, old.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
        DEFAULT_SETTINGS.filter(d => keys.indexOf(d[0]) < 0).forEach(d => old.appendRow(d));
      }
      return;
    }
    const sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (key === 'settings') sh.getRange(2, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    if (key === 'gifts') sh.getRange(2, 1, DEFAULT_GIFTS.length, 4).setValues(DEFAULT_GIFTS);
  });
}
function rows_(key) {
  const sh = sheet_(key);
  const last = sh.getLastRow();
  const cols = SHEETS[key].cols;
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    cols.forEach((c, j) => { o[c] = r[j]; });
    return o;
  }).filter(o => o[cols[0]] !== '');
}
// 「=」などで始まる文字が数式として扱われないようにする
function safe_(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }
function append_(key, obj) {
  sheet_(key).appendRow(SHEETS[key].cols.map(c => obj[c] === undefined ? '' : safe_(obj[c])));
}
function setCell_(key, row, col, value) {
  sheet_(key).getRange(row, SHEETS[key].cols.indexOf(col) + 1).setValue(safe_(value));
}
function settings_() {
  const s = {};
  rows_('settings').forEach(r => { s[r.key] = r.value; });
  DEFAULT_SETTINGS.forEach(d => { if (s[d[0]] === undefined) s[d[0]] = d[1]; });
  return s;
}
function setSetting_(key, value) {
  const r = rows_('settings').find(x => x.key === key);
  if (r) setCell_('settings', r._row, 'value', value);
  else append_('settings', { key: key, value: value, memo: '' });
}
function hash_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'nocturne:' + text, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function newId_(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function nowIso_() { return new Date().toISOString(); }
function monthOf_(iso) { return Utilities.formatDate(new Date(iso), TZ, 'yyyy-MM'); }
function dayOf_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max); }
function roleKey_(role) {
  if (role === 'host') return 'hosts';
  if (role === 'guest') return 'guests';
  throw new Error('ホストかお客様を選んでください');
}
function me_(b) {
  if (!b.token) return null;
  const key = roleKey_(b.role);
  const u = rows_(key).find(r => r.token === b.token);
  return u ? { key: key, user: u } : null;
}

// ---------- 読み込み ----------
function readAll_(b) {
  const s = settings_();
  const hosts = rows_('hosts');
  const guests = rows_('guests');
  const tributes = rows_('tributes');
  const gifts = rows_('gifts').map(g => ({ id: String(g.id), name: g.name, price: Number(g.price) || 0, icon: g.icon || 'redeem' }))
    .sort((a, b) => a.price - b.price);
  const month = monthOf_(nowIso_());
  const dayAgo = Date.now() - 24 * 3600 * 1000;
  const guestMap = {};
  guests.forEach(g => { guestMap[g.id] = g; });

  const stat = {};
  hosts.forEach(h => { stat[h.id] = { total: 0, month: 0, auraPrice: 0, patrons: {} }; });
  const guestGiven = {};
  tributes.forEach(t => {
    const price = Number(t.price) || 0;
    const st = stat[t.hostId];
    guestGiven[t.guestId] = (guestGiven[t.guestId] || 0) + price;
    if (!st) return;
    st.total += price;
    if (monthOf_(t.at) === month) st.month += price;
    if (new Date(t.at).getTime() >= dayAgo) st.auraPrice = Math.max(st.auraPrice, price);
    st.patrons[t.guestId] = (st.patrons[t.guestId] || 0) + price;
  });

  const out = {
    settings: { clubName: s.clubName, initialPoints: Number(s.initialPoints) || 0, dailyBonus: Number(s.dailyBonus) || 0, needsKey: !!String(s.roomKey || '').trim(),
                blogPoints: Number(s.blogPoints) || 0, blogDailyMax: Number(s.blogDailyMax) || 0, blogMinChars: Number(s.blogMinChars) || 0 },
    posts: rows_('blog').slice(-100).reverse().map(p => ({ id: p.id, at: p.at, guestId: p.guestId, title: p.title, body: p.body, points: Number(p.points) || 0 })),
    month: month,
    gifts: gifts,
    hosts: hosts.map(h => {
      const st = stat[h.id];
      const patrons = Object.keys(st.patrons).map(id => ({ id: id, total: st.patrons[id] }))
        .sort((a, b) => b.total - a.total).slice(0, 3);
      return { id: h.id, name: h.name, icon: h.icon, catch: h.catch, total: st.total, month: st.month, auraPrice: st.auraPrice, patrons: patrons };
    }),
    guests: guests.map(g => ({ id: g.id, name: g.name, icon: g.icon, given: guestGiven[g.id] || 0 })),
    tributes: tributes.slice(-60).reverse().map(t => ({
      id: t.id, at: t.at, guestId: t.guestId, hostId: t.hostId, giftName: t.giftName, price: Number(t.price) || 0, icon: t.icon, message: t.message
    }))
  };
  const me = me_(b);
  if (me) {
    out.me = { role: b.role, id: me.user.id, name: me.user.name };
    const field = me.key === 'guests' ? 'guestId' : 'hostId';
    const mine = tributes.filter(t => t[field] === me.user.id);
    out.me.byPartner = {};
    mine.forEach(t => {
      const other = me.key === 'guests' ? t.hostId : t.guestId;
      out.me.byPartner[other] = (out.me.byPartner[other] || 0) + (Number(t.price) || 0);
    });
    out.me.history = mine.slice(-50).reverse().map(t => ({
      id: t.id, at: t.at, guestId: t.guestId, hostId: t.hostId, giftName: t.giftName, price: Number(t.price) || 0, icon: t.icon, message: t.message
    }));
    if (me.key === 'guests') {
      out.me.points = Number(me.user.points) || 0;
      out.me.bonusReady = (Number(s.dailyBonus) || 0) > 0 && String(me.user.lastBonus) !== dayOf_(new Date());
      const used = String(me.user.blogDay) === dayOf_(new Date()) ? Number(me.user.blogCount) || 0 : 0;
      out.me.blogLeft = Math.max(0, (Number(s.blogDailyMax) || 0) - used);
    }
  }
  return out;
}

// ---------- 登録・ログイン ----------
function register_(b) {
  const s = settings_();
  const key = roleKey_(b.role);
  const roomKey = String(s.roomKey || '').trim();
  if (roomKey && clean_(b.key, 100) !== roomKey) throw new Error('合言葉がちがいます');
  const name = clean_(b.name, 20);
  const pass = String(b.password || '');
  if (!name) throw new Error('名前を入れてください');
  if (pass.length < 4) throw new Error('パスワードは4文字以上にしてください');
  if (rows_(key).some(r => String(r.name) === name)) throw new Error('その名前はもう登録されています');
  const id = newId_(key === 'hosts' ? 'h' : 'c');
  const token = Utilities.getUuid();
  const obj = { id: id, name: name, icon: clean_(b.icon, 500), passHash: hash_(pass), token: token, createdAt: nowIso_() };
  if (key === 'hosts') obj.catch = clean_(b.catch, 40);
  else { obj.points = Number(s.initialPoints) || 0; obj.lastBonus = ''; }
  append_(key, obj);
  return { ok: true, id: id, token: token };
}
function login_(b) {
  const key = roleKey_(b.role);
  const name = clean_(b.name, 20);
  const u = rows_(key).find(r => String(r.name) === name);
  if (!u || u.passHash !== hash_(String(b.password || ''))) throw new Error('名前かパスワードがちがいます');
  let token = u.token;
  if (!token) { token = Utilities.getUuid(); setCell_(key, u._row, 'token', token); }
  return { ok: true, id: u.id, token: token };
}
function uploadIcon_(b) {
  const m = String(b.image || '').match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('画像が読めませんでした');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'icon_' + Date.now() + '.jpg');
  const folders = DriveApp.getFoldersByName('ホストクラブ_アイコン');
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('ホストクラブ_アイコン');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: 'https://lh3.googleusercontent.com/d/' + file.getId() };
}
function updateProfile_(b) {
  const me = me_(b);
  if (!me) throw new Error('ログインし直してください');
  if (b.icon !== undefined) setCell_(me.key, me.user._row, 'icon', clean_(b.icon, 500));
  if (b.catch !== undefined && me.key === 'hosts') setCell_(me.key, me.user._row, 'catch', clean_(b.catch, 40));
  if (b.password) {
    if (String(b.password).length < 4) throw new Error('パスワードは4文字以上にしてください');
    setCell_(me.key, me.user._row, 'passHash', hash_(String(b.password)));
  }
  return { ok: true };
}

// ---------- ポイント ----------
function claimBonus_(b) {
  b.role = 'guest';
  const me = me_(b);
  if (!me) throw new Error('ログインし直してください');
  const today = dayOf_(new Date());
  if (String(me.user.lastBonus) === today) throw new Error('今日のボーナスはもう受け取っています');
  const bonus = Number(settings_().dailyBonus) || 0;
  if (bonus <= 0) throw new Error('今はボーナスはありません');
  const points = (Number(me.user.points) || 0) + bonus;
  setCell_('guests', me.user._row, 'points', points);
  setCell_('guests', me.user._row, 'lastBonus', today);
  return { ok: true, bonus: bonus, points: points };
}
function tribute_(b) {
  b.role = 'guest';
  const me = me_(b);
  if (!me) throw new Error('ログインし直してください');
  const host = rows_('hosts').find(h => h.id === b.hostId);
  if (!host) throw new Error('ホストが見つかりません');
  const gift = rows_('gifts').find(g => String(g.id) === String(b.giftId));
  if (!gift) throw new Error('貢ぎ物が見つかりません');
  const count = Math.max(1, Math.min(99, Math.floor(Number(b.count) || 1)));
  const price = (Number(gift.price) || 0) * count;
  const points = Number(me.user.points) || 0;
  if (price > points) throw new Error('ポイントが足りません');
  setCell_('guests', me.user._row, 'points', points - price);
  const t = {
    id: newId_('t'), at: nowIso_(), guestId: me.user.id, hostId: host.id, giftId: gift.id,
    giftName: count > 1 ? gift.name + ' ×' + count : gift.name, price: price, icon: gift.icon || 'redeem', message: clean_(b.message, 60)
  };
  append_('tributes', t);
  return { ok: true, points: points - price, tribute: t };
}

// ---------- ブログ（お客様だけ書ける。書くとポイント） ----------
function postBlog_(b) {
  b.role = 'guest';
  const me = me_(b);
  if (!me) throw new Error('お客様としてログインしてください');
  const s = settings_();
  const title = clean_(b.title, 40);
  const body = String(b.body == null ? '' : b.body).replace(/\r\n?/g, '\n').trim().slice(0, 2000);
  if (!body) throw new Error('本文を書いてください');
  const today = dayOf_(new Date());
  const used = String(me.user.blogDay) === today ? Number(me.user.blogCount) || 0 : 0;
  const chars = body.replace(/\s/g, '').length;
  let points = 0;
  if (used < (Number(s.blogDailyMax) || 0) && chars >= (Number(s.blogMinChars) || 0)) points = Number(s.blogPoints) || 0;
  const post = { id: newId_('b'), at: nowIso_(), guestId: me.user.id, title: title, body: body, points: points };
  append_('blog', post);
  if (points > 0) {
    setCell_('guests', me.user._row, 'points', (Number(me.user.points) || 0) + points);
    setCell_('guests', me.user._row, 'blogDay', today);
    setCell_('guests', me.user._row, 'blogCount', used + 1);
  }
  return { ok: true, post: post, earned: points, chars: chars };
}
function deleteBlog_(b) {
  b.role = 'guest';
  const me = me_(b);
  if (!me) throw new Error('ログインし直してください');
  const p = rows_('blog').find(x => x.id === b.id);
  if (!p || p.guestId !== me.user.id) throw new Error('この記事は消せません');
  sheet_('blog').deleteRow(p._row);
  return { ok: true };
}

// ---------- 管理 ----------
function admin_(b) {
  const s = settings_();
  if (String(b.adminPassword || '') !== String(s.adminPassword)) throw new Error('管理パスワードがちがいます');
  switch (b.op) {
    case 'check':
      return { ok: true, settings: { clubName: s.clubName, initialPoints: s.initialPoints, dailyBonus: s.dailyBonus, roomKey: s.roomKey,
                                     blogPoints: s.blogPoints, blogDailyMax: s.blogDailyMax, blogMinChars: s.blogMinChars },
               guests: rows_('guests').map(g => ({ id: g.id, name: g.name, points: Number(g.points) || 0 })) };
    case 'saveSettings': {
      const NUM = ['initialPoints', 'dailyBonus', 'blogPoints', 'blogDailyMax', 'blogMinChars'];
      ['clubName', 'roomKey'].concat(NUM).forEach(k => {
        if (b.settings && b.settings[k] !== undefined) setSetting_(k, NUM.indexOf(k) >= 0 ? Math.max(0, Math.floor(Number(b.settings[k]) || 0)) : clean_(b.settings[k], 60));
      });
      return { ok: true };
    }
    case 'saveGifts': {
      const list = (b.gifts || []).map(g => [clean_(g.id, 20) || newId_('g'), safe_(clean_(g.name, 20)), Math.max(1, Math.floor(Number(g.price) || 0)), clean_(g.icon, 40) || 'redeem'])
        .filter(g => g[1]);
      const sh = sheet_('gifts');
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
      if (list.length) sh.getRange(2, 1, list.length, 4).setValues(list);
      return { ok: true };
    }
    case 'givePoints': {
      const amount = Math.floor(Number(b.amount) || 0);
      const targets = b.guestId === 'all' ? rows_('guests') : rows_('guests').filter(g => g.id === b.guestId);
      if (!targets.length) throw new Error('お客様が見つかりません');
      targets.forEach(g => setCell_('guests', g._row, 'points', Math.max(0, (Number(g.points) || 0) + amount)));
      return { ok: true };
    }
    case 'deleteUser': {
      const key = roleKey_(b.role);
      const u = rows_(key).find(r => r.id === b.id);
      if (!u) throw new Error('見つかりません');
      sheet_(key).deleteRow(u._row);
      return { ok: true };
    }
    case 'deleteBlog': {
      const p = rows_('blog').find(x => x.id === b.id);
      if (!p) throw new Error('見つかりません');
      sheet_('blog').deleteRow(p._row);
      return { ok: true };
    }
    case 'resetPassword': {
      const key = roleKey_(b.role);
      const u = rows_(key).find(r => r.id === b.id);
      if (!u) throw new Error('見つかりません');
      if (String(b.password || '').length < 4) throw new Error('パスワードは4文字以上にしてください');
      setCell_(key, u._row, 'passHash', hash_(String(b.password)));
      return { ok: true };
    }
    default:
      return { error: '不明な管理操作です' };
  }
}
