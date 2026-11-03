/**
 * Last Ticket Resort — order book backend.
 * Lives inside a Google Sheet (Extensions > Apps Script).
 * Deploy > New deployment > Web app: Execute as "Me", access "Anyone".
 *
 * Sheets (created by setup()):
 *   Quotes   – every buy / sell quote
 *   Matches  – deals you confirm
 *   Settings – ADMIN_KEY, FEE
 */

var QUOTE_COLS = ['id','created','side','event','category','city','date','qty','openQty',
  'minPrice','maxPrice','askPrice','faceValue','name','phone','notes','status','fromPlace','toPlace'];
var MATCH_COLS = ['id','created','event','date','buyId','sellId','qty','price','fee','status','note','sellPrice','margin'];

/* ---------- one-time setup: run this once from the editor ---------- */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, 'Quotes', QUOTE_COLS);
  ensureSheet_(ss, 'Matches', MATCH_COLS);
  var s = ensureSheet_(ss, 'Settings', ['key', 'value']);
  if (!getSetting_('ADMIN_KEY')) s.appendRow(['ADMIN_KEY', Utilities.getUuid().replace(/-/g, '').slice(0, 16)]);
  if (!getSetting_('FEE')) s.appendRow(['FEE', 99]);
  var first = ss.getSheetByName('Sheet1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);
  Logger.log('Admin key: ' + getSetting_('ADMIN_KEY'));
}

function ensureSheet_(ss, name, cols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(cols); sh.setFrozenRows(1); sh.getRange(1, 1, 1, cols.length).setFontWeight('bold'); }
  return sh;
}

function getSetting_(key) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) return '';
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === key) return v[i][1];
  return '';
}

/* ---------- helpers ---------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function rows_(name, cols) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  var v = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var c = 0; c < cols.length; c++) {
      var x = v[i][c];
      if (x instanceof Date) x = Utilities.formatDate(x, 'Asia/Kolkata', cols[c] === 'date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
      o[cols[c]] = x;
    }
    out.push(o);
  }
  return out;
}
function setCell_(name, cols, row, field, value) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name).getRange(row, cols.indexOf(field) + 1).setValue(value);
}
function clean_(s, n) { return String(s == null ? '' : s).replace(/[\u0000-\u001f<>]/g, ' ').trim().slice(0, n || 80); }
function num_(x) { if (x === '' || x == null) return null; var n = Number(x); return isFinite(n) && n >= 0 ? Math.round(n) : null; }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'); }
function norm_(s) { return String(s).toLowerCase().replace(/\s+/g, ' ').trim(); }
function eventKey_(q) { return [norm_(q.event), q.date, norm_(q.city), norm_(q.fromPlace || ''), norm_(q.toPlace || '')].join('|'); }
function isAdmin_(key) { return key && String(key) === String(getSetting_('ADMIN_KEY')); }

/* ---------- GET: public board / admin data ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'admin') {
      if (!isAdmin_(p.key)) return json_({ ok: false, error: 'Wrong admin key.' });
      return json_({ ok: true, fee: Number(getSetting_('FEE')) || 0, quotes: rows_('Quotes', QUOTE_COLS), matches: rows_('Matches', MATCH_COLS), events: board_() });
    }
    return json_({ ok: true, events: board_().map(function (e) { return { event: e.event, date: e.date, city: e.city, category: e.category, fromPlace: e.fromPlace, toPlace: e.toPlace, demand: e.demand, supply: e.supply, buyers: e.buyers, sellers: e.sellers }; }) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** Aggregated demand / supply per event. No names or phone numbers. */
function board_() {
  var t = today_();
  var quotes = rows_('Quotes', QUOTE_COLS).filter(function (q) { return q.status === 'OPEN' && Number(q.openQty) > 0 && (!q.date || q.date >= t); });
  var matches = rows_('Matches', MATCH_COLS).filter(function (m) { return m.status !== 'CANCELLED'; });
  var ev = {};
  quotes.forEach(function (q) {
    var k = eventKey_(q);
    var e = ev[k] || (ev[k] = { event: q.event, date: q.date, city: q.city, category: q.category, fromPlace: q.fromPlace || '', toPlace: q.toPlace || '',
      demand: 0, supply: 0, buyers: 0, sellers: 0, bestBid: null, bestAsk: null, bidLow: null, lastPrice: null, matched: 0 });
    var n = Number(q.openQty);
    if (q.side === 'BUY') {
      e.demand += n; e.buyers++;
      e.bestBid = e.bestBid == null ? Number(q.maxPrice) : Math.max(e.bestBid, Number(q.maxPrice));
      e.bidLow = e.bidLow == null ? Number(q.minPrice) : Math.min(e.bidLow, Number(q.minPrice));
    } else {
      e.supply += n; e.sellers++;
      e.bestAsk = e.bestAsk == null ? Number(q.askPrice) : Math.min(e.bestAsk, Number(q.askPrice));
    }
  });
  matches.forEach(function (m) {
    var k = norm_(m.event) + '|' + m.date + '|';
    for (var key in ev) if (key.indexOf(k) === 0) { ev[key].lastPrice = Number(m.price); ev[key].matched += Number(m.qty); }
  });
  return Object.keys(ev).map(function (k) { return ev[k]; });
}

/* ---------- POST: submit quote / admin actions ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var b = JSON.parse(e.postData.contents || '{}');
    if (b.action === 'quote') return json_(addQuote_(b));
    if (!isAdmin_(b.key)) return json_({ ok: false, error: 'Wrong admin key.' });
    if (b.action === 'match') return json_(addMatch_(b));
    if (b.action === 'quoteStatus') return json_(quoteStatus_(b));
    if (b.action === 'matchStatus') return json_(matchStatus_(b));
    if (b.action === 'setFee') { setFee_(b.fee); return json_({ ok: true }); }
    return json_({ ok: false, error: 'Unknown action.' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function addQuote_(b) {
  if (b.website) return { ok: true }; // honeypot: bots fill hidden field
  ensureQuoteHeader_();
  var side = b.side === 'SELL' ? 'SELL' : 'BUY';
  var q = {
    event: clean_(b.event), category: clean_(b.category, 30), city: clean_(b.city, 40), date: clean_(b.date, 10),
    qty: num_(b.qty), minPrice: num_(b.minPrice), maxPrice: num_(b.maxPrice), askPrice: num_(b.askPrice),
    faceValue: num_(b.faceValue), name: clean_(b.name, 60), phone: clean_(b.phone, 20), notes: clean_(b.notes, 300),
    fromPlace: clean_(b.fromPlace, 40), toPlace: clean_(b.toPlace, 40)
  };
  if (!q.event) return { ok: false, error: 'Add the event or purpose.' };
  if (!!q.fromPlace !== !!q.toPlace) return { ok: false, error: 'Fill in both From and To, or leave both empty.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(q.date) || q.date < today_()) return { ok: false, error: 'Add an upcoming event date.' };
  if (!q.qty || q.qty > 50) return { ok: false, error: 'Number of tickets must be between 1 and 50.' };
  if (!q.name || !/^[+\d][\d\s-]{8,18}$/.test(q.phone)) return { ok: false, error: 'Add your name and a valid phone number.' };
  if (/\b(train|irctc|pnr|flight|airline)\b/i.test(q.event + ' ' + q.notes)) return { ok: false, error: 'Train and flight tickets are name-bound and cannot be traded here.' };
  if (side === 'BUY') {
    if (q.minPrice == null || q.maxPrice == null || q.maxPrice < q.minPrice) return { ok: false, error: 'Give a price range where max is at least min.' };
    q.askPrice = '';
  } else {
    if (q.askPrice == null) return { ok: false, error: 'Add your asking price per ticket.' };
    q.minPrice = ''; q.maxPrice = '';
  }
  var id = (side === 'BUY' ? 'B-' : 'S-') + Utilities.getUuid().slice(0, 8).toUpperCase();
  var row = [id, new Date(), side, q.event, q.category, q.city, q.date, q.qty, q.qty,
    q.minPrice, q.maxPrice, q.askPrice, q.faceValue == null ? '' : q.faceValue, q.name, "'" + q.phone, q.notes, 'OPEN', q.fromPlace, q.toPlace];
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Quotes').appendRow(row);
  return { ok: true, id: id };
}

function ensureQuoteHeader_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Quotes');
  if (sh.getLastColumn() < QUOTE_COLS.length) sh.getRange(1, 1, 1, QUOTE_COLS.length).setValues([QUOTE_COLS]).setFontWeight('bold');
}

function findQuote_(id) {
  var all = rows_('Quotes', QUOTE_COLS);
  for (var i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
  return null;
}

function addMatch_(b) {
  var buy = findQuote_(b.buyId), sell = findQuote_(b.sellId);
  var qty = num_(b.qty), price = num_(b.buyPrice != null ? b.buyPrice : b.price), sellPrice = num_(b.sellPrice != null ? b.sellPrice : price);
  if (!buy || !sell || buy.side !== 'BUY' || sell.side !== 'SELL') return { ok: false, error: 'Quote not found.' };
  if (buy.status !== 'OPEN' || sell.status !== 'OPEN') return { ok: false, error: 'One of these quotes is no longer open.' };
  if (!qty || qty > Number(buy.openQty) || qty > Number(sell.openQty)) return { ok: false, error: 'Quantity is more than what is open.' };
  if (price == null || sellPrice == null) return { ok: false, error: 'Add the buyer and seller prices.' };
  if (sellPrice > price) return { ok: false, error: 'Seller price is above the buyer price.' };
  var msh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches');
  if (msh.getLastColumn() < MATCH_COLS.length) msh.getRange(1, 1, 1, MATCH_COLS.length).setValues([MATCH_COLS]).setFontWeight('bold');
  var id = 'M-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  var fee = Number(getSetting_('FEE')) || 0;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Matches').appendRow(
    [id, new Date(), buy.event, buy.date, buy.id, sell.id, qty, price, fee, 'PROPOSED', clean_(b.note, 200), sellPrice, (price - sellPrice) * qty]);
  [buy, sell].forEach(function (q) {
    var left = Number(q.openQty) - qty;
    setCell_('Quotes', QUOTE_COLS, q._row, 'openQty', left);
    if (left <= 0) setCell_('Quotes', QUOTE_COLS, q._row, 'status', 'FILLED');
  });
  return { ok: true, id: id };
}

function quoteStatus_(b) {
  var q = findQuote_(b.id);
  if (!q) return { ok: false, error: 'Quote not found.' };
  var s = ['OPEN', 'CANCELLED', 'FILLED'].indexOf(b.status) >= 0 ? b.status : 'OPEN';
  setCell_('Quotes', QUOTE_COLS, q._row, 'status', s);
  return { ok: true };
}

function matchStatus_(b) {
  var all = rows_('Matches', MATCH_COLS), m = null;
  for (var i = 0; i < all.length; i++) if (all[i].id === b.id) m = all[i];
  if (!m) return { ok: false, error: 'Match not found.' };
  var s = ['PROPOSED', 'MEETING_SET', 'DONE', 'CANCELLED'].indexOf(b.status) >= 0 ? b.status : 'PROPOSED';
  if (s === 'CANCELLED' && m.status !== 'CANCELLED') {
    [m.buyId, m.sellId].forEach(function (qid) {
      var q = findQuote_(qid);
      if (q) { setCell_('Quotes', QUOTE_COLS, q._row, 'openQty', Number(q.openQty) + Number(m.qty)); setCell_('Quotes', QUOTE_COLS, q._row, 'status', 'OPEN'); }
    });
  }
  setCell_('Matches', MATCH_COLS, m._row, 'status', s);
  return { ok: true };
}

function setFee_(fee) {
  var n = num_(fee); if (n == null) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === 'FEE') { sh.getRange(i + 1, 2).setValue(n); return; }
  sh.appendRow(['FEE', n]);
}
