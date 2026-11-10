/*****************************************************************
 * 車輛銷售管理系統 — 後端主程式（backend.js）
 *
 * 這個檔案放在 GitHub，由 Google Apps Script 裡的 Code.gs 自動下載執行，
 * 所以更新後端只要改這個檔案，不用再到 Apps Script 貼程式。
 *
 * ⚠ 這個檔案是公開的：不要在這裡放任何密碼或個資。
 *   密碼（API_KEY）只放在 Apps Script 的 Code.gs 和 Google 的指令碼屬性裡。
 *****************************************************************/
const BACKEND_VERSION = 'v25';   // 每次改後端就加 1

// API_KEY 由 Apps Script 裡的 Code.gs 傳進來；執行 setup 後會另外存起來。
const KEY_PLACEHOLDER = '請改成你自己的密碼';
function apiKey_() {
  const k = String(API_KEY).trim();
  if (k && k !== KEY_PLACEHOLDER) return k;
  return PropertiesService.getScriptProperties().getProperty('API_KEY') || '';
}
const REMIND_EMAIL = '';                // 提醒信寄到哪；留白 = 寄給自己
const REMIND_HOUR = 9;                  // 每天幾點寄提醒
const TZ = 'Asia/Taipei';
const RENEW_WARN_DAYS = 60;             // 第一年保險到期前幾天開始列入續保提醒

const SHEET = { db: '車輛資料庫', sales: '銷售進行中', taxi: '計程車', cash: '零用金', pw: '帳號密碼' };
const FIELDS = [
  'id', '類型', '月份', '客戶姓名', '電話',
  '車牌', '廠牌', '車型', '年式', '顏色', '車身號碼', '引擎號碼',
  '配件', '保險公司', '強制險到期', '任意險到期',
  '領牌日', '交車日', '特殊備註', '流程', '狀態',
  '建立時間', '更新時間', '歸檔時間',
  '投保日期', '險種', '保險金額',   // v2 新增（放在最後，舊資料不會錯位）
  '續保狀態', '續保日期', '續保保險公司', '續保險種', '續保金額', '未續保原因', '聯絡紀錄',   // v3 新增
  '小百貨已訂購', '洗車工單已開', '洗車工單號碼', 'ETAG已完成', '小百貨已安裝', '交車已上傳', '不聯繫已註記', '整備通知時間',   // v4 新增
  '等級', '價位', '拖車公司', '拖車安排日', '拖車抵達日', '折扣', '已付金額', '折後金額', '車款餘額',   // v5 新增（計程車）
  '成交價', '訂金日期', '訂金金額', '收款紀錄', '已收金額', '未收車款',   // v6 新增（分次收款）
  '訂單已交助理', '訂金付款方式', '訂金刷卡日', '訂金簽單已印', '訂金收據已交助理', '系統已按預估', '小百貨已交助理',
  '車籍紙本已填', 'EIP領取單已打', '身分證已收', '印章已收', '和安已試算', '強制險已開',
  '額外配件已確認', '手冊已畫線', '溫馨提醒已寫', '名片已放',   // v7 新增（流程檢查項目）
  '總價', '車價', '配件成本', '補貼息', '趴數',   // v8 新增（趴數試算）
  '訂單編號', '受訂日', '型號', '契約書號', '訂約人', '使用人', '原廠狀態', '合約配件金額'   // v9 新增（原廠受訂報告匯入）
];
// 零用金分頁有自己的欄位
const CASH_FIELDS = ['id', '日期', '類型', '分類', '項目', '金額', '備註', '建立時間', '更新時間'];
// 帳號密碼分頁
const PW_FIELDS = ['id', '系統名稱', '網址', '帳號', '密碼', '備註', '建立時間', '更新時間'];
const OWN_FIELDS = { cash: CASH_FIELDS, pw: PW_FIELDS };
const fieldsOf_ = key => OWN_FIELDS[key] || FIELDS;

const PLATE_STEP = { sales: '領牌', taxi: '登記領牌資料' };

/* ---------- 初始化 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  const k = String(API_KEY).trim();
  if (k && k !== KEY_PLACEHOLDER) PropertiesService.getScriptProperties().setProperty('API_KEY', k);
  else if (!apiKey_()) throw new Error('請先在 Apps Script 的 Code.gs 把 API_KEY 改成你自己的密碼，再執行 setup');
  Object.values(SHEET).forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    const key = Object.keys(SHEET).find(k => SHEET[k] === name);
    const F = fieldsOf_(key);
    sh.getRange(1, 1, 1, F.length).setValues([F])
      .setFontWeight('bold').setBackground('#16202b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), F.length).setNumberFormat('@'); // 全部存純文字，避免日期/電話被轉格式
  });
  ['工作表1', 'Sheet1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'dailyReminder')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1)
    .atHour(REMIND_HOUR).inTimezone(TZ).create();
}

/* ---------- API ---------- */
function doGet() {
  return ContentService.createTextOutput('車輛管理 API 運作中（後端 ' + BACKEND_VERSION + '）');
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  const key = apiKey_();
  if (!key) return out_({ ok: false, error: '後端還沒設定密碼：請在 Apps Script 把 API_KEY 改成你的密碼，執行 setup，再部署新版本' });
  if (String(p.key || '').trim() !== key) return out_({ ok: false, error: '密碼錯誤：App 設定頁的密碼要和 Apps Script 裡的 API_KEY 一模一樣（大小寫、半形）。改過 API_KEY 的話，要再「部署 → 管理部署作業 → 新版本」才會生效' });

  const lock = LockService.getScriptLock();
  try {
    if (p.action !== 'list') lock.waitLock(15000);
    switch (p.action) {
      case 'list':    return out_({ ok: true, data: listAll_() });
      case 'save':    return out_({ ok: true, record: save_(p.sheet, p.record) });
      case 'bulk':    return out_(Object.assign({ ok: true }, bulk_(p.items || [])));
      case 'delete':  remove_(p.sheet, p.id); return out_({ ok: true });
      case 'move':    return out_({ ok: true, record: move_(p.from, p.to, p.record) });
      case 'archive': return out_({ ok: true, record: archive_(p.id) });
      default:        return out_({ ok: false, error: '未知的動作：' + p.action });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 資料存取 ---------- */
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(key) {
  if (!SHEET[key]) throw new Error('未知的工作表：' + key);
  const sh = book_().getSheetByName(SHEET[key]);
  if (!sh) throw new Error('找不到工作表「' + SHEET[key] + '」，請先執行 setup');
  return sh;
}

function read_(key) {
  const sh = sheet_(key);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const F = fieldsOf_(key);
  return sh.getRange(2, 1, n - 1, F.length).getDisplayValues()
    .filter(r => r[0])
    .map(r => {
      const o = {};
      F.forEach((f, i) => { o[f] = r[i]; });
      if (OWN_FIELDS[key]) return o;
      try { o['流程'] = o['流程'] ? JSON.parse(o['流程']) : {}; } catch (e) { o['流程'] = {}; }
      return o;
    });
}

function listAll_() {
  const opt = k => book_().getSheetByName(SHEET[k]) ? read_(k) : [];
  return { db: read_('db'), sales: read_('sales'), taxi: read_('taxi'), cash: opt('cash'), pw: opt('pw') };
}

function findRow_(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function save_(key, rec) {
  const sh = sheet_(key);
  const now = stamp_();
  rec = Object.assign({}, rec);
  if (!rec.id) { rec.id = Utilities.getUuid().slice(0, 8); rec['建立時間'] = now; }
  rec['更新時間'] = now;
  const F = fieldsOf_(key);
  const row = F.map(f => {
    const v = rec[f];
    if (v == null) return '';
    return typeof v === 'object' ? JSON.stringify(v) : String(v);
  });
  let r = findRow_(sh, rec.id);
  if (!r) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, F.length).setNumberFormat('@').setValues([row]);
  return rec;
}

/** 一次新增很多筆（匯入用）。同一張訂單（訂單編號）已經存在就略過。 */
function bulk_(items) {
  const now = stamp_();
  const groups = {};
  items.forEach(it => { if (SHEET[it.sheet]) (groups[it.sheet] = groups[it.sheet] || []).push(it.record); });
  const seen = {};
  ['db', 'sales', 'taxi'].forEach(k => read_(k).forEach(r => { if (r['訂單編號']) seen[r['訂單編號']] = true; }));
  let added = 0, skipped = 0;
  Object.keys(groups).forEach(key => {
    const sh = sheet_(key); const F = fieldsOf_(key);
    const rows = [];
    groups[key].forEach(rec => {
      if (rec['訂單編號'] && seen[rec['訂單編號']]) { skipped++; return; }
      if (rec['訂單編號']) seen[rec['訂單編號']] = true;
      rec = Object.assign({}, rec, { id: Utilities.getUuid().slice(0, 8), 建立時間: now, 更新時間: now });
      rows.push(F.map(f => { const v = rec[f]; return v == null ? '' : typeof v === 'object' ? JSON.stringify(v) : String(v); }));
    });
    if (rows.length) {
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, F.length).setNumberFormat('@').setValues(rows);
      added += rows.length;
    }
  });
  return { added: added, skipped: skipped };
}

/** 把一筆資料從一個分頁搬到另一個（例如一般車轉計程車）：先存新的，再刪舊的 */
function move_(from, to, rec) {
  if (!SHEET[from] || !SHEET[to] || from === to) throw new Error('搬移的分頁不對');
  const saved = save_(to, rec);
  remove_(from, rec.id);
  return saved;
}

function remove_(key, id) {
  const sh = sheet_(key);
  const r = findRow_(sh, id);
  if (r) sh.deleteRow(r);
}

function archive_(id) {
  const rec = read_('sales').find(x => x.id === id);
  if (!rec) throw new Error('找不到這台車，可能已經歸檔');
  rec['狀態'] = '已歸檔';
  rec['歸檔時間'] = stamp_();
  const saved = save_('db', rec);
  remove_('sales', id);
  return saved;
}

function stamp_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

/* ---------- 每日提醒（觸發器每天執行） ---------- */
function dailyReminder(force) {
  const now = new Date();
  const y = +Utilities.formatDate(now, TZ, 'yyyy');
  const m = +Utilities.formatDate(now, TZ, 'M');
  const d = +Utilities.formatDate(now, TZ, 'd');
  const lastDay = new Date(y, m, 0).getDate();
  const left = lastDay - d;                 // 0~4 = 月底最後五天
  const parts = [];
  let subject = '';

  if (left <= 4 || force === true) {
    const pending = []
      .concat(read_('sales').map(r => ({ t: '一般', r, done: !!r['流程'][PLATE_STEP.sales] })))
      .concat(read_('taxi').filter(r => r['狀態'] !== '結案')
        .map(r => ({ t: '計程車', r, done: !!(r['流程'][PLATE_STEP.taxi] || r['流程']['領牌資料回傳']) })))
      .filter(x => !x.done);

    subject = left === 0 ? '【今天月底】領牌最後一天' : `【領牌提醒】距月底還有 ${left} 天`;
    if (pending.length) {
      subject += `，${pending.length} 台未領牌`;
      parts.push('<h3>尚未領牌</h3><ul>' + pending.map(x =>
        `<li>[${x.t}] ${esc_(x.r['客戶姓名'] || '未填客戶')}｜${esc_([x.r['廠牌'], x.r['車型']].filter(Boolean).join(' '))}｜${esc_(x.r['月份'])}</li>`
      ).join('') + '</ul>');
    } else {
      subject += '，已全部領牌';
      parts.push('<p>進行中的車都已領牌，沒有待辦。</p>');
    }
  }

  // 每週一：第二年續保
  if (Utilities.formatDate(now, TZ, 'u') === '1' || force === true) {
    const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const year = today.slice(0, 4);
    const c = { done: 0, lost: 0, overdue: 0 };
    const todo = [];
    read_('db').forEach(r => {
      const exp = firstExpiry_(r); if (!exp) return;
      const days = daysBetween_(today, exp); const st = r['續保狀態'];
      if (exp.slice(0, 4) === year) {
        if (st === '已續保') c.done++;
        else if (st === '未續保') c.lost++;
        else if (days < 0) c.overdue++;
      }
      if (!st && days <= RENEW_WARN_DAYS) todo.push({ r, exp, days });
    });
    const total = c.done + c.lost + c.overdue;
    const rate = total ? Math.round(c.done / total * 100) + '%' : '—';
    if (!subject) subject = `【續保提醒】${todo.length} 台要聯絡續保，今年繼續率 ${rate}`;
    parts.push(`<h3>今年第二年繼續率：${rate}</h3><p>應續保 ${total} 台：已續保 ${c.done}、沒有續保 ${c.lost}、過期未處理 ${c.overdue}</p>`);
    if (todo.length) {
      parts.push('<h3>要聯絡續保</h3><ul>' + todo.sort((a, b) => a.days - b.days).map(x =>
        `<li>${esc_(x.r['車牌'] || '無車牌')} ${esc_(x.r['客戶姓名'])}｜到期 ${esc_(x.exp)}（${x.days < 0 ? '已過期 ' + (-x.days) + ' 天' : '剩 ' + x.days + ' 天'}）${x.r['電話'] ? '｜' + esc_(x.r['電話']) : ''}</li>`
      ).join('') + '</ul>');
    }
  }

  if (!parts.length) return;
  MailApp.sendEmail({
    to: REMIND_EMAIL || Session.getEffectiveUser().getEmail(),
    subject: subject,
    htmlBody: parts.join('')
  });
}

/** 手動測試：不管今天幾號都寄一封提醒信 */
function testReminder() { dailyReminder(true); }

function firstExpiry_(r) {
  const ok = v => /^\d{4}-\d{2}-\d{2}$/.test(v || '');
  if (ok(r['任意險到期'])) return r['任意險到期'];
  if (ok(r['強制險到期'])) return r['強制險到期'];
  if (ok(r['投保日期'])) {
    const d = new Date(r['投保日期']); d.setFullYear(d.getFullYear() + 1);
    return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  }
  return '';
}

function daysBetween_(a, b) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b || '')) return null;
  return Math.round((new Date(b) - new Date(a)) / 86400000);
}

function esc_(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
