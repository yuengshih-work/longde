/**
 * 龍德路住宅整修工程｜專案頁後端（Google Apps Script）
 * 資料：Google 試算表（自動建立）；照片與簽名：Google 雲端硬碟資料夾（自動建立）
 *
 * 部署帳號：yuengshih@gmail.com（試算表與照片資料夾都會建立在此帳號的雲端硬碟）
 * 首次使用：
 *   1. 在編輯器上方選擇函式 setup → 執行（授權）
 *   2. 專案設定 → 指令碼屬性，修改 ADMIN_PASS（平安住管理密碼）
 *   3. 部署 → 新增部署作業 → 網頁應用程式；執行身分：我；存取權：所有人
 *   4. 把部署網址貼到 index.html 的 API 常數
 *
 * ── 編輯功能（2026-10-28 新增）──
 *   editUpdate（施工紀錄：改內容、移除／補上傳照片）、editDoc（合約文件：改名稱、移除／補上傳頁面）、editPhoto（照片位置與說明）
 *   日期一律以台灣時區輸出，避免試算表自動轉日期後在網頁上少一天。
 *
 * ── 記帳系統連動（2026-10-21 新增）──
 *   專案頁「標記收款／取消收款」時，自動同步到平安住記帳系統的收款紀錄。
 *   專案設定 → 指令碼屬性，新增：
 *     BOOK_URL         記帳系統 Apps Script 網址（/exec 結尾）
 *     BOOK_TOKEN       pian-site-sync-2026（與記帳系統 SYNC_TOKEN 相同）
 *     BOOK_CONTRACT_ID 記帳系統中此工程的合約ID（在記帳系統「合約」頁可複製；或改填 BOOK_SITE_ID 案場ID）
 *   可選：BOOK_AMOUNTS  各期金額，以逗號分隔（記帳系統未建立付款期程時才會用到）
 *   未設定 BOOK_URL 時完全不影響原本功能。
 *
 * ── 規格及追加工程確認書（線上簽署，2026-11-17 新增）──
 *   kv：confirms（所有確認書：圖面頁、確認項目、業主簽署、平安住核可）、contractorSig（承辦簽名，核可時自動帶入）
 *   可建立多份（封窗、規格確認、追加工程…）；業主免登入勾選全部項目後簽名；
 *   核可時產生「已簽署 PDF」存入雲端硬碟，並自動加入合約文件（追加工程→追加合約，其餘→變更簽認）。
 *   舊版 kv：seal／sealClient／sealApprove 會自動轉入 confirms。
 *   更新後請「管理部署作業 → 編輯 → 版本：新版本 → 部署」，網址不變。
 */

const SHEETS = {
  kv: ['key', 'json', 'updatedAt'],
  photos: ['id', 'cat', 'floor', 'caption', 'fileId', 'updateId', 'createdAt'],
  updates: ['id', 'date', 'stage', 'title', 'note', 'photoIds', 'createdAt'],
  docs: ['id', 'type', 'title', 'date', 'files', 'createdAt'],
  log: ['time', 'action', 'detail']
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SHEET_ID');
  if (!ssId) {
    const ss = SpreadsheetApp.create('龍德路工程專案資料');
    ssId = ss.getId();
    props.setProperty('SHEET_ID', ssId);
  }
  const ss = SpreadsheetApp.openById(ssId);
  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (!props.getProperty('FOLDER_ID')) {
    const folder = DriveApp.createFolder('龍德路工程專案照片');
    props.setProperty('FOLDER_ID', folder.getId());
  }
  if (!props.getProperty('ADMIN_PASS')) props.setProperty('ADMIN_PASS', 'change-me-' + Math.floor(Math.random() * 1e6));
  Logger.log('完成。試算表：https://docs.google.com/spreadsheets/d/' + ssId);
  Logger.log('ADMIN_PASS = ' + props.getProperty('ADMIN_PASS'));
}

/* ---------- helpers ---------- */
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function ss_() { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function sheet_(n) { return ss_().getSheetByName(n); }
function folder_() { return DriveApp.getFolderById(prop_('FOLDER_ID')); }
function out_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return new Date().toISOString(); }
function uid_() { return Utilities.getUuid().slice(0, 8) + Date.now().toString(36); }
function log_(action, detail) { try { sheet_('log').appendRow([now_(), action, String(detail || '').slice(0, 300)]); } catch (e) {} }

function rows_(name) {
  const sh = sheet_(name); const v = sh.getDataRange().getValues(); const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? tw_(r[i], k) : r[i]); return o; });
}
// 試算表會把「2026-10-28」自動轉成日期；以台灣時區輸出才不會少一天
function tw_(d, key) {
  return key === 'date' ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd') : Utilities.formatDate(d, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss'+08:00'");
}
function setText_(name, r, col, v) { sheet_(name).getRange(r, col).setNumberFormat('@').setValue(v == null ? '' : String(v)); }
function photoTrash_(pid) {
  const pr = findRow_('photos', pid);
  if (pr > 0) { trash_(sheet_('photos').getRange(pr, 5).getValue()); sheet_('photos').deleteRow(pr); }
}
function findRow_(name, id) {
  const v = sheet_(name).getRange(1, 1, sheet_(name).getLastRow(), 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 1;
  return -1;
}
function kvGet_(key) {
  const r = findRow_('kv', key); if (r < 0) return null;
  try { return JSON.parse(sheet_('kv').getRange(r, 2).getValue()); } catch (e) { return null; }
}
function kvSet_(key, val) {
  const sh = sheet_('kv'); const r = findRow_('kv', key);
  const row = [key, JSON.stringify(val), now_()];
  if (r < 0) sh.appendRow(row); else sh.getRange(r, 1, 1, 3).setValues([row]);
}
function kvDel_(key) { const r = findRow_('kv', key); if (r > 0) sheet_('kv').deleteRow(r); }

function saveImage_(dataUrl, name) {
  const m = String(dataUrl).match(/^data:(image\/[a-z+]+|application\/pdf);base64,(.+)$/);
  if (!m) throw new Error('檔案格式錯誤（僅接受圖片或 PDF）');
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  const file = folder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}
function trash_(fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {} }


/* ---------- GET ---------- */
function doGet(e) {
  try {
    const isAdmin = (e.parameter.pass || '') === prop_('ADMIN_PASS');
    if (e.parameter.action === 'file') { // 管理端合成已簽署 PDF 用：讀取專案資料夾內的圖片
      if (!isAdmin) throw new Error('管理密碼錯誤');
      const f = DriveApp.getFileById(e.parameter.id);
      if (!f.getParents().hasNext() || f.getParents().next().getId() !== prop_('FOLDER_ID')) throw new Error('檔案不在專案資料夾');
      const b = f.getBlob();
      return out_({ ok: true, data: 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()) });
    }
    return out_({
      confirms: confs_(true),
      contractorSig: isAdmin ? kvGet_('contractorSig') : null,
      ok: true,
      fields: kvGet_('fields') || {},
      docs: rows_('docs').map(d => Object.assign(d, { files: JSON.parse(d.files || '[]') })),
      status: kvGet_('status') || { stages: {}, checks: {}, paid: {} },
      photos: rows_('photos'),
      updates: rows_('updates').map(u => Object.assign(u, { photoIds: String(u.photoIds || '').split(',').filter(Boolean) })),
      admin: isAdmin
    });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- POST ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents || '{}');
    const isAdmin = b.pass && b.pass === prop_('ADMIN_PASS');
    const needAdmin = () => { if (!isAdmin) throw new Error('管理密碼錯誤'); };
    let res = { ok: true };

    switch (b.action) {
      case 'checkAdmin':
        needAdmin(); break;

      case 'saveFields':
        needAdmin();
        kvSet_('fields', { reviewDays: b.reviewDays || '', startDate: b.startDate || '', endDate: b.endDate || '' });
        log_('saveFields', JSON.stringify(b).replace(/"pass":"[^"]*",?/, ''));
        break;

      case 'addDoc': {
        needAdmin(); const id = uid_();
        sheet_('docs').appendRow([id, b.type || '', b.title || '', b.date || '', '[]', now_()]);
        res.id = id; log_('addDoc', b.title); break;
      }
      case 'uploadDocFile': {
        needAdmin(); const r = findRow_('docs', b.docId); if (r < 0) throw new Error('找不到文件');
        const fileId = saveImage_(b.data, (b.name || 'doc') );
        const cell = sheet_('docs').getRange(r, 5);
        const files = JSON.parse(cell.getValue() || '[]');
        files.push({ id: fileId, mime: String(b.data).slice(5, String(b.data).indexOf(';')), name: b.name || '' });
        cell.setValue(JSON.stringify(files)); res.fileId = fileId; break;
      }
      case 'deleteDoc': {
        needAdmin(); const r = findRow_('docs', b.id);
        if (r > 0) { JSON.parse(sheet_('docs').getRange(r, 5).getValue() || '[]').forEach(f => trash_(f.id)); sheet_('docs').deleteRow(r); }
        log_('deleteDoc', b.id); break;
      }
      case 'saveStatus': {
        needAdmin();
        const before = (kvGet_('status') || {}).paid || {};
        kvSet_('status', { stages: b.stages || {}, checks: b.checks || {}, paid: b.paid || {} });
        log_('saveStatus', '');
        if (JSON.stringify(before) !== JSON.stringify(b.paid || {})) res.book = pushPaidToBook_(b.paid || {});
        break;
      }

      case 'uploadPhoto': {
        needAdmin();
        const id = uid_();
        const fileId = saveImage_(b.data, (b.cat || 'photo') + '_' + (b.floor || '') + '_' + id + '.jpg');
        sheet_('photos').appendRow([id, b.cat || 'before', b.floor || '', b.caption || '', fileId, b.updateId || '', now_()]);
        res.id = id; res.fileId = fileId; break;
      }
      case 'deletePhoto': {
        needAdmin(); const r = findRow_('photos', b.id);
        if (r > 0) { trash_(sheet_('photos').getRange(r, 5).getValue()); sheet_('photos').deleteRow(r); }
        break;
      }
      case 'addUpdate': {
        needAdmin(); const id = uid_();
        sheet_('updates').appendRow([id, b.date || '', b.stage || '', b.title || '', b.note || '', (b.photoIds || []).join(','), now_()]);
        res.id = id; log_('addUpdate', b.title); break;
      }
      case 'deleteUpdate': {
        needAdmin(); const r = findRow_('updates', b.id);
        if (r > 0) {
          const ids = String(sheet_('updates').getRange(r, 6).getValue() || '').split(',').filter(Boolean);
          sheet_('updates').deleteRow(r);
          ids.forEach(pid => { const pr = findRow_('photos', pid); if (pr > 0) { trash_(sheet_('photos').getRange(pr, 5).getValue()); sheet_('photos').deleteRow(pr); } });
        }
        break;
      }
      case 'editUpdate': {
        needAdmin(); const r = findRow_('updates', b.id); if (r < 0) throw new Error('找不到這筆紀錄');
        const sh = sheet_('updates'); const rm = b.removePhotoIds || [];
        const ids = String(sh.getRange(r, 6).getValue() || '').split(',').filter(Boolean).filter(x => rm.indexOf(x) < 0).concat(b.addPhotoIds || []);
        rm.forEach(photoTrash_);
        const r2 = findRow_('updates', b.id);
        setText_('updates', r2, 2, b.date || '');
        sh.getRange(r2, 3, 1, 4).setValues([[b.stage || '', b.title || '', b.note || '', ids.join(',')]]);
        log_('editUpdate', b.title); break;
      }
      case 'editDoc': {
        needAdmin(); const r = findRow_('docs', b.id); if (r < 0) throw new Error('找不到文件');
        const sh = sheet_('docs'); const rm = b.removeFileIds || [];
        const files = JSON.parse(sh.getRange(r, 5).getValue() || '[]');
        rm.forEach(trash_);
        sh.getRange(r, 2, 1, 2).setValues([[b.type || '', b.title || '']]);
        setText_('docs', r, 4, b.date || '');
        sh.getRange(r, 5).setValue(JSON.stringify(files.filter(f => rm.indexOf(f.id) < 0)));
        log_('editDoc', b.title); break;
      }
      case 'editPhoto': {
        needAdmin(); const r = findRow_('photos', b.id); if (r < 0) throw new Error('找不到照片');
        if (b.floor != null) sheet_('photos').getRange(r, 3).setValue(b.floor);
        if (b.caption != null) sheet_('photos').getRange(r, 4).setValue(b.caption);
        break;
      }
      /* ---------- 規格及追加工程確認書 ---------- */
      case 'confPublish': { // 新增或發布新版
        needAdmin();
        const list = confs_(); let c = b.id ? list.find(x => x.id === b.id) : null;
        if (b.id && !c) throw new Error('找不到這份確認書');
        const items = (b.items || []).map(t => String(t || '').trim()).filter(Boolean).map((t, i) => ({ n: i + 1, text: t.slice(0, 300) }));
        if (!items.length) throw new Error('請至少列出一個確認項目');
        const title = String(b.title || '').trim(); if (!title) throw new Error('請填寫標題');
        if (!c) { c = { id: uid_(), ver: 0, pages: [], client: null, approve: null, createdAt: now_() }; list.push(c); }
        const ver = c.ver + 1;
        let pages = c.pages || [];
        if (b.pages && b.pages.length) {
          pages = b.pages.map((d, i) => saveImage_(d, title + '_v' + ver + '_p' + (i + 1) + '.jpg'));
          (c.pages || []).forEach(id => { if (!c.client && !c.approve) trash_(id); }); // 未簽過的舊圖直接清掉
        }
        if (!pages.length) throw new Error('請上傳圖面或附件（PDF 或圖片）');
        const signed = c.client && c.client.ver === c.ver;
        if (c.approve) unapprove_(c);
        Object.assign(c, { ver: ver, type: b.type || '規格確認', title: title.slice(0, 80), amount: String(b.amount || '').trim().slice(0, 40),
          statement: String(b.statement || '').trim().slice(0, 300), items: items, pages: pages, publishedAt: now_() });
        if (signed) c.client = Object.assign({}, c.client, { stale: true });
        saveConfs_(list); res.id = c.id; log_('confPublish', c.title + ' v' + ver); break;
      }
      case 'confDelete': {
        needAdmin(); const list = confs_(); const i = list.findIndex(x => x.id === b.id); if (i < 0) break;
        const c = list[i]; if (c.approve) throw new Error('已核可的確認書請先撤銷核可再刪除');
        list.splice(i, 1); saveConfs_(list); log_('confDelete', c.title); break;
      }
      case 'confSign': { // 業主：免登入
        const list = confs_(); const c = list.find(x => x.id === b.id); if (!c) throw new Error('找不到這份確認書');
        if (Number(b.ver) !== c.ver) throw new Error('確認書已更新，請重新整理頁面後再簽署');
        if (c.client && c.client.ver === c.ver && !c.client.stale) throw new Error('此版已完成簽署');
        const name = String(b.name || '').trim(); if (!name) throw new Error('請填寫姓名');
        const checked = (b.checked || []).map(Number);
        if (!c.items.every(it => checked.indexOf(it.n) >= 0)) throw new Error('請勾選全部確認項目');
        if (!/^data:image\/png;base64,/.test(String(b.signature || ''))) throw new Error('簽名格式錯誤');
        if (c.approve) unapprove_(c);
        c.client = { name: name.slice(0, 40), sigFileId: saveImage_(b.signature, c.title + '_業主簽名_v' + c.ver + '_' + uid_() + '.png'), signedAt: now_(), ver: c.ver, checked: checked };
        saveConfs_(list); log_('confSign', c.title + ' ' + name + ' v' + c.ver); break;
      }
      case 'confReset': {
        needAdmin(); const list = confs_(); const c = list.find(x => x.id === b.id); if (!c) break;
        if (c.approve) unapprove_(c); c.client = null; saveConfs_(list); log_('confReset', c.title); break;
      }
      case 'confApprove': {
        needAdmin();
        const list = confs_(); const c = list.find(x => x.id === b.id); if (!c) throw new Error('找不到這份確認書');
        if (!c.client || c.client.ver !== c.ver || c.client.stale) throw new Error('業主尚未依目前版本簽署');
        if (b.clientSignedAt !== c.client.signedAt) throw new Error('業主簽署資料已變動，請重新整理');
        let k = kvGet_('contractorSig');
        if (b.signature) { k = { fileId: saveImage_(b.signature, '平安住承辦簽名_' + uid_() + '.png'), rep: b.rep || '' }; kvSet_('contractorSig', k); }
        if (!k) throw new Error('請先簽名');
        if (b.rep && b.rep !== k.rep) { k.rep = b.rep; kvSet_('contractorSig', k); }
        if (c.approve) unapprove_(c);
        const name = c.title + '_已簽署_第' + c.ver + '版.pdf';
        const pdfFileId = saveImage_(b.pdf, name);
        const docId = uid_(); const day = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
        sheet_('docs').appendRow([docId, c.type === '追加工程' ? '追加合約' : '變更簽認', c.title + '（第' + c.ver + '版・已簽署）', '', JSON.stringify([{ id: pdfFileId, mime: 'application/pdf', name: name }]), now_()]);
        setText_('docs', findRow_('docs', docId), 4, day);
        c.approve = { rep: k.rep, sigFileId: k.fileId, approvedAt: b.approvedAt || now_(), ver: c.ver, clientSignedAt: c.client.signedAt, pdfFileId: pdfFileId, docId: docId };
        saveConfs_(list); log_('confApprove', c.title + ' ' + k.rep + ' v' + c.ver); break;
      }
      case 'confUnapprove': {
        needAdmin(); const list = confs_(); const c = list.find(x => x.id === b.id); if (!c) break;
        unapprove_(c); saveConfs_(list); log_('confUnapprove', c.title); break;
      }
      default:
        throw new Error('未知的操作');
    }
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}


function unapprove_(c) { // c：確認書物件（會直接修改，呼叫端負責存檔）
  const a = c.approve; if (!a) return;
  if (a.docId) { const r = findRow_('docs', a.docId); if (r > 0) sheet_('docs').deleteRow(r); }
  if (a.pdfFileId) trash_(a.pdfFileId);
  c.approve = null;
}
function saveConfs_(list) { kvSet_('confirms', list); }
function confs_(needLock) { // doPost 已持有鎖；doGet 轉換舊資料時另外取鎖
  const list = kvGet_('confirms'); if (list) return list;
  const s = kvGet_('seal'); if (!s) return [];
  // 舊版封窗確認書轉入
  const c = { id: 'seal', ver: s.ver, type: '規格確認', title: s.title || '外觀立面封窗確認書', amount: '', statement: '本人已確認上列封窗及保留窗位置，同意依本圖施工。',
    items: (s.items || []).map(it => ({ n: it.n, text: [it.floor, it.face, it.note].filter(Boolean).join(' ') })),
    pages: [s.fileId], publishedAt: s.publishedAt, createdAt: s.publishedAt, client: kvGet_('sealClient'), approve: kvGet_('sealApprove') };
  const migrate = () => { if (!kvGet_('confirms')) { kvSet_('confirms', [c]); kvDel_('seal'); kvDel_('sealClient'); kvDel_('sealApprove'); } };
  if (!needLock) migrate();
  else { const lock = LockService.getScriptLock(); if (lock.tryLock(5000)) { try { migrate(); } finally { lock.releaseLock(); } } }
  return [c];
}

/* ---------- 記帳系統連動：收款狀態 ---------- */
function pushPaidToBook_(paid) {
  const url = prop_('BOOK_URL');
  if (!url) return { skipped: true };
  try {
    const amounts = String(prop_('BOOK_AMOUNTS') || '').split(',').map(x => Number(String(x).trim()) || 0).filter(x => x > 0);
    const r = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, followRedirects: true,
      payload: JSON.stringify({
        action: 'projectPayment', token: prop_('BOOK_TOKEN') || '',
        contractId: prop_('BOOK_CONTRACT_ID') || '', siteId: prop_('BOOK_SITE_ID') || '',
        paid: paid, amounts: amounts
      })
    });
    const j = JSON.parse(r.getContentText());
    log_('bookSync', JSON.stringify(j).slice(0, 280));
    return j;
  } catch (e) {
    log_('bookSyncError', String(e.message || e));
    return { ok: false, error: String(e.message || e) };
  }
}
