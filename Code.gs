/**
 * Cohen Expense Flow — Apps Script backend
 * -----------------------------------------
 * One project does everything:
 *   - serves the mobile PWA UI (doGet -> index.html)
 *   - stores receipt images in Google Drive (runs as YOU, no service account)
 *   - writes the ledger to Google Sheets
 *   - calls OpenAI vision for receipt extraction (key stays server-side)
 *
 * First-time setup: open the editor, run `setup()` once, grant permissions.
 * Then set Script Properties:  PASSCODE, OPENAI_API_KEY
 * Deploy: Deploy > New deployment > Web app > Execute as: Me, Access: Anyone.
 */

/* ------------------------------------------------------------------ *
 * Config
 * ------------------------------------------------------------------ */

const ROOT_FOLDER_NAME = 'Cohen Expense Flow';
const SHEET_NAME       = 'Cohen Expense Flow Ledger';
const DEFAULT_PEOPLE   = ['Leah', 'Moshe', 'Other'];
const DEFAULT_ORGS     = ['Hamsa Nomads', 'Moshe House', 'Personal', 'Work', 'Other'];

const TABS = {
  Organizations:  ['organization_id','organization_name','status','notes','created_at'],
  EventGroups:    ['event_group_id','organization_id','organization_name','event_group_name','status','notes','created_at'],
  Events:         ['event_id','organization_id','organization_name','event_group_id','event_group_name','event_name','event_date','status','budget','notes','created_at'],
  Receipts:       ['receipt_id','date','organization_id','organization_name','event_group_id','event_group_name','event_id','event_name','paid_by','merchant','subtotal','tax','tip','total','currency','payment_method','drive_file_id','drive_url','status','ai_confidence','created_at','approved_at','notes'],
  Items:          ['item_id','receipt_id','date','organization_name','event_group_name','event_name','paid_by','merchant','item_name','normalized_item_name','quantity','unit','item_total','unit_price','category','needed_for','kosher_sensitive','reimbursable','ai_confidence','notes'],
  PeopleBalances: ['event_id','event_name','person','total_paid','should_pay','reimbursement_due','last_updated'],
  PriceChecks:    ['price_check_id','item_id','receipt_id','event_name','item_name','original_store','original_price','original_unit_price','cheaper_store','cheaper_price','cheaper_unit_price','savings','link','kosher_status_checked','confidence','status','checked_at','notes'],
  AuditLog:       ['timestamp','user','action','receipt_id','event_id','details'],
  _Config:        ['key','value'],
  _Folders:       ['path','folder_id']
};

function props_()  { return PropertiesService.getScriptProperties(); }
function prop_(k)  { return props_().getProperty(k); }
function setProp_(k,v){ props_().setProperty(k, String(v)); }

/* ------------------------------------------------------------------ *
 * Web entry point — serves the UI
 * ------------------------------------------------------------------ */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cohen Expense Flow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ------------------------------------------------------------------ *
 * Auth gate — every server call validates the passcode
 * ------------------------------------------------------------------ */

function checkPass_(passcode) {
  const expected = prop_('PASSCODE');
  if (!expected) throw new Error('App not configured: set PASSCODE in Script Properties.');
  if (String(passcode) !== String(expected)) throw new Error('Access not allowed.');
}

/** Lightweight login check used by the client. */
function authenticate(passcode) {
  checkPass_(passcode);
  return { ok: true };
}

/* ------------------------------------------------------------------ *
 * One-time setup
 * ------------------------------------------------------------------ */

function setup() {
  const root = getOrCreateChildFolder_(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  setProp_('ROOT_FOLDER_ID', root.getId());
  ['Receipts','Exports','Logs'].forEach(n => getOrCreateChildFolder_(root, n));

  let ss;
  const existingId = prop_('SHEET_ID');
  if (existingId) {
    try { ss = SpreadsheetApp.openById(existingId); } catch (e) { ss = null; }
  }
  if (!ss) { ss = SpreadsheetApp.create(SHEET_NAME); setProp_('SHEET_ID', ss.getId()); }

  Object.keys(TABS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(TABS[name]);
    if (name.charAt(0) === '_') sh.hideSheet();
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  // Seed organizations + people if empty
  const orgSheet = ss.getSheetByName('Organizations');
  if (orgSheet.getLastRow() === 1) {
    DEFAULT_ORGS.forEach((name, i) => {
      orgSheet.appendRow(['ORG-' + pad_(i + 1, 4), name, 'active', '', new Date()]);
    });
  }
  setConfig_('people', DEFAULT_PEOPLE.join(','));
  if (!getConfig_('next_receipt')) setConfig_('next_receipt', '1');

  return {
    sheetUrl: ss.getUrl(),
    rootFolderUrl: root.getUrl(),
    rootFolderId: root.getId(),
    sheetId: ss.getId(),
    note: 'Now set PASSCODE and OPENAI_API_KEY in Project Settings > Script Properties.'
  };
}

/* ------------------------------------------------------------------ *
 * Sheet helpers
 * ------------------------------------------------------------------ */

function ss_()   { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function tab_(n) { return ss_().getSheetByName(n); }

function read_(name) {
  const sh = tab_(name);
  const values = sh.getDataRange().getValues();
  const headers = values.shift() || [];
  const rows = values.map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
  return { sh, headers, rows };
}

function append_(name, obj) {
  const sh = tab_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => (obj[h] === undefined || obj[h] === null) ? '' : obj[h]));
}

function findRowIndex_(name, idCol, idVal) {
  const { headers, sh } = read_(name);
  const col = headers.indexOf(idCol);
  if (col < 0) return -1;
  const vals = sh.getRange(2, col + 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(idVal)) return i + 2;
  return -1;
}

function updateRow_(name, rowIndex, patch) {
  const sh = tab_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  headers.forEach((h, i) => { if (patch[h] !== undefined) row[i] = patch[h]; });
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function getConfig_(key) {
  const { rows } = read_('_Config');
  const r = rows.find(x => x.key === key);
  return r ? r.value : null;
}
function setConfig_(key, value) {
  const idx = findRowIndex_('_Config', 'key', key);
  if (idx > 0) updateRow_('_Config', idx, { value: value });
  else append_('_Config', { key: key, value: value });
}

function pad_(n, w) { return String(n).padStart(w, '0'); }

function nextReceiptId_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const n = parseInt(getConfig_('next_receipt') || '1', 10);
    setConfig_('next_receipt', String(n + 1));
    return 'R-' + pad_(n, 6);
  } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------------------ *
 * Drive helpers
 * ------------------------------------------------------------------ */

function getOrCreateChildFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** Resolve /Receipts/Org/Group/Event, caching ids in _Folders. */
function folderForReceipt_(orgName, groupName, eventName) {
  const segments = ['Receipts', orgName, groupName, eventName].filter(Boolean);
  const path = segments.join('/');
  const cached = getFolderCache_(path);
  if (cached) { try { return DriveApp.getFolderById(cached); } catch (e) {} }

  let folder = DriveApp.getFolderById(prop_('ROOT_FOLDER_ID'));
  segments.forEach(seg => { folder = getOrCreateChildFolder_(folder, seg); });
  setFolderCache_(path, folder.getId());
  return folder;
}

function getFolderCache_(path) {
  const { rows } = read_('_Folders');
  const r = rows.find(x => x.path === path);
  return r ? r.folder_id : null;
}
function setFolderCache_(path, id) {
  const idx = findRowIndex_('_Folders', 'path', path);
  if (idx > 0) updateRow_('_Folders', idx, { folder_id: id });
  else append_('_Folders', { path: path, folder_id: id });
}

function token_(s) {
  return String(s == null ? '' : s)
    .replace(/[^A-Za-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '') || 'NA';
}
function totalToken_(t) {
  if (t == null || t === '') return 'NA';
  return String(t).replace(/\./g, '-');
}

/* ------------------------------------------------------------------ *
 * Client-facing API (called via google.script.run)
 * ------------------------------------------------------------------ */

/** Everything the add-receipt form needs. */
function bootstrap(passcode) {
  checkPass_(passcode);
  const orgs   = read_('Organizations').rows.filter(r => r.organization_id);
  const groups = read_('EventGroups').rows.filter(r => r.event_group_id);
  const events = read_('Events').rows.filter(r => r.event_id);
  const people = (getConfig_('people') || DEFAULT_PEOPLE.join(',')).split(',').map(s => s.trim());
  return {
    organizations: orgs.map(o => ({ id: o.organization_id, name: o.organization_name, status: o.status })),
    eventGroups:   groups.map(g => ({ id: g.event_group_id, orgId: g.organization_id, orgName: g.organization_name, name: g.event_group_name })),
    events:        events.map(e => ({
      id: e.event_id, orgId: e.organization_id, orgName: e.organization_name,
      groupId: e.event_group_id, groupName: e.event_group_name, name: e.event_name,
      date: e.event_date ? Utilities.formatDate(new Date(e.event_date), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      status: e.status, budget: e.budget
    })),
    people: people
  };
}

/** Create org/group/event on demand. Returns the new event. */
function createEvent(passcode, p) {
  checkPass_(passcode);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const now = new Date();

    // org
    let org = read_('Organizations').rows.find(o => o.organization_name === p.organizationName);
    if (!org) {
      const id = 'ORG-' + pad_(read_('Organizations').rows.length + 1, 4);
      append_('Organizations', { organization_id: id, organization_name: p.organizationName, status: 'active', created_at: now });
      org = { organization_id: id, organization_name: p.organizationName };
    }
    // group
    let grp = read_('EventGroups').rows.find(g => g.event_group_name === p.eventGroupName && g.organization_id === org.organization_id);
    if (!grp) {
      const id = 'EG-' + pad_(read_('EventGroups').rows.length + 1, 4);
      append_('EventGroups', { event_group_id: id, organization_id: org.organization_id, organization_name: org.organization_name, event_group_name: p.eventGroupName, status: 'active', created_at: now });
      grp = { event_group_id: id, event_group_name: p.eventGroupName };
    }
    // event
    const evId = 'EV-' + pad_(read_('Events').rows.length + 1, 4);
    append_('Events', {
      event_id: evId, organization_id: org.organization_id, organization_name: org.organization_name,
      event_group_id: grp.event_group_id, event_group_name: grp.event_group_name,
      event_name: p.eventName, event_date: p.eventDate || '', status: 'active',
      budget: p.budget || '', created_at: now
    });
    audit_('create_event', '', evId, p.eventName);
    return { id: evId, orgId: org.organization_id, orgName: org.organization_name, groupId: grp.event_group_id, groupName: grp.event_group_name, name: p.eventName, status: 'active' };
  } finally { lock.releaseLock(); }
}

/**
 * Upload a receipt image, create a draft row, run AI extraction.
 * p = { imageBase64, mimeType, event:{id,name,orgId,orgName,groupId,groupName}, paidBy, note }
 */
function uploadReceipt(passcode, p) {
  checkPass_(passcode);
  const receiptId = nextReceiptId_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // 1. Save image to Drive with a temporary name
  const blob = Utilities.newBlob(Utilities.base64Decode(p.imageBase64), p.mimeType || 'image/jpeg',
    'PENDING__' + today + '__' + token_(p.paidBy) + '__' + receiptId + '.jpg');
  const folder = folderForReceipt_(p.event.orgName, p.event.groupName, p.event.name);
  const file = folder.createFile(blob);
  const fileId = file.getId();
  const url = file.getUrl();

  // 2. Draft row (status needs_review until a human approves)
  append_('Receipts', {
    receipt_id: receiptId, date: today,
    organization_id: p.event.orgId, organization_name: p.event.orgName,
    event_group_id: p.event.groupId, event_group_name: p.event.groupName,
    event_id: p.event.id, event_name: p.event.name,
    paid_by: p.paidBy, currency: 'USD',
    drive_file_id: fileId, drive_url: url,
    status: 'needs_review', created_at: new Date(), notes: p.note || ''
  });
  audit_('upload', receiptId, p.event.id, p.event.name);

  // 3. AI extraction
  let extraction;
  try { extraction = extractWithOpenAI_(p.imageBase64, p.mimeType || 'image/jpeg'); }
  catch (err) { extraction = blankExtraction_('AI error: ' + err.message); }

  // 4. Patch row + rename file when merchant/total are known
  const idx = findRowIndex_('Receipts', 'receipt_id', receiptId);
  updateRow_('Receipts', idx, {
    merchant: extraction.merchant_name || '',
    subtotal: numOrBlank_(extraction.subtotal),
    tax:      numOrBlank_(extraction.tax),
    tip:      numOrBlank_(extraction.tip),
    total:    numOrBlank_(extraction.total),
    payment_method: extraction.payment_method || '',
    ai_confidence:  extraction.overall_confidence || 'low'
  });

  if (extraction.merchant_name || extraction.total != null) {
    const finalName = [today, token_(p.event.orgName), token_(p.event.groupName), token_(p.event.name),
      token_(p.paidBy), token_(extraction.merchant_name), totalToken_(extraction.total), receiptId].join('__') + '.jpg';
    try { file.setName(finalName); } catch (e) {}
  }

  return {
    receiptId: receiptId,
    driveFileId: fileId,
    driveUrl: url,
    status: 'needs_review',
    extraction: extraction,
    paidBy: p.paidBy,
    date: today,
    event: p.event
  };
}

/** Re-run AI on an already-uploaded receipt. */
function rerunAI(passcode, receiptId) {
  checkPass_(passcode);
  const { rows } = read_('Receipts');
  const r = rows.find(x => x.receipt_id === receiptId);
  if (!r) throw new Error('Receipt not found.');
  const file = DriveApp.getFileById(r.drive_file_id);
  const b64 = Utilities.base64Encode(file.getBlob().getBytes());
  const extraction = extractWithOpenAI_(b64, file.getBlob().getContentType());
  const idx = findRowIndex_('Receipts', 'receipt_id', receiptId);
  updateRow_('Receipts', idx, {
    merchant: extraction.merchant_name || '', subtotal: numOrBlank_(extraction.subtotal),
    tax: numOrBlank_(extraction.tax), tip: numOrBlank_(extraction.tip),
    total: numOrBlank_(extraction.total), ai_confidence: extraction.overall_confidence || 'low'
  });
  return { extraction: extraction, driveUrl: r.drive_url };
}

/** Load a receipt for the review screen (draft or saved). */
function getReceipt(passcode, receiptId) {
  checkPass_(passcode);
  const r = read_('Receipts').rows.find(x => x.receipt_id === receiptId);
  if (!r) throw new Error('Receipt not found.');
  const items = read_('Items').rows.filter(i => i.receipt_id === receiptId);
  return { receipt: r, items: items };
}

/**
 * Approve + save (or save as draft).
 * p = { receiptId, merchant, date, subtotal, tax, tip, total, paidBy, event{...},
 *       items:[{name,normalized_name,quantity,unit,total_price,unit_price,category,needed_for,kosher_sensitive,reimbursable,confidence}],
 *       notes, approve:bool }
 */
function saveReceipt(passcode, p) {
  checkPass_(passcode);
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    if (p.approve && (!p.event || !p.event.id)) throw new Error('Pick an event before approving.');

    const idx = findRowIndex_('Receipts', 'receipt_id', p.receiptId);
    if (idx < 0) throw new Error('Receipt not found.');

    updateRow_('Receipts', idx, {
      merchant: p.merchant || '', date: p.date || '',
      subtotal: numOrBlank_(p.subtotal), tax: numOrBlank_(p.tax), tip: numOrBlank_(p.tip), total: numOrBlank_(p.total),
      paid_by: p.paidBy || '',
      organization_id: p.event.orgId, organization_name: p.event.orgName,
      event_group_id: p.event.groupId, event_group_name: p.event.groupName,
      event_id: p.event.id, event_name: p.event.name,
      status: p.approve ? 'approved' : 'draft',
      approved_at: p.approve ? new Date() : '',
      notes: p.notes || ''
    });

    // Rewrite item rows for this receipt
    deleteItemsFor_(p.receiptId);
    (p.items || []).forEach((it, i) => {
      append_('Items', {
        item_id: 'IT-' + p.receiptId + '-' + pad_(i + 1, 3),
        receipt_id: p.receiptId, date: p.date || '',
        organization_name: p.event.orgName, event_group_name: p.event.groupName, event_name: p.event.name,
        paid_by: p.paidBy || '', merchant: p.merchant || '',
        item_name: it.name || '', normalized_item_name: it.normalized_name || it.name || '',
        quantity: numOrBlank_(it.quantity), unit: it.unit || '',
        item_total: numOrBlank_(it.total_price), unit_price: numOrBlank_(it.unit_price),
        category: it.category || 'other', needed_for: it.needed_for || '',
        kosher_sensitive: !!it.kosher_sensitive, reimbursable: it.reimbursable !== false,
        ai_confidence: it.confidence || 'medium', notes: ''
      });
    });

    if (p.approve) {
      recomputeBalances_(p.event.id, p.event.name);
      audit_('approve', p.receiptId, p.event.id, p.merchant || '');
    } else {
      audit_('save_draft', p.receiptId, p.event.id, '');
    }
    return { ok: true, status: p.approve ? 'approved' : 'draft' };
  } finally { lock.releaseLock(); }
}

function deleteReceipt(passcode, receiptId) {
  checkPass_(passcode);
  const idx = findRowIndex_('Receipts', 'receipt_id', receiptId);
  if (idx < 0) throw new Error('Receipt not found.');
  const r = read_('Receipts').rows.find(x => x.receipt_id === receiptId);
  updateRow_('Receipts', idx, { status: 'rejected', notes: 'Deleted by user' });
  deleteItemsFor_(receiptId);
  if (r && r.event_id) recomputeBalances_(r.event_id, r.event_name);
  audit_('delete', receiptId, r ? r.event_id : '', '');
  return { ok: true };
}

function deleteItemsFor_(receiptId) {
  const sh = tab_('Items');
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][1]) === String(receiptId)) sh.deleteRow(i + 1);
}

/* ------------------------------------------------------------------ *
 * Dashboards
 * ------------------------------------------------------------------ */

function reviewQueue(passcode) {
  checkPass_(passcode);
  const open = ['needs_review', 'draft', 'error'];
  return read_('Receipts').rows
    .filter(r => open.indexOf(r.status) >= 0)
    .map(r => ({ receiptId: r.receipt_id, merchant: r.merchant, total: r.total, eventName: r.event_name, paidBy: r.paid_by, status: r.status, date: fmtDate_(r.date), driveUrl: r.drive_url }));
}

function listEvents(passcode) {
  checkPass_(passcode);
  const events = read_('Events').rows.filter(e => e.event_id);
  const receipts = read_('Receipts').rows.filter(r => r.status === 'approved');
  return events.map(e => {
    const rs = receipts.filter(r => r.event_id === e.event_id);
    const total = sum_(rs.map(r => num_(r.total)));
    const byPerson = {};
    rs.forEach(r => { byPerson[r.paid_by] = (byPerson[r.paid_by] || 0) + num_(r.total); });
    return {
      id: e.event_id, name: e.event_name, orgName: e.organization_name, groupName: e.event_group_name,
      status: e.status, budget: num_(e.budget), total: total, receiptCount: rs.length,
      paidByLeah: byPerson['Leah'] || 0, paidByMoshe: byPerson['Moshe'] || 0
    };
  });
}

function getEvent(passcode, eventId) {
  checkPass_(passcode);
  const e = read_('Events').rows.find(x => x.event_id === eventId);
  if (!e) throw new Error('Event not found.');
  const receipts = read_('Receipts').rows.filter(r => r.event_id === eventId && r.status === 'approved');
  const items = read_('Items').rows.filter(i => i.event_name === e.event_name);
  const balances = read_('PeopleBalances').rows.filter(b => b.event_id === eventId);

  const byCat = {};
  items.forEach(i => { byCat[i.category || 'other'] = (byCat[i.category || 'other'] || 0) + num_(i.item_total); });
  const byPerson = {};
  receipts.forEach(r => { byPerson[r.paid_by] = (byPerson[r.paid_by] || 0) + num_(r.total); });

  return {
    event: { id: e.event_id, name: e.event_name, orgName: e.organization_name, groupName: e.event_group_name, budget: num_(e.budget), status: e.status },
    total: sum_(receipts.map(r => num_(r.total))),
    byCategory: byCat,
    byPerson: byPerson,
    balances: balances.map(b => ({ person: b.person, paid: num_(b.total_paid), shouldPay: num_(b.should_pay), due: num_(b.reimbursement_due) })),
    receipts: receipts.map(r => ({ receiptId: r.receipt_id, merchant: r.merchant, total: num_(r.total), paidBy: r.paid_by, date: fmtDate_(r.date), driveUrl: r.drive_url })),
    items: items.map(i => ({ name: i.item_name, total: num_(i.item_total), category: i.category, kosher: i.kosher_sensitive }))
  };
}

/* ------------------------------------------------------------------ *
 * Balances + audit
 * ------------------------------------------------------------------ */

function recomputeBalances_(eventId, eventName) {
  const receipts = read_('Receipts').rows.filter(r => r.event_id === eventId && r.status === 'approved');
  const paid = {};
  receipts.forEach(r => { if (r.paid_by) paid[r.paid_by] = (paid[r.paid_by] || 0) + num_(r.total); });

  // Equal split among contributors who aren't "Other"
  const splitters = Object.keys(paid).filter(p => p !== 'Other');
  const eventTotal = sum_(receipts.map(r => num_(r.total)));
  const share = splitters.length ? eventTotal / splitters.length : 0;

  // Clear old balance rows for this event, then write fresh
  const sh = tab_('PeopleBalances');
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) if (String(data[i][0]) === String(eventId)) sh.deleteRow(i + 1);

  Object.keys(paid).forEach(person => {
    const shouldPay = person === 'Other' ? 0 : share;
    append_('PeopleBalances', {
      event_id: eventId, event_name: eventName, person: person,
      total_paid: round2_(paid[person]), should_pay: round2_(shouldPay),
      reimbursement_due: round2_(paid[person] - shouldPay), last_updated: new Date()
    });
  });
}

function audit_(action, receiptId, eventId, details) {
  let user = '';
  try { user = Session.getActiveUser().getEmail(); } catch (e) {}
  append_('AuditLog', { timestamp: new Date(), user: user, action: action, receipt_id: receiptId || '', event_id: eventId || '', details: details || '' });
}

/* ------------------------------------------------------------------ *
 * OpenAI extraction
 * ------------------------------------------------------------------ */

const RECEIPT_CATEGORIES = ['food','supplies','marketing','travel','housing','utilities','decor','printing','other'];

function extractWithOpenAI_(base64, mimeType) {
  const key = prop_('OPENAI_API_KEY');
  if (!key) return blankExtraction_('OPENAI_API_KEY not set.');
  const model = prop_('OPENAI_MODEL') || 'gpt-4o-mini';

  const system =
    'You read a single retail/restaurant receipt image and return STRICT JSON only — no prose, no markdown. ' +
    'Schema: {merchant_name:string|null, date:string|null (YYYY-MM-DD), subtotal:number|null, tax:number|null, ' +
    'tip:number|null, total:number|null, currency:"USD", payment_method:string|null, items:[{name:string, ' +
    'normalized_name:string|null, quantity:number|null, unit:string|null, total_price:number|null, ' +
    'unit_price:number|null, category:one of ' + RECEIPT_CATEGORIES.join('|') + ', needed_for:string|null, ' +
    'kosher_sensitive:boolean, reimbursable:boolean, confidence:"low"|"medium"|"high"}], warnings:string[], ' +
    'overall_confidence:"low"|"medium"|"high"}. ' +
    'Rules: preserve original item names; also give a clean normalized_name. Never invent prices — use null if unreadable. ' +
    'If the image is blurry or the total is unreadable, add a warning and use null. If a category is uncertain use "other" with low confidence. ' +
    'Mark kosher_sensitive=true for any food where kosher status could matter (meat, dairy, prepared/packaged food). Default reimbursable=true.';

  const payload = {
    model: model,
    temperature: 0,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: system },
      { role: 'user', content: [
        { type: 'text', text: 'Extract this receipt as JSON per the schema.' },
        { type: 'image_url', image_url: { url: 'data:' + mimeType + ';base64,' + base64 } }
      ]}
    ]
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) return blankExtraction_('OpenAI HTTP ' + res.getResponseCode());

  let content = '';
  try { content = JSON.parse(res.getContentText()).choices[0].message.content; }
  catch (e) { return blankExtraction_('Bad OpenAI response.'); }

  return coerceExtraction_(content);
}

function coerceExtraction_(raw) {
  let obj;
  try { obj = JSON.parse(String(raw).replace(/```json|```/g, '').trim()); }
  catch (e) { return blankExtraction_('Could not parse AI JSON.'); }

  const items = Array.isArray(obj.items) ? obj.items.map(it => ({
    name: str_(it.name), normalized_name: str_(it.normalized_name) || str_(it.name),
    quantity: numOrNull_(it.quantity), unit: str_(it.unit),
    total_price: numOrNull_(it.total_price), unit_price: numOrNull_(it.unit_price),
    category: RECEIPT_CATEGORIES.indexOf(it.category) >= 0 ? it.category : 'other',
    needed_for: str_(it.needed_for), kosher_sensitive: !!it.kosher_sensitive,
    reimbursable: it.reimbursable !== false, confidence: conf_(it.confidence)
  })) : [];

  return {
    merchant_name: str_(obj.merchant_name) || null,
    date: str_(obj.date) || null,
    subtotal: numOrNull_(obj.subtotal), tax: numOrNull_(obj.tax),
    tip: numOrNull_(obj.tip), total: numOrNull_(obj.total),
    currency: 'USD', payment_method: str_(obj.payment_method) || null,
    items: items,
    warnings: Array.isArray(obj.warnings) ? obj.warnings.map(String) : [],
    overall_confidence: conf_(obj.overall_confidence)
  };
}

function blankExtraction_(warning) {
  return { merchant_name: null, date: null, subtotal: null, tax: null, tip: null, total: null,
    currency: 'USD', payment_method: null, items: [], warnings: [warning], overall_confidence: 'low' };
}

/* ------------------------------------------------------------------ *
 * Tiny utils
 * ------------------------------------------------------------------ */

function num_(v)        { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function numOrNull_(v)  { if (v === null || v === undefined || v === '') return null; const n = parseFloat(v); return isNaN(n) ? null : n; }
function numOrBlank_(v) { const n = numOrNull_(v); return n === null ? '' : n; }
function str_(v)        { return v == null ? '' : String(v).trim(); }
function conf_(v)       { return ['low','medium','high'].indexOf(v) >= 0 ? v : 'low'; }
function sum_(a)        { return a.reduce((x, y) => x + y, 0); }
function round2_(n)     { return Math.round(n * 100) / 100; }
function fmtDate_(d)    { if (!d) return ''; try { return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch (e) { return String(d); } }
