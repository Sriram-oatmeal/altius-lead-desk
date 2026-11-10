/**
 * Altius Leads – Google Sheet backend
 * Receives enquiries from the landing page and serves the Lead Desk.
 *
 * SETUP (one time):
 * 1. Paste this whole file into Extensions → Apps Script of your "Altius Leads" Google Sheet.
 * 2. Run the function "setup" once and allow access when Google asks.
 * 3. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone → Deploy.
 * 4. Copy the Web app URL. Put it in the landing page and in the Lead Desk settings.
 */

// Private key for the Lead Desk. Keep it secret: anyone with it can read and edit your leads.
const SECRET = 'altius-xilUC_C2vKtITelR0Aesq6Nt';

// Email alert for every new website lead (true = on).
const EMAIL_ALERTS = true;

const SHEET_NAME = 'Leads';
const COLS = ['id', 'createdAt', 'updatedAt', 'name', 'phone', 'source', 'date', 'budget', 'config',
  'timeline', 'location', 'status', 'brigadeReg', 'nextAction', 'nextDate', 'value', 'notes', 'campaign'];
const TZ = 'Asia/Kolkata';

/* ---------- setup ---------- */
function setup() {
  const sh = sheet_();
  sh.getRange(1, 1, 1, COLS.length).setFontWeight('bold').setBackground('#13213A').setFontColor('#FFFFFF');
  sh.autoResizeColumns(1, COLS.length);
  // Ask for mail permission now, so alerts work later.
  if (EMAIL_ALERTS) MailApp.getRemainingDailyQuota();
  Logger.log('Ready. Now deploy as a Web app.');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------- helpers ---------- */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/^[=+\-@]+(?=\D)/, '').slice(0, max || 300); }

function readAll_(sh) {
  const values = sh.getDataRange().getDisplayValues();
  const head = values.shift() || COLS;
  return values.filter(r => r[0]).map(r => {
    const o = {};
    head.forEach((k, j) => { o[k] = r[j]; });
    o.value = o.value ? Number(o.value) : null;
    return o;
  });
}

function rowIndex_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function writeRow_(sh, row, lead) {
  const vals = COLS.map(k => clean_(lead[k] == null ? '' : lead[k], k === 'notes' ? 2000 : 300));
  // Plain text, so phone numbers and dates are stored exactly as typed.
  sh.getRange(row, 1, 1, COLS.length).setNumberFormat('@').setValues([vals]);
}

function newId_() { return 'l' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

/* ---------- health check (open the URL in a browser to test) ---------- */
function doGet() {
  return json_({ ok: true, service: 'Altius Leads', time: new Date().toISOString() });
}

/* ---------- write ---------- */
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_();

    // 1) Enquiry from the public landing page (no key needed, add-only)
    if (body.action === 'web') {
      if (body.website) return json_({ ok: true }); // spam trap filled in by bots
      const name = clean_(body.name, 80).trim();
      const phone = clean_(body.phone, 30).trim();
      if (!name || phone.replace(/\D/g, '').length < 10) return json_({ ok: false, error: 'invalid' });

      const src = String(body.utm_source || '').toLowerCase();
      const source = /fb|facebook|insta|meta|ig/.test(src) ? 'Meta ads'
        : /google|gads|adwords/.test(src) ? 'Google ads'
        : /whatsapp|wa/.test(src) ? 'WhatsApp'
        : 'Landing page';
      const lead = {
        id: newId_(), createdAt: new Date().toISOString(), updatedAt: Date.now(),
        name, phone, source, date: today_(),
        budget: body.budget, config: body.config, timeline: body.when,
        location: /abroad|nri/i.test(body.where || '') ? 'NRI' : 'India',
        status: 'New', brigadeReg: '',
        nextAction: 'Call within 5 minutes and qualify', nextDate: today_(), value: '',
        notes: 'Website form · lives in: ' + clean_(body.where, 40) + ' · site visit: ' + clean_(body.visit, 40),
        campaign: [body.utm_source, body.utm_medium, body.utm_campaign].filter(Boolean).map(x => clean_(x, 60)).join(' / ')
      };
      writeRow_(sh, sh.getLastRow() + 1, lead);
      if (EMAIL_ALERTS) {
        try {
          MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
            'New Altius lead: ' + name + ' (' + lead.location + ', ' + (lead.budget || '') + ')',
            'Call now: ' + phone + '\n\n' +
            'Home type: ' + lead.config + '\nBudget: ' + lead.budget + '\nPlanning to buy: ' + lead.timeline +
            '\nLives in: ' + body.where + '\nSite visit: ' + body.visit + '\nSource: ' + source +
            (lead.campaign ? ' (' + lead.campaign + ')' : '') +
            '\n\nWhatsApp: https://wa.me/' + (phone.replace(/\D/g, '').length === 10 ? '91' : '') + phone.replace(/\D/g, ''));
        } catch (err) { /* email quota or permission – lead is still saved */ }
      }
      return json_({ ok: true, id: lead.id });
    }

    // 2) Lead Desk actions (key required)
    if (body.key !== SECRET) return json_({ ok: false, error: 'bad_key' });

    if (body.action === 'list') {
      return json_({ ok: true, leads: readAll_(sh), at: Date.now() });
    }
    if (body.action === 'add') {
      const lead = Object.assign({}, body.lead, { id: newId_(), createdAt: new Date().toISOString(), updatedAt: Date.now() });
      writeRow_(sh, sh.getLastRow() + 1, lead);
      return json_({ ok: true, id: lead.id });
    }
    if (body.action === 'update') {
      const row = rowIndex_(sh, body.id);
      if (row < 0) return json_({ ok: false, error: 'not_found' });
      const cur = {};
      const vals = sh.getRange(row, 1, 1, COLS.length).getDisplayValues()[0];
      COLS.forEach((k, j) => { cur[k] = vals[j]; });
      const lead = Object.assign(cur, body.lead, { id: body.id, updatedAt: Date.now() });
      writeRow_(sh, row, lead);
      return json_({ ok: true });
    }
    if (body.action === 'delete') {
      const row = rowIndex_(sh, body.id);
      if (row > 0) sh.deleteRow(row);
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
