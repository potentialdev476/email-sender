const express    = require('express');
const multer     = require('multer');
const nodemailer = require('nodemailer');
const ExcelJS    = require('exceljs');
const path       = require('path');
const fs         = require('fs');

const app  = express();
const PORT = 3000;

// ── Global crash guards ───────────────────────────────────────────────────────
process.on('uncaughtException', err => {
  // Let EADDRINUSE bubble — nodemon needs the process to exit so it can retry
  if (err.code === 'EADDRINUSE') { process.exit(1); }
  console.error('[uncaughtException]', err.message);
});
process.on('unhandledRejection', err => {
  console.error('[unhandledRejection]', err?.message || err);
});

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── Directories ───────────────────────────────────────────────────────────────
const UPLOAD_DIR = path.join(__dirname, 'uploads');
const DATA_DIR   = path.join(__dirname, 'data');
const STATE_FILE = path.join(DATA_DIR, 'state.json');
const LOGS_FILE  = path.join(DATA_DIR, 'logs.json');

[UPLOAD_DIR, DATA_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d); });

const upload = multer({
  dest: UPLOAD_DIR,
  fileFilter: (req, file, cb) => {
    if (file.originalname.endsWith('.xlsx')) cb(null, true);
    else cb(new Error('Only .xlsx files are allowed'));
  },
  limits: { fileSize: 16 * 1024 * 1024 },
});

// ── Helpers ───────────────────────────────────────────────────────────────────
function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

// ── State ─────────────────────────────────────────────────────────────────────
const state = {
  smtp: { host: '', port: 587, username: '', password: '', fromName: '', useTLS: true, signature: '' },
  categories:            [],
  templates:             [],
  emailList:             [],
  selectedCategoryId:    null,
  templateRotationIndex: 0,
  intervalSeconds:       30,
  campaignVars:          {},   // custom variables e.g. { platform: 'LinkedIn' }
  isRunning:             false,
  currentIndex:          0,
  logs:                  [],
  timer:                 null,
  nextSendAt:            null,
};

// ── Persistence ───────────────────────────────────────────────────────────────
let _saveStateTimer = null;
function saveState() {
  clearTimeout(_saveStateTimer);
  _saveStateTimer = setTimeout(() => {
    try {
      const snapshot = {
        smtp:                  state.smtp,
        categories:            state.categories,
        templates:             state.templates,
        emailList:             state.emailList,
        selectedCategoryId:    state.selectedCategoryId,
        templateRotationIndex: state.templateRotationIndex,
        intervalSeconds:       state.intervalSeconds,
        campaignVars:          state.campaignVars,
        currentIndex:          state.currentIndex,
      };
      fs.writeFileSync(STATE_FILE, JSON.stringify(snapshot, null, 2), 'utf8');
    } catch (e) {
      console.error('[persist] Failed to save state:', e.message);
    }
  }, 400);
}

let _saveLogsTimer = null;
function saveLogs() {
  clearTimeout(_saveLogsTimer);
  _saveLogsTimer = setTimeout(() => {
    try {
      fs.writeFileSync(LOGS_FILE, JSON.stringify(state.logs, null, 2), 'utf8');
    } catch (e) {
      console.error('[persist] Failed to save logs:', e.message);
    }
  }, 800);
}

function loadPersistedData() {
  // Load state
  if (fs.existsSync(STATE_FILE)) {
    try {
      const saved = JSON.parse(fs.readFileSync(STATE_FILE, 'utf8'));
      Object.assign(state.smtp,       saved.smtp       || {});
      state.categories            = saved.categories            || [];
      state.templates             = saved.templates             || [];
      state.emailList             = saved.emailList             || [];
      state.selectedCategoryId    = saved.selectedCategoryId    || null;
      state.templateRotationIndex = saved.templateRotationIndex || 0;
      state.intervalSeconds       = saved.intervalSeconds ?? saved.intervalMinutes * 60 ?? 30;
      state.campaignVars          = saved.campaignVars          || {};
      state.currentIndex          = saved.currentIndex          || 0;
      console.log(`[persist] State loaded — ${state.categories.length} categories, ` +
                  `${state.templates.length} templates, ${state.emailList.length} recipients`);
      if (state.currentIndex > 0 && state.currentIndex < state.emailList.length) {
        console.log(`[persist] ⚠️  Unfinished job detected: ${state.currentIndex}/${state.emailList.length} sent`);
      }
    } catch (e) {
      console.error('[persist] Failed to load state.json:', e.message);
    }
  }
  // Load logs
  if (fs.existsSync(LOGS_FILE)) {
    try {
      state.logs = JSON.parse(fs.readFileSync(LOGS_FILE, 'utf8'));
      console.log(`[persist] ${state.logs.length} log entries loaded`);
    } catch (e) {
      console.error('[persist] Failed to load logs.json:', e.message);
    }
  }
}

// ── Logging ───────────────────────────────────────────────────────────────────
function addLog(level, message, recipient = null) {
  const entry = {
    id:        genId(),
    timestamp: new Date().toLocaleString('en-US', { hour12: false }),
    level,
    message,
    recipient,
  };
  state.logs.unshift(entry);
  if (state.logs.length > 2000) state.logs.length = 2000;
  saveLogs();
  return entry;
}

// ── Template helpers ──────────────────────────────────────────────────────────
function getCategoryTemplates(catId) {
  return state.templates.filter(t => t.categoryId === (catId ?? state.selectedCategoryId));
}

function getNextTemplate() {
  const templates = getCategoryTemplates();
  if (!templates.length) return null;
  const template = templates[state.templateRotationIndex % templates.length];
  state.templateRotationIndex = (state.templateRotationIndex + 1) % templates.length;
  return template;
}

function replacePlaceholders(text, name, email) {
  let result = text
    .replace(/\{\{name\}\}/g, name)
    .replace(/\{\{email\}\}/g, email);
  // Apply campaign variables (e.g. {{platform}}, {{jobTitle}}, etc.)
  for (const [key, val] of Object.entries(state.campaignVars || {})) {
    result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), val);
  }
  return result;
}

// ── Email sender ──────────────────────────────────────────────────────────────
async function sendEmail(recipient, template) {
  const { name, email } = recipient;
  const cfg = state.smtp;

  const transporter = nodemailer.createTransport({
    host: cfg.host, port: cfg.port,
    secure: !cfg.useTLS, requireTLS: cfg.useTLS,
    auth: { user: cfg.username, pass: cfg.password },
    connectionTimeout: 15000,
  });

  const subject   = replacePlaceholders(template.subject, name, email);
  const body      = replacePlaceholders(template.body,    name, email);
  const sig       = cfg.signature || '';
  const bodyIsHtml = /<[a-z][\s\S]*>/i.test(body);
  const sigIsHtml  = /<[a-z][\s\S]*>/i.test(sig);

  let htmlBody = null;
  let textBody = null;

  if (bodyIsHtml || sigIsHtml) {
    // Convert plain-text body to HTML if needed, then append signature
    const bodyHtml = bodyIsHtml ? body : body.replace(/\n/g, '<br>');
    htmlBody = sig ? `${bodyHtml}<br><br>${sig}` : bodyHtml;
  } else {
    // Both body and signature are plain text
    textBody = sig ? `${body}\n\n${sig.replace(/<[^>]+>/g, '')}` : body;
  }

  await transporter.sendMail({
    from: `"${cfg.fromName}" <${cfg.username}>`,
    to: email, subject,
    ...(htmlBody ? { html: htmlBody } : { text: textBody }),
  });
}

// ── Scheduler ─────────────────────────────────────────────────────────────────
function scheduleNext() {
  if (!state.isRunning) return;

  const idx = state.currentIndex;
  if (idx >= state.emailList.length) {
    addLog('info', '✅ All emails sent. Scheduler finished.');
    state.isRunning  = false;
    state.nextSendAt = null;
    saveState();
    return;
  }

  const recipient = state.emailList[idx];
  const delayMs   = idx === 0 ? 0 : state.intervalSeconds * 1000;
  state.nextSendAt = new Date(Date.now() + delayMs);

  state.timer = setTimeout(async () => {
    if (!state.isRunning) return;

    const template = getNextTemplate();
    if (!template) {
      addLog('error', '❌ No templates in selected category. Stopping.');
      state.isRunning = false;
      return;
    }

    try {
      await sendEmail(recipient, template);
      addLog('success',
        `[${template.name}] → ${recipient.name} <${recipient.email}>`,
        recipient.email);
    } catch (err) {
      addLog('error',
        `[${template.name}] ✗ ${recipient.name} <${recipient.email}>: ${err.message}`,
        recipient.email);
    }

    state.currentIndex++;
    saveState();       // persist progress after every send
    scheduleNext();
  }, delayMs);
}

// ═══════════════════════════════════════════════════════════════════════════════
// API Routes
// ═══════════════════════════════════════════════════════════════════════════════

// ── SMTP ──────────────────────────────────────────────────────────────────────
app.get('/api/smtp', (req, res) => {
  const cfg = { ...state.smtp, password: state.smtp.password ? '●●●●●●●●' : '' };
  res.json(cfg);
});

app.post('/api/smtp', (req, res) => {
  const { host, port, username, password, fromName, useTLS, signature } = req.body;
  state.smtp = {
    host:      host || '',
    port:      parseInt(port) || 587,
    username:  username || '',
    password:  password && password !== '●●●●●●●●' ? password : state.smtp.password,
    fromName:  fromName || '',
    useTLS:    useTLS !== false,
    signature: signature ?? state.smtp.signature ?? '',
  };
  saveState();
  res.json({ ok: true, message: 'SMTP settings saved.' });
});

app.post('/api/smtp/test', async (req, res) => {
  let transporter;
  try {
    const cfg = state.smtp;
    transporter = nodemailer.createTransport({
      host: cfg.host, port: cfg.port,
      secure: !cfg.useTLS, requireTLS: cfg.useTLS,
      auth: { user: cfg.username, pass: cfg.password },
      connectionTimeout: 10000,
      socketTimeout: 10000,
    });
    // Suppress unhandled 'error' events on the transporter
    transporter.on('error', () => {});
    await transporter.verify();
    res.json({ ok: true, message: '✅ SMTP connection successful!' });
  } catch (err) {
    res.status(400).json({ ok: false, message: `❌ ${err.message}` });
  } finally {
    try { transporter?.close(); } catch (_) {}
  }
});

// ── Categories ────────────────────────────────────────────────────────────────
app.get('/api/categories', (req, res) => {
  res.json(state.categories.map(c => ({
    ...c, templateCount: getCategoryTemplates(c.id).length,
  })));
});

app.post('/api/categories', (req, res) => {
  const name = req.body.name?.trim();
  if (!name) return res.status(400).json({ ok: false, message: 'Name is required.' });
  const category = { id: genId(), name };
  state.categories.push(category);
  saveState();
  res.json({ ok: true, category });
});

app.put('/api/categories/:id', (req, res) => {
  const cat = state.categories.find(c => c.id === req.params.id);
  if (!cat) return res.status(404).json({ ok: false, message: 'Category not found.' });
  cat.name = req.body.name?.trim() || cat.name;
  saveState();
  res.json({ ok: true, category: cat });
});

app.delete('/api/categories/:id', (req, res) => {
  const idx = state.categories.findIndex(c => c.id === req.params.id);
  if (idx === -1) return res.status(404).json({ ok: false, message: 'Category not found.' });
  state.categories.splice(idx, 1);
  state.templates = state.templates.filter(t => t.categoryId !== req.params.id);
  if (state.selectedCategoryId === req.params.id) state.selectedCategoryId = null;
  saveState();
  res.json({ ok: true });
});

// ── Templates ─────────────────────────────────────────────────────────────────
app.get('/api/templates', (req, res) => {
  const { categoryId } = req.query;
  res.json(categoryId
    ? state.templates.filter(t => t.categoryId === categoryId)
    : state.templates);
});

app.post('/api/templates', (req, res) => {
  const { categoryId, name, subject, body } = req.body;
  if (!categoryId || !name?.trim() || !subject?.trim() || !body?.trim())
    return res.status(400).json({ ok: false, message: 'All fields are required.' });
  if (!state.categories.find(c => c.id === categoryId))
    return res.status(400).json({ ok: false, message: 'Category not found.' });
  const template = { id: genId(), categoryId, name: name.trim(), subject: subject.trim(), body: body.trim() };
  state.templates.push(template);
  saveState();
  res.json({ ok: true, template });
});

app.put('/api/templates/:id', (req, res) => {
  const tmpl = state.templates.find(t => t.id === req.params.id);
  if (!tmpl) return res.status(404).json({ ok: false, message: 'Template not found.' });
  if (req.body.name)    tmpl.name    = req.body.name.trim();
  if (req.body.subject) tmpl.subject = req.body.subject.trim();
  if (req.body.body)    tmpl.body    = req.body.body.trim();
  saveState();
  res.json({ ok: true, template: tmpl });
});

app.delete('/api/templates/:id', (req, res) => {
  const idx = state.templates.findIndex(t => t.id === req.params.id);
  if (idx === -1) return res.status(404).json({ ok: false, message: 'Template not found.' });
  state.templates.splice(idx, 1);
  saveState();
  res.json({ ok: true });
});

// ── Select Category ───────────────────────────────────────────────────────────
app.post('/api/select-category', (req, res) => {
  const { categoryId } = req.body;
  if (categoryId && !state.categories.find(c => c.id === categoryId))
    return res.status(400).json({ ok: false, message: 'Category not found.' });
  state.selectedCategoryId    = categoryId || null;
  state.templateRotationIndex = 0;
  saveState();
  const cat       = state.categories.find(c => c.id === categoryId);
  const tmplCount = getCategoryTemplates(categoryId).length;
  res.json({
    ok: true,
    message: cat
      ? `✅ Category "${cat.name}" selected — ${tmplCount} template(s)`
      : 'Category cleared.',
  });
});

// ── Campaign Variables ────────────────────────────────────────────────────────
app.get('/api/campaign-vars', (req, res) => {
  res.json(state.campaignVars || {});
});

app.post('/api/campaign-vars', (req, res) => {
  const incoming = req.body;
  if (typeof incoming !== 'object' || Array.isArray(incoming))
    return res.status(400).json({ ok: false, message: 'Invalid format.' });
  const cleaned = {};
  for (const [k, v] of Object.entries(incoming)) {
    const key = k.trim().replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
    if (key && key !== 'name' && key !== 'email')
      cleaned[key] = String(v);
  }
  state.campaignVars = cleaned;
  saveState();
  res.json({ ok: true, vars: state.campaignVars });
});

// ── Settings ──────────────────────────────────────────────────────────────────
app.post('/api/settings', (req, res) => {
  if (req.body.intervalSeconds !== undefined)
    state.intervalSeconds = Math.max(1, parseInt(req.body.intervalSeconds));
  saveState();
  res.json({ ok: true, intervalSeconds: state.intervalSeconds });
});

// ── Upload XLSX ───────────────────────────────────────────────────────────────

// Safely extract a plain string from any ExcelJS cell value type
function cellText(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v.trim();
  if (typeof v === 'number') return String(v).trim();
  if (typeof v === 'boolean') return '';
  if (typeof v === 'object') {
    if (v.text)     return String(v.text).trim();          // hyperlink cell
    if (v.richText) return v.richText.map(r => r.text || '').join('').trim(); // rich text
    if (v.result !== undefined) return String(v.result).trim(); // formula result
    if (v.error)    return '';                             // formula error
  }
  return String(v).trim();
}

app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, message: 'No file uploaded.' });
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(req.file.path);
    const ws = wb.worksheets[0];

    const emailList = [];
    let skipped = 0, rowNum = 0;

    ws.eachRow(row => {
      rowNum++;
      if (rowNum === 1) return; // skip header row

      const name  = cellText(row.getCell(1));
      const email = cellText(row.getCell(2));

      if (name && email && email.includes('@')) {
        emailList.push({ name, email });
      } else {
        // Only count as skipped if the row isn't completely empty
        if (name || email) skipped++;
      }
    });

    state.emailList             = emailList;
    state.currentIndex          = 0;
    state.templateRotationIndex = 0;
    fs.unlink(req.file.path, () => {});
    saveState();

    addLog('info', `📋 Loaded ${emailList.length} recipients from "${req.file.originalname}" (${skipped} invalid rows skipped)`);
    res.json({ ok: true, count: emailList.length, skipped, preview: emailList.slice(0, 5) });
  } catch (err) {
    res.status(500).json({ ok: false, message: err.message });
  }
});

// ── Scheduler ─────────────────────────────────────────────────────────────────
app.post('/api/scheduler/start', (req, res) => {
  if (state.isRunning)
    return res.status(400).json({ ok: false, message: 'Already running.' });
  if (!state.emailList.length)
    return res.status(400).json({ ok: false, message: 'No email list loaded.' });
  if (!state.selectedCategoryId)
    return res.status(400).json({ ok: false, message: 'No category selected. Go to Recipients tab.' });
  if (!getCategoryTemplates().length)
    return res.status(400).json({ ok: false, message: 'Selected category has no templates.' });
  if (!state.smtp.host || !state.smtp.username)
    return res.status(400).json({ ok: false, message: 'SMTP is not configured.' });

  if (state.currentIndex >= state.emailList.length) state.currentIndex = 0;

  state.isRunning = true;
  const cat     = state.categories.find(c => c.id === state.selectedCategoryId);
  const pending = state.emailList.length - state.currentIndex;
  addLog('info',
    `🚀 Started — ${pending} recipients remaining | Category: "${cat?.name}" | ` +
    `${getCategoryTemplates().length} templates | Interval: ${state.intervalSeconds}s`);
  scheduleNext();
  res.json({ ok: true, message: 'Scheduler started.' });
});

app.post('/api/scheduler/stop', (req, res) => {
  if (!state.isRunning)
    return res.status(400).json({ ok: false, message: 'Not running.' });
  state.isRunning = false;
  if (state.timer) { clearTimeout(state.timer); state.timer = null; }
  state.nextSendAt = null;
  addLog('info', `🛑 Scheduler stopped — progress saved (${state.currentIndex}/${state.emailList.length})`);
  saveState();
  res.json({ ok: true, message: 'Scheduler stopped.' });
});

app.post('/api/recipients/clear', (req, res) => {
  if (state.isRunning)
    return res.status(400).json({ ok: false, message: 'Stop the scheduler first.' });
  state.emailList             = [];
  state.currentIndex          = 0;
  state.templateRotationIndex = 0;
  saveState();
  addLog('info', '🗑 Recipient list cleared.');
  res.json({ ok: true, message: 'Recipients cleared.' });
});

// ── Status ────────────────────────────────────────────────────────────────────
app.get('/api/status', (req, res) => {
  const total   = state.emailList.length;
  const sent    = state.currentIndex;
  const cat     = state.categories.find(c => c.id === state.selectedCategoryId);
  const resumed = !state.isRunning && sent > 0 && sent < total;
  res.json({
    isRunning:            state.isRunning,
    total, sent,
    remaining:            Math.max(0, total - sent),
    intervalSeconds:      state.intervalSeconds,
    nextSendAt:           state.nextSendAt,
    selectedCategoryId:   state.selectedCategoryId,
    selectedCategoryName: cat?.name || null,
    templateCount:        getCategoryTemplates().length,
    resumable:            resumed,   // ← unfinished job detected
  });
});

// ── Logs ──────────────────────────────────────────────────────────────────────
app.get('/api/logs', (req, res) => {
  const sinceId = req.query.since_id;
  let logs = state.logs;
  if (sinceId) {
    const idx = logs.findIndex(l => l.id === sinceId);
    if (idx > 0)   logs = logs.slice(0, idx);
    if (idx === 0) logs = [];
  }
  res.json({ logs: logs.slice(0, 500) });
});

app.delete('/api/logs', (req, res) => {
  state.logs = [];
  saveLogs();
  res.json({ ok: true });
});

// ── Email list ────────────────────────────────────────────────────────────────
app.get('/api/email-list', (req, res) => {
  res.json({
    count: state.emailList.length,
    currentIndex: state.currentIndex,
    list: state.emailList,
  });
});

// ── Boot ──────────────────────────────────────────────────────────────────────
loadPersistedData();

const server = app.listen(PORT, () => {
  console.log(`\n🚀 Email Sender Bot running at http://localhost:${PORT}`);
  console.log(`💾 Data saved to: ${DATA_DIR}\n`);
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n❌ Port ${PORT} is already in use. Stop the other process and try again.\n`);
    process.exit(1);
  } else {
    console.error('[server error]', err.message);
  }
});
