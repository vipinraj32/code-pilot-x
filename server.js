// Cafe Compute Indore — Codex credit distribution
// Rules: only checked-in attendees can claim, 1 email = 1 link, a link is never given twice,
// and the same email always gets back the same link.
//
// Storage: Upstash Redis when its env vars are set (required on Vercel, whose disk is not persistent),
// otherwise a local JSON file in ./output (for running on your own machine).

const fs = require('fs');
const path = require('path');
const express = require('express');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3000;
const DATA_DIR = path.join(__dirname, 'data');
const OUTPUT_DIR = path.join(__dirname, 'output');
const STORE_FILE = path.join(OUTPUT_DIR, 'assignments.json');
const STORE_CSV = path.join(OUTPUT_DIR, 'assignments.csv');
const SHEET_EXT = /\.(xlsx|xls|csv)$/i;

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const URL_RE = /^https?:\/\/\S+$/i;
const CHECKIN_HEADER_RE = /check.?in|checked|attended/i;
const EMAIL_HEADER_RE = /^e-?mail(\s*(address|id))?$/i;
const NOT_CHECKED_IN = new Set(['', 'no', 'n', 'false', '0', 'not checked in', 'absent', 'pending']);

const normalizeEmail = (e) => String(e || '').trim().toLowerCase();

// ---------------------------------------------------------------- spreadsheet loading

let attendees = new Set(); // checked-in emails: allowed to claim
let registered = new Set(); // every email in the attendee sheet, checked in or not
let links = [];
let dataSignature = null;

function sheetFiles() {
  if (!fs.existsSync(DATA_DIR)) return [];
  return fs.readdirSync(DATA_DIR)
    .filter((f) => SHEET_EXT.test(f) && !f.startsWith('~$'))
    .map((f) => path.join(DATA_DIR, f));
}

function currentSignature() {
  return sheetFiles().map((f) => `${f}:${fs.statSync(f).mtimeMs}`).join('|');
}

// Every sheet in every file under data/ is classified automatically:
// mostly emails -> attendee sheet, mostly http(s) links -> credit link sheet.
function loadData() {
  const nextAttendees = new Set();
  const nextRegistered = new Set();
  const nextLinks = [];
  const seenLinks = new Set();
  const summary = [];

  for (const file of sheetFiles()) {
    const wb = XLSX.read(fs.readFileSync(file));
    for (const sheetName of wb.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '', raw: false });
      const cells = rows.flat().map((c) => String(c).trim()).filter(Boolean);
      const emailCount = cells.filter((c) => EMAIL_RE.test(c)).length;
      const urlCount = cells.filter((c) => URL_RE.test(c)).length;
      const label = `${path.basename(file)} › ${sheetName}`;

      if (emailCount === 0 && urlCount === 0) continue;

      const header = (rows[0] || []).map((h) => String(h).trim());
      const emailCol = header.findIndex((h) => EMAIL_HEADER_RE.test(h));

      if (emailCol !== -1 || emailCount >= urlCount) {
        // Attendee sheet (e.g. a Luma export). If it has a check-in column such as "checked_in_at",
        // a row only counts when that cell is filled (a check-in time / "yes"); empty = not checked in.
        const checkCol = header.findIndex((h) => CHECKIN_HEADER_RE.test(h));
        const hasHeader = emailCol !== -1 || checkCol !== -1;
        let count = 0;
        let pending = 0;
        rows.forEach((row, i) => {
          if (hasHeader && i === 0) return;
          // With an email column, read only that column — other columns (e.g. "referred_by")
          // can hold someone else's email, who must not be able to claim through this row.
          const emails = (emailCol !== -1 ? [row[emailCol]] : row).map(normalizeEmail).filter((v) => EMAIL_RE.test(v));
          const checkedIn = checkCol === -1 || !NOT_CHECKED_IN.has(String(row[checkCol]).trim().toLowerCase());
          for (const v of emails) {
            nextRegistered.add(v);
            if (checkedIn && !nextAttendees.has(v)) { nextAttendees.add(v); count++; }
            else if (!checkedIn) pending++;
          }
        });
        const filter = checkCol !== -1 ? ` (checked in via "${header[checkCol]}"; ${pending} registered but not checked in)` : '';
        summary.push(`  attendees  ${String(count).padStart(4)}  ← ${label}${filter}`);
      } else {
        let count = 0;
        for (const c of cells) {
          if (URL_RE.test(c) && !seenLinks.has(c)) { seenLinks.add(c); nextLinks.push(c); count++; }
        }
        summary.push(`  links      ${String(count).padStart(4)}  ← ${label}`);
      }
    }
  }

  attendees = nextAttendees;
  registered = nextRegistered;
  links = nextLinks;
  dataSignature = currentSignature();
  console.log(`[data] loaded ${attendees.size} checked-in attendees, ${links.length} credit links`);
  summary.forEach((s) => console.log(s));
  if (!summary.length) console.log('  (no spreadsheets found — put your .xlsx/.csv files in ./data)');
}

function ensureFreshData() {
  // Lets organisers update the sheets during the event without restarting (local mode).
  if (dataSignature === null || currentSignature() !== dataSignature) {
    try { loadData(); } catch (err) { console.error('[data] reload failed, keeping previous data:', err.message); }
  }
}

// ---------------------------------------------------------------- stores

// Local: JSON file. Each claim runs synchronously, so two requests can never get the same link.
function createFileStore() {
  let store = { assignments: {} };
  if (fs.existsSync(STORE_FILE)) store = JSON.parse(fs.readFileSync(STORE_FILE, 'utf8'));

  function save() {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    const tmp = `${STORE_FILE}.tmp`;
    fs.writeFileSync(tmp, JSON.stringify(store, null, 2));
    fs.renameSync(tmp, STORE_FILE);
    fs.writeFileSync(STORE_CSV, toCsv(Object.entries(store.assignments).map(([email, a]) => ({ email, ...a }))));
  }

  return {
    name: 'file',
    async claim(email, pool) {
      const existing = store.assignments[email];
      if (existing) return { status: 'already_claimed', link: existing.link, claimedAt: existing.claimedAt };
      if (!attendees.has(email)) return { status: 'not_checked_in' };

      const used = new Set(Object.values(store.assignments).map((a) => a.link));
      const link = pool.find((l) => !used.has(l));
      if (!link) return { status: 'exhausted' };

      store.assignments[email] = { link, claimedAt: new Date().toISOString() };
      save();
      return { status: 'claimed', link, claimedAt: store.assignments[email].claimedAt };
    },
    async given() { return Object.keys(store.assignments).length; },
    async all() { return Object.entries(store.assignments).map(([email, a]) => ({ email, ...a })); },
  };
}

// Vercel: Upstash Redis. The Lua script runs atomically inside Redis, so even many
// serverless instances at once can never hand the same link to two emails.
const CLAIM_SCRIPT = `
local existing = redis.call('HGET', KEYS[1], ARGV[1])
if existing then return {'already_claimed', existing, redis.call('HGET', KEYS[3], ARGV[1])} end
if ARGV[3] ~= '1' then return {'not_checked_in'} end
for i = 4, #ARGV do
  if redis.call('HSETNX', KEYS[2], ARGV[i], ARGV[1]) == 1 then
    redis.call('HSET', KEYS[1], ARGV[1], ARGV[i])
    redis.call('HSET', KEYS[3], ARGV[1], ARGV[2])
    return {'claimed', ARGV[i], ARGV[2]}
  end
end
return {'exhausted'}
`;
const K_BY_EMAIL = 'credits:by_email';
const K_BY_LINK = 'credits:by_link';
const K_TIME = 'credits:claimed_at';

function createRedisStore(url, token) {
  const { Redis } = require('@upstash/redis');
  const redis = new Redis({ url, token, automaticDeserialization: false });
  return {
    name: 'redis',
    async claim(email, pool) {
      const now = new Date().toISOString();
      const isAttendee = attendees.has(email) ? '1' : '0';
      const [status, link, claimedAt] = await redis.eval(CLAIM_SCRIPT, [K_BY_EMAIL, K_BY_LINK, K_TIME], [email, now, isAttendee, ...pool]);
      return link ? { status, link, claimedAt } : { status };
    },
    async given() { return redis.hlen(K_BY_EMAIL); },
    async all() {
      const [byEmail, times] = await Promise.all([redis.hgetall(K_BY_EMAIL), redis.hgetall(K_TIME)]);
      return Object.entries(byEmail || {}).map(([email, link]) => ({ email, link, claimedAt: (times || {})[email] }));
    },
  };
}

function toCsv(rows) {
  return ['email,link,claimed_at'].concat(rows.map((r) => `${r.email},${r.link},${r.claimedAt || ''}`)).join('\n');
}

const REDIS_URL = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
const REDIS_TOKEN = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
if (process.env.VERCEL && !REDIS_URL) {
  console.error('[store] Running on Vercel without Redis — claims will NOT persist. Connect Upstash Redis to this project.');
}
const store = REDIS_URL && REDIS_TOKEN ? createRedisStore(REDIS_URL, REDIS_TOKEN) : createFileStore();

// ---------------------------------------------------------------- simple per-IP rate limit

const hits = new Map();
function rateLimit(req, res, next) {
  const now = Date.now();
  const ip = req.headers['x-forwarded-for']?.split(',')[0].trim() || req.ip;
  const recent = (hits.get(ip) || []).filter((t) => now - t < 60_000);
  recent.push(now);
  hits.set(ip, recent);
  if (recent.length > 20) return res.status(429).json({ status: 'rate_limited' });
  next();
}

// ---------------------------------------------------------------- server

const app = express();
app.set('trust proxy', true);
app.use(express.json({ limit: '2kb' }));
app.use(express.static(path.join(__dirname, 'public')));

app.post('/api/claim', rateLimit, async (req, res) => {
  const email = normalizeEmail(req.body && req.body.email);
  if (!EMAIL_RE.test(email) || email.length > 254) return res.status(400).json({ status: 'invalid_email' });
  ensureFreshData();
  try {
    const result = await store.claim(email, links);
    // Tell registered-but-not-yet-checked-in guests to check in first, rather than "not found".
    if (result.status === 'not_checked_in' && registered.has(email)) result.status = 'not_checked_in_yet';
    if (result.status === 'claimed') console.log(`[claim] ${email} → link #${links.indexOf(result.link) + 1}`);
    res.json(result);
  } catch (err) {
    console.error('[claim] failed:', err);
    res.status(500).json({ status: 'error' });
  }
});

app.get('/api/stats', async (req, res) => {
  ensureFreshData();
  try {
    const given = await store.given();
    res.json({ total: links.length, given, remaining: Math.max(links.length - given, 0) });
  } catch (err) {
    res.status(500).json({ status: 'error' });
  }
});

// Download who-got-which-link as CSV: /api/export?key=<ADMIN_KEY>  (disabled unless ADMIN_KEY is set)
app.get('/api/export', async (req, res) => {
  if (!process.env.ADMIN_KEY || req.query.key !== process.env.ADMIN_KEY) return res.status(404).end();
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="codex-credit-claims.csv"');
  res.send(toCsv(await store.all()));
});

ensureFreshData();
console.log(`[store] using ${store.name} storage`);

module.exports = app;

if (require.main === module) {
  app.listen(PORT, () => console.log(`\n☕ Cafe Compute Indore credits running at http://localhost:${PORT}\n`));
}
