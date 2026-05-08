const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const crypto  = require('crypto');
const path    = require('path');
const fs      = require('fs');
const db      = require('./db');

const app  = express();
const PORT = process.env.PORT || 3003;

app.use(express.json({ limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const upload = multer({
  dest: process.env.UPLOADS_PATH || path.join(__dirname, 'uploads'),
  limits: { fileSize: 100 * 1024 * 1024 }
});

// ─── helpers ─────────────────────────────────────────────────────────────────
function parseDate(v) {
  if (!v) return null;
  const m = String(v).trim().match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : String(v).trim();
}

// Parse a delimited line, respecting quoted fields
function parseLine(line, delim) {
  const vals = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQ = !inQ; }
    else if (ch === delim && !inQ) { vals.push(cur); cur = ''; }
    else { cur += ch; }
  }
  vals.push(cur);
  return vals.map(v => v.trim());
}

// Parse TXT/CSV as plain text — preserves full tracking numbers (no Excel float conversion)
function parsePlainText(filePath) {
  const content = fs.readFileSync(filePath, 'utf8');
  const lines   = content.split(/\r?\n/).map(l => l.trimEnd()).filter(l => l.length);
  if (lines.length < 2) return [];
  const delim   = lines[0].includes('\t') ? '\t' : ',';
  const headers = parseLine(lines[0], delim).map(h => h.replace(/^﻿/, '')); // strip BOM
  const rows = [];
  for (let i = 1; i < lines.length; i++) {
    const vals = parseLine(lines[i], delim);
    if (!vals.some(v => v)) continue;
    const row = {};
    headers.forEach((h, idx) => { row[h] = vals[idx] !== undefined ? vals[idx] : ''; });
    rows.push(row);
  }
  return rows;
}
function hash(...p) { return crypto.createHash('md5').update(p.join('|')).digest('hex'); }
function fileHash(fp) { return crypto.createHash('md5').update(fs.readFileSync(fp)).digest('hex'); }

// ─── dedup helper — runs after every upload ───────────────────────────────────
function runDedup() {
  // Pass 1: remove rows where row_hash is identical (exact duplicates)
  const p1 = db.prepare(`
    DELETE FROM removal_shipments
    WHERE rowid NOT IN (
      SELECT MIN(rowid) FROM removal_shipments GROUP BY row_hash
    )
  `).run();

  // Pass 2: remove rows where order_id + sku + tracking_number + shipment_date match
  // (catches cases where hash differs due to minor field variation but record is the same)
  const p2 = db.prepare(`
    DELETE FROM removal_shipments
    WHERE rowid NOT IN (
      SELECT MIN(rowid) FROM removal_shipments
      GROUP BY order_id, sku, tracking_number, shipment_date
    )
  `).run();

  const removed = p1.changes + p2.changes;
  if (removed > 0) console.log(`[dedup] removed ${removed} duplicate rows`);
  return removed;
}

// Detect carrier from carrier code + tracking number
function normalizeCarrier(carrier, tracking) {
  const c = (carrier || '').toUpperCase();
  const t = (tracking || '').trim();
  if (t.startsWith('TBA'))  return 'Amazon';
  if (t.startsWith('1Z'))   return 'UPS';
  if (t.match(/^9\d{15,}/)) return 'USPS';
  if (c.includes('UPS'))    return 'UPS';
  if (c.includes('USPS'))   return 'USPS';
  if (c.includes('AMZL') || c.includes('AMAZON')) return 'Amazon';
  if (c.includes('FEDEX'))  return 'FedEx';
  // Amazon carrier codes
  if (c.includes('EXLA') || c.includes('ABNT') || c.includes('AMZL')) return 'Amazon';
  return carrier || 'Unknown';
}

function trackingUrl(carrier, tracking) {
  const t = (tracking || '').trim();
  const c = normalizeCarrier(carrier, tracking);
  if (c === 'UPS')    return `https://www.ups.com/track?tracknum=${t}`;
  if (c === 'USPS')   return `https://tools.usps.com/go/TrackConfirmAction?tLabels=${t}`;
  if (c === 'Amazon') return `https://track.amazon.com/tracking/${t}`;
  if (c === 'FedEx')  return `https://www.fedex.com/fedextrack/?trknbr=${t}`;
  return `https://www.ups.com/track?tracknum=${t}`;
}

// ─── prepared statements ─────────────────────────────────────────────────────
const insertShipment = db.prepare(`
  INSERT OR IGNORE INTO removal_shipments
  (request_date, order_id, shipment_date, sku, fnsku, disposition,
   shipped_quantity, carrier, tracking_number, removal_order_type, row_hash)
  VALUES
  (@request_date, @order_id, @shipment_date, @sku, @fnsku, @disposition,
   @shipped_quantity, @carrier, @tracking_number, @removal_order_type, @row_hash)
`);

function buildRecord(row) {
  const reqDate  = parseDate(row['request-date']  || row['Request Date']  || '');
  const shipDate = parseDate(row['shipment-date'] || row['Shipment Date'] || '');
  const orderId  = String(row['order-id']  || row['Order ID']  || '').trim();
  const sku      = String(row['sku']       || row['SKU']       || '').trim();
  const tracking = String(row['tracking-number'] || row['Tracking Number'] || '').trim();
  const carrier  = String(row['carrier']   || row['Carrier']   || '').trim();
  return {
    request_date:        reqDate  || '',
    order_id:            orderId,
    shipment_date:       shipDate || '',
    sku,
    fnsku:               String(row['fnsku'] || row['FNSKU'] || '').trim(),
    disposition:         String(row['disposition']         || row['Disposition']         || '').trim(),
    shipped_quantity:    parseInt(row['shipped-quantity']  || row['Shipped Quantity']    || 1) || 1,
    carrier,
    tracking_number:     tracking,
    removal_order_type:  String(row['removal-order-type'] || row['Removal Order Type']  || '').trim(),
    row_hash:            hash(orderId, sku, tracking, shipDate || ''),
  };
}

// ─── carrier tracking helpers ─────────────────────────────────────────────────
// Node 18+ has built-in fetch; use https module as fallback
const https = require('https');

function httpsRequest(url, options = {}) {
  return new Promise((resolve, reject) => {
    const u = new URL(url);
    const reqOpts = {
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: options.method || 'GET',
      headers: options.headers || {},
      timeout: 12000,
    };
    const req = https.request(reqOpts, res => {
      let data = '';
      res.on('data', chunk => { data += chunk; });
      res.on('end', () => resolve({
        ok: res.statusCode >= 200 && res.statusCode < 300,
        status: res.statusCode,
        text: () => Promise.resolve(data),
        json: () => Promise.resolve(JSON.parse(data))
      }));
    });
    req.on('error', reject);
    req.on('timeout', () => { req.destroy(); reject(new Error('Request timed out')); });
    if (options.body) req.write(options.body);
    req.end();
  });
}

function fetchWithTimeout(url, options = {}, ms = 12000) {
  if (typeof fetch === 'function') {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, { ...options, signal: ctrl.signal }).finally(() => clearTimeout(timer));
  }
  return httpsRequest(url, options);
}

// ─── UPS OAuth2 Token Cache ───────────────────────────────────────────────────
// Set UPS_CLIENT_ID and UPS_CLIENT_SECRET in Railway env vars.
// Register free at https://developer.ups.com → My Apps → Add App → Tracking API
let upsTokenCache = { token: null, expires: 0 };

async function getUPSToken() {
  if (upsTokenCache.token && Date.now() < upsTokenCache.expires) {
    return upsTokenCache.token;
  }
  const clientId     = process.env.UPS_CLIENT_ID     || '';
  const clientSecret = process.env.UPS_CLIENT_SECRET || '';
  if (!clientId || !clientSecret) return null;

  const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
  try {
    const res = await fetchWithTimeout('https://onlinetools.ups.com/security/v1/oauth/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Authorization': `Basic ${credentials}`,
      },
      body: 'grant_type=client_credentials',
    }, 15000);
    if (!res.ok) { console.error('[UPS] token fetch failed:', res.status); return null; }
    const data = await res.json();
    upsTokenCache.token   = data.access_token;
    upsTokenCache.expires = Date.now() + (data.expires_in - 60) * 1000;
    console.log('[UPS] token refreshed, expires in', data.expires_in, 's');
    return upsTokenCache.token;
  } catch (e) {
    console.error('[UPS] token error:', e.message);
    return null;
  }
}

// UPS Developer API v1 — requires UPS_CLIENT_ID + UPS_CLIENT_SECRET env vars
// Get free API access at https://developer.ups.com
async function checkUPS(trackingNum) {
  try {
    const token = await getUPSToken();
    if (!token) return null;

    const url = `https://onlinetools.ups.com/api/track/v1/details/${encodeURIComponent(trackingNum)}`;
    const res = await fetchWithTimeout(url, {
      method: 'GET',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
        'transId': `removal-${Date.now()}`,
        'transactionSrc': 'removaltracker',
      },
    }, 12000);

    if (!res.ok) {
      const errText = await res.text().catch(() => '');
      console.error(`[UPS] tracking error for ${trackingNum}: HTTP ${res.status}`, errText.slice(0, 200));
      return null;
    }

    const data = await res.json();
    const pkg  = data?.trackResponse?.shipment?.[0]?.package?.[0];
    if (!pkg) return null;

    const act  = pkg.activity?.[0];
    if (!act)  return null;

    const type = act.status?.type;
    const status = type === 'D' ? 'Delivered'
                 : type === 'O' ? 'Out for Delivery'
                 : type === 'X' ? 'Exception'
                 : 'In Transit';
    const loc  = [act.location?.address?.city, act.location?.address?.stateProvince]
                   .filter(Boolean).join(', ');
    return { status, description: act.status?.description || '', location: loc };
  } catch (e) {
    console.error(`[UPS] check error for ${trackingNum}:`, e.message);
    return null;
  }
}

// USPS Web Tools API — requires free USPS_USER_ID from https://www.usps.com/business/web-tools-apis/
// Set USPS_USER_ID in Railway env vars.
async function checkUSPS(trackingNum) {
  try {
    const userId = process.env.USPS_USER_ID || '';
    if (!userId) return null;
    const xml = `<TrackFieldRequest USERID="${userId}"><TrackID ID="${trackingNum}"/></TrackFieldRequest>`;
    const url = `https://secure.shippingapis.com/ShippingAPI.dll?API=TrackV2&XML=${encodeURIComponent(xml)}`;
    const res = await fetchWithTimeout(url, {}, 12000);
    if (!res.ok) return null;
    const text = await res.text();
    // Check for USPS error response
    if (text.includes('<Error>') || text.includes('<Description>')) {
      const desc = text.match(/<Description>(.*?)<\/Description>/)?.[1] || '';
      if (desc) console.error(`[USPS] error for ${trackingNum}:`, desc);
      return null;
    }
    const event = text.match(/<Event>(.*?)<\/Event>/)?.[1]      || '';
    const city  = text.match(/<EventCity>(.*?)<\/EventCity>/)?.[1] || '';
    const state = text.match(/<EventState>(.*?)<\/EventState>/)?.[1] || '';
    const lower = event.toLowerCase();
    const status = lower.includes('delivered') ? 'Delivered'
                 : lower.includes('out for delivery') ? 'Out for Delivery'
                 : lower.includes('exception') || lower.includes('alert') ? 'Exception'
                 : 'In Transit';
    return { status, description: event, location: [city, state].filter(Boolean).join(', ') };
  } catch (e) {
    console.error(`[USPS] check error for ${trackingNum}:`, e.message);
    return null;
  }
}

// Amazon TBA tracking — Amazon Logistics does not expose a public server-side API.
// The track.amazon.com page requires JavaScript/session to render.
// This function marks Amazon packages as "In Transit" (requires manual check or
// Amazon SP-API credentials for automated status).
async function checkAmazon(trackingNum) {
  // Amazon TBA packages cannot be checked via public API without browser session.
  // Return null so they stay at current status rather than being overwritten with bad data.
  return null;
}

function updateOrderStatuses() {
  const orders = db.prepare(`SELECT DISTINCT order_id FROM removal_shipments WHERE order_id != ''`).all();
  const upsertOrder = db.prepare(`INSERT OR IGNORE INTO removal_orders (order_id) VALUES (?)`);
  const updateOrder = db.prepare(`UPDATE removal_orders SET status = ?, updated_at = datetime('now') WHERE order_id = ?`);

  for (const { order_id } of orders) {
    const trackings = db.prepare(`
      SELECT DISTINCT ts.status
      FROM removal_shipments s
      LEFT JOIN tracking_status ts ON s.tracking_number = ts.tracking_number
      WHERE s.order_id = ? AND s.tracking_number != '' AND ts.status IS NOT NULL
    `).all(order_id);

    if (!trackings.length) continue;
    const statuses = trackings.map(t => t.status);
    const orderStatus = statuses.every(s => s === 'Delivered') ? 'Delivered'
      : statuses.some(s => s === 'Out for Delivery') ? 'Out for Delivery'
      : statuses.some(s => s === 'Exception') ? 'Exception'
      : statuses.some(s => s === 'In Transit') ? 'In Transit'
      : 'In Transit';

    upsertOrder.run(order_id);
    updateOrder.run(orderStatus, order_id);
  }
}

async function refreshTracking(batchSize = 50) {
  // Get undelivered tracking numbers not checked in last 4 hours
  const rows = db.prepare(`
    SELECT DISTINCT s.tracking_number, s.carrier
    FROM removal_shipments s
    LEFT JOIN tracking_status ts ON s.tracking_number = ts.tracking_number
    WHERE s.tracking_number != ''
      AND (ts.status IS NULL OR ts.status NOT IN ('Delivered', 'Return to Sender'))
      AND (ts.last_checked IS NULL OR ts.last_checked < datetime('now', '-4 hours'))
    ORDER BY ts.last_checked ASC NULLS FIRST
    LIMIT ?
  `).all(batchSize);

  const upsert = db.prepare(`
    INSERT INTO tracking_status (tracking_number, status, description, location, last_checked)
    VALUES (?, ?, ?, ?, datetime('now'))
    ON CONFLICT(tracking_number) DO UPDATE SET
      status = excluded.status, description = excluded.description,
      location = excluded.location, last_checked = excluded.last_checked
  `);

  // Log which carriers are configured
  const hasUPS  = !!(process.env.UPS_CLIENT_ID && process.env.UPS_CLIENT_SECRET);
  const hasUSPS = !!process.env.USPS_USER_ID;
  console.log(`[tracking] config: UPS=${hasUPS ? '✓' : '✗ (need UPS_CLIENT_ID+UPS_CLIENT_SECRET)'} USPS=${hasUSPS ? '✓' : '✗ (need USPS_USER_ID)'} Amazon=✗ (no public API)`);
  console.log(`[tracking] checking ${rows.length} packages...`);

  let checked = 0, updated = 0, skipped = 0;
  for (const row of rows) {
    try {
      const carrier = normalizeCarrier(row.carrier, row.tracking_number);
      let result = null;

      if (carrier === 'UPS' && hasUPS) {
        result = await checkUPS(row.tracking_number);
      } else if (carrier === 'USPS' && hasUSPS) {
        result = await checkUSPS(row.tracking_number);
      } else if (carrier === 'Amazon') {
        // Amazon TBA: no public API — skip (don't overwrite with Unknown)
        skipped++;
        continue;
      } else if (!hasUPS && carrier === 'UPS') {
        skipped++;
        continue;
      } else if (!hasUSPS && carrier === 'USPS') {
        skipped++;
        continue;
      }

      const status = result?.status || 'Unknown';
      upsert.run(row.tracking_number, status, result?.description || '', result?.location || '');
      if (result) updated++;
      checked++;
      await new Promise(r => setTimeout(r, 300)); // rate limit delay between requests
    } catch (e) {
      console.error(`tracking check failed for ${row.tracking_number}:`, e.message);
    }
  }

  // Update order statuses based on their tracking results
  updateOrderStatuses();
  console.log(`[tracking] done: checked=${checked} updated=${updated} skipped=${skipped}`);
  return { checked, updated, skipped, remaining: rows.length };
}

// ─── in-memory jobs ───────────────────────────────────────────────────────────
const jobs = {};

// ─── POST /api/upload ─────────────────────────────────────────────────────────
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const jobId = crypto.randomBytes(8).toString('hex');
  jobs[jobId] = { status: 'processing', progress: 0 };
  res.json({ success: true, jobId });

  setImmediate(() => {
    const job = jobs[jobId];
    try {
      const ext  = path.extname(req.file.originalname).toLowerCase();
      let rows;
      if (ext === '.txt' || ext === '.tsv' || ext === '.csv') {
        // Use plain-text parser to preserve full tracking numbers
        rows = parsePlainText(req.file.path);
      } else {
        const wb   = XLSX.readFile(req.file.path, { type: 'file', raw: false });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const raw   = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        rows = raw.map(r => Object.fromEntries(Object.entries(r).map(([k,v]) => [k.trim(), v])));
      }
      console.log(`[${jobId}] rows=${rows.length} (${ext})`);

      let added = 0, skipped = 0;
      db.transaction(rows => {
        for (const row of rows) {
          const r = insertShipment.run(buildRecord(row));
          if (r.changes > 0) added++; else skipped++;
        }
      })(rows);

      // Auto-create order metadata rows for new order IDs
      const orderIds = [...new Set(rows.map(r => String(r['order-id'] || r['Order ID'] || '').trim()).filter(Boolean))];
      for (const oid of orderIds) {
        db.prepare(`INSERT OR IGNORE INTO removal_orders (order_id) VALUES (?)`).run(oid);
      }

      // Auto-dedup: remove any duplicates introduced by this upload
      const dupsRemoved = runDedup();
      if (dupsRemoved > 0) {
        skipped += dupsRemoved;
        console.log(`[${jobId}] auto-dedup removed ${dupsRemoved} duplicate rows`);
      }

      try {
        const fh = fileHash(req.file.path);
        db.prepare('INSERT OR IGNORE INTO upload_log (filename,file_hash,rows_added,rows_skipped) VALUES (?,?,?,?)')
          .run(req.file.originalname, fh, added, skipped);
      } catch(_) {}

      job.status = 'done'; job.added = added; job.skipped = skipped; job.dupsRemoved = dupsRemoved; job.progress = 100;
      console.log(`[${jobId}] done — +${added} added, ${skipped} skipped, ${dupsRemoved} dups removed`);
    } catch(e) {
      console.error(`[${jobId}] error:`, e.message);
      job.status = 'error'; job.error = e.message;
    } finally {
      try { fs.unlinkSync(req.file.path); } catch(_) {}
    }
  });
});

// ─── GET /api/job/:id ─────────────────────────────────────────────────────────
app.get('/api/job/:id', (req, res) => {
  const job = jobs[req.params.id];
  if (!job) return res.status(404).json({ error: 'not found' });
  res.json(job);
});

// ─── GET /api/stats ───────────────────────────────────────────────────────────
app.get('/api/stats', (req, res) => {
  const totals = db.prepare(`
    SELECT
      COUNT(DISTINCT order_id)    as total_orders,
      COUNT(DISTINCT tracking_number) as total_packages,
      SUM(shipped_quantity)       as total_units,
      COUNT(DISTINCT sku)         as total_skus
    FROM removal_shipments WHERE order_id != ''
  `).get();

  const byCarrier = db.prepare(`
    SELECT
      CASE
        WHEN tracking_number LIKE 'TBA%'            THEN 'Amazon'
        WHEN tracking_number LIKE '1Z%'             THEN 'UPS'
        WHEN tracking_number GLOB '9[0-9][0-9][0-9]*' THEN 'USPS'
        WHEN carrier LIKE '%UPS%'                   THEN 'UPS'
        WHEN carrier LIKE '%USPS%'                  THEN 'USPS'
        WHEN carrier LIKE '%AMZL%'                  THEN 'Amazon'
        WHEN carrier LIKE '%EXLA%'                  THEN 'Amazon'
        WHEN carrier LIKE '%ABNT%'                  THEN 'Amazon'
        WHEN carrier LIKE '%FEDEX%'                 THEN 'FedEx'
        ELSE carrier
      END as carrier_name,
      COUNT(DISTINCT tracking_number) as packages,
      SUM(shipped_quantity) as units
    FROM removal_shipments
    WHERE order_id != ''
    GROUP BY carrier_name ORDER BY packages DESC
  `).all();

  const byDisposition = db.prepare(`
    SELECT disposition, SUM(shipped_quantity) as units, COUNT(*) as rows
    FROM removal_shipments WHERE disposition != ''
    GROUP BY disposition ORDER BY units DESC
  `).all();

  const byStatus = db.prepare(`
    SELECT status, COUNT(*) as cnt FROM removal_orders GROUP BY status ORDER BY cnt DESC
  `).all();

  const recentOrders = db.prepare(`
    SELECT s.order_id, MIN(s.request_date) as request_date,
      MAX(s.shipment_date) as last_shipment,
      COUNT(DISTINCT s.tracking_number) as packages,
      SUM(s.shipped_quantity) as units,
      o.destination, o.status
    FROM removal_shipments s
    LEFT JOIN removal_orders o ON s.order_id = o.order_id
    WHERE s.order_id != ''
    GROUP BY s.order_id ORDER BY request_date DESC LIMIT 5
  `).all();

  const trackingStatusSummary = db.prepare(`
    SELECT ts.status, COUNT(DISTINCT s.tracking_number) as count
    FROM removal_shipments s
    LEFT JOIN tracking_status ts ON s.tracking_number = ts.tracking_number
    WHERE s.tracking_number != ''
    GROUP BY ts.status ORDER BY count DESC
  `).all();

  res.json({ totals, byCarrier, byDisposition, byStatus, recentOrders, trackingStatusSummary });
});

// ─── GET /api/orders ─────────────────────────────────────────────────────────
app.get('/api/orders', (req, res) => {
  const page   = parseInt(req.query.page) || 1;
  const limit  = 50;
  const offset = (page - 1) * limit;
  const search = req.query.search || '';
  const status = req.query.status || '';

  let where = `WHERE s.order_id != ''`;
  const params = [];
  if (search) { where += ` AND (s.order_id LIKE ? OR o.destination LIKE ? OR s.tracking_number LIKE ?)`; params.push(`%${search}%`,`%${search}%`,`%${search}%`); }
  if (status) { where += ` AND o.status = ?`; params.push(status); }

  const total = db.prepare(`
    SELECT COUNT(DISTINCT s.order_id) as n
    FROM removal_shipments s LEFT JOIN removal_orders o ON s.order_id = o.order_id ${where}
  `).get(...params).n;

  const orders = db.prepare(`
    SELECT s.order_id,
      MIN(s.request_date)  as request_date,
      MAX(s.shipment_date) as last_shipment,
      COUNT(DISTINCT s.tracking_number) as packages,
      SUM(s.shipped_quantity) as units,
      COUNT(DISTINCT s.sku) as skus,
      GROUP_CONCAT(DISTINCT CASE WHEN s.tracking_number LIKE 'TBA%' THEN 'Amazon'
        WHEN s.tracking_number LIKE '1Z%' THEN 'UPS'
        WHEN s.tracking_number GLOB '9[0-9]*' THEN 'USPS'
        WHEN s.carrier LIKE '%UPS%' THEN 'UPS'
        WHEN s.carrier LIKE '%USPS%' THEN 'USPS'
        WHEN s.carrier LIKE '%AMZL%' THEN 'Amazon'
        ELSE s.carrier END) as carriers,
      o.destination, o.status, o.notes
    FROM removal_shipments s
    LEFT JOIN removal_orders o ON s.order_id = o.order_id
    ${where}
    GROUP BY s.order_id ORDER BY request_date DESC
    LIMIT ? OFFSET ?
  `).all(...params, limit, offset);

  res.json({ orders, total, page, pages: Math.ceil(total / limit) });
});

// ─── GET /api/order/:id — full detail of one order ───────────────────────────
app.get('/api/order/:id', (req, res) => {
  const orderId = req.params.id;
  const meta = db.prepare(`SELECT * FROM removal_orders WHERE order_id = ?`).get(orderId) || { order_id: orderId, destination: '', status: 'In Transit', notes: '' };

  // Group by tracking number
  const packages = db.prepare(`
    SELECT tracking_number, carrier, MIN(shipment_date) as shipment_date,
      SUM(shipped_quantity) as units, COUNT(*) as item_count
    FROM removal_shipments WHERE order_id = ? AND tracking_number != ''
    GROUP BY tracking_number ORDER BY shipment_date ASC
  `).all(orderId);

  // All items
  const items = db.prepare(`
    SELECT sku, fnsku, disposition, shipped_quantity, carrier, tracking_number, shipment_date
    FROM removal_shipments WHERE order_id = ?
    ORDER BY tracking_number, shipment_date
  `).all(orderId);

  // Add tracking URLs to packages
  const packagesWithUrls = packages.map(p => ({
    ...p,
    carrier_name: normalizeCarrier(p.carrier, p.tracking_number),
    tracking_url: trackingUrl(p.carrier, p.tracking_number),
  }));

  res.json({ meta, packages: packagesWithUrls, items });
});

// ─── PATCH /api/order/:id — update destination / status / notes ──────────────
app.patch('/api/order/:id', (req, res) => {
  const orderId = req.params.id;
  const { destination, status, notes } = req.body;
  db.prepare(`INSERT OR IGNORE INTO removal_orders (order_id) VALUES (?)`).run(orderId);
  db.prepare(`UPDATE removal_orders SET destination=COALESCE(?,destination), status=COALESCE(?,status), notes=COALESCE(?,notes), updated_at=datetime('now') WHERE order_id=?`)
    .run(destination ?? null, status ?? null, notes ?? null, orderId);
  res.json({ success: true });
});

// ─── GET /api/tracking — all unique tracking numbers ─────────────────────────
app.get('/api/tracking', (req, res) => {
  const page   = parseInt(req.query.page) || 1;
  const limit  = 50;
  const offset = (page - 1) * limit;
  const search = req.query.search || '';
  const carrier= req.query.carrier || '';

  let where = `WHERE s.tracking_number != ''`;
  const params = [];
  if (search)  { where += ` AND (s.tracking_number LIKE ? OR s.order_id LIKE ?)`; params.push(`%${search}%`, `%${search}%`); }
  if (carrier) {
    if (carrier === 'UPS')    where += ` AND s.tracking_number LIKE '1Z%'`;
    if (carrier === 'USPS')   where += ` AND s.tracking_number GLOB '9[0-9]*'`;
    if (carrier === 'Amazon') where += ` AND s.tracking_number LIKE 'TBA%'`;
  }

  const total = db.prepare(`SELECT COUNT(DISTINCT s.tracking_number) as n FROM removal_shipments s ${where}`).get(...params).n;

  const rows = db.prepare(`
    SELECT s.tracking_number, s.order_id, s.carrier,
      MIN(s.shipment_date) as shipment_date,
      SUM(s.shipped_quantity) as units,
      COUNT(*) as item_count,
      GROUP_CONCAT(DISTINCT s.disposition) as dispositions,
      ts.status as tracking_status,
      ts.location,
      ts.last_checked
    FROM removal_shipments s
    LEFT JOIN tracking_status ts ON s.tracking_number = ts.tracking_number
    ${where}
    GROUP BY s.tracking_number ORDER BY s.shipment_date DESC
    LIMIT ? OFFSET ?
  `).all(...params, limit, offset);

  const result = rows.map(r => ({
    ...r,
    carrier_name: normalizeCarrier(r.carrier, r.tracking_number),
    tracking_url: trackingUrl(r.carrier, r.tracking_number),
  }));

  res.json({ rows: result, total, page, pages: Math.ceil(total / limit) });
});

// ─── GET /api/tracking/refresh — manually trigger a tracking check ────────────
app.get('/api/tracking/refresh', async (req, res) => {
  const batch = parseInt(req.query.batch) || 100;
  const hasUPS  = !!(process.env.UPS_CLIENT_ID && process.env.UPS_CLIENT_SECRET);
  const hasUSPS = !!process.env.USPS_USER_ID;
  const configured = [];
  if (hasUPS)  configured.push('UPS');
  if (hasUSPS) configured.push('USPS');
  const msg = configured.length
    ? `Checking up to ${batch} ${configured.join('+')} tracking numbers in background...`
    : `No carrier credentials configured. Set UPS_CLIENT_ID+UPS_CLIENT_SECRET (UPS) or USPS_USER_ID (USPS) in Railway env vars.`;
  res.json({ success: configured.length > 0, message: msg, configured });
  if (configured.length > 0) {
    refreshTracking(batch).catch(e => console.error('refresh error:', e.message));
  }
});

// ─── GET /api/tracking/status-summary — counts by status ─────────────────────
app.get('/api/tracking/status-summary', (req, res) => {
  const summary = db.prepare(`
    SELECT status, COUNT(*) as count
    FROM tracking_status
    GROUP BY status ORDER BY count DESC
  `).all();
  const lastChecked = db.prepare(`SELECT MAX(last_checked) as ts FROM tracking_status`).get()?.ts;
  const totalTracking = db.prepare(`SELECT COUNT(DISTINCT tracking_number) as n FROM removal_shipments WHERE tracking_number != ''`).get().n;
  const checkedCount  = db.prepare(`SELECT COUNT(*) as n FROM tracking_status`).get().n;
  const hasUPS  = !!(process.env.UPS_CLIENT_ID && process.env.UPS_CLIENT_SECRET);
  const hasUSPS = !!process.env.USPS_USER_ID;
  res.json({ summary, lastChecked, totalTracking, checkedCount,
    config: { ups: hasUPS, usps: hasUSPS, amazon: false } });
});

// ─── GET /api/tracking/:tn — items in one tracking number ────────────────────
app.get('/api/tracking/:tn', (req, res) => {
  const tn = req.params.tn;
  const items = db.prepare(`
    SELECT sku, fnsku, disposition, shipped_quantity, order_id, shipment_date
    FROM removal_shipments WHERE tracking_number = ?
  `).all(tn);
  const carrier = items[0]?.carrier || '';
  res.json({ tracking_number: tn, carrier_name: normalizeCarrier(carrier, tn), tracking_url: trackingUrl(carrier, tn), items });
});

// ─── POST /api/upload-addresses ──────────────────────────────────────────────
app.post('/api/upload-addresses', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const ext = path.extname(req.file.originalname).toLowerCase();
    let rows;
    if (ext === '.txt' || ext === '.tsv' || ext === '.csv') {
      rows = parsePlainText(req.file.path);
    } else {
      const wb    = XLSX.readFile(req.file.path, { type: 'file', raw: false });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const raw   = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      rows = raw.map(r => Object.fromEntries(Object.entries(r).map(([k,v]) => [k.trim(), v])));
    }

    let updated = 0, created = 0;
    // Build map: order_id → { address, order_type, order_status }
    // Aggregate per order_id: sum shipped/requested qty, keep first address
    const orderMap = {};
    for (const row of rows) {
      const orderId = String(row['Order ID'] || row['order-id'] || row['order_id'] || '').trim();
      if (!orderId) continue;
      const addr  = String(row[' Shipping Address'] || row['Shipping Address'] || row['shipping_address'] || row['shipping-address'] || '').trim()
                      .replace(/\r\n/g, '\n').replace(/\r/g, '\n');
      const oType = String(row['Order Type']   || row['order-type']   || '').trim();
      const oStat = String(row['Order Status'] || row['order-status'] || '').trim();
      const shipQty = parseInt(row['Shipped Quantity']   || row['shipped-quantity']   || row['Shipped_Quantity']   || 0) || 0;
      const reqQty  = parseInt(row['Requested Quantity'] || row['requested-quantity'] || row['Requested_Quantity'] || 0) || 0;

      if (!orderMap[orderId]) {
        orderMap[orderId] = { addr, oType, oStat, shipQty: 0, reqQty: 0 };
      }
      // Sum quantities across all SKU rows for this order
      orderMap[orderId].shipQty += shipQty;
      orderMap[orderId].reqQty  += reqQty;
      // Keep first non-empty address/type/status
      if (!orderMap[orderId].addr  && addr)  orderMap[orderId].addr  = addr;
      if (!orderMap[orderId].oType && oType) orderMap[orderId].oType = oType;
      if (!orderMap[orderId].oStat && oStat) orderMap[orderId].oStat = oStat;
    }

    const upsert = db.prepare(`INSERT INTO removal_orders
        (order_id, shipping_address, order_type, order_status, addr_shipped_qty, addr_requested_qty)
      VALUES (?, ?, ?, ?, ?, ?)
      ON CONFLICT(order_id) DO UPDATE SET
        shipping_address   = excluded.shipping_address,
        order_type         = CASE WHEN excluded.order_type  != '' THEN excluded.order_type  ELSE order_type  END,
        order_status       = CASE WHEN excluded.order_status!= '' THEN excluded.order_status ELSE order_status END,
        addr_shipped_qty   = excluded.addr_shipped_qty,
        addr_requested_qty = excluded.addr_requested_qty,
        updated_at         = datetime('now')`);

    db.transaction(map => {
      for (const [orderId, info] of Object.entries(map)) {
        const existing = db.prepare('SELECT id FROM removal_orders WHERE order_id = ?').get(orderId);
        upsert.run(orderId, info.addr, info.oType, info.oStat, info.shipQty, info.reqQty);
        if (existing) updated++; else created++;
      }
    })(orderMap);

    try { fs.unlinkSync(req.file.path); } catch(_) {}
    res.json({ success: true, updated, created, total: Object.keys(orderMap).length });
  } catch(e) {
    try { fs.unlinkSync(req.file.path); } catch(_) {}
    res.status(500).json({ error: e.message });
  }
});

// ─── GET /api/addresses — grouped by shipping address ────────────────────────
app.get('/api/addresses', (req, res) => {
  const search  = req.query.search   || '';
  const orderId = req.query.order_id || '';
  const fnsku   = req.query.fnsku    || '';

  // When filtering by order_id or fnsku we need to join removal_shipments
  const needsJoin = !!(orderId || fnsku);

  let where  = `WHERE o.shipping_address != ''`;
  const params = [];

  if (search) {
    where += ` AND (o.shipping_address LIKE ? OR o.order_id LIKE ?)`;
    params.push(`%${search}%`, `%${search}%`);
  }
  if (orderId) {
    where += ` AND o.order_id LIKE ?`;
    params.push(`%${orderId}%`);
  }
  if (fnsku) {
    // Filter to addresses that have at least one shipment with matching fnsku
    where += ` AND o.order_id IN (SELECT DISTINCT order_id FROM removal_shipments WHERE fnsku LIKE ?)`;
    params.push(`%${fnsku}%`);
  }

  // Query removal_orders only (no JOIN) to avoid row multiplication on SUM
  const rows = db.prepare(`
    SELECT
      o.shipping_address,
      COUNT(DISTINCT o.order_id)           as order_count,
      COALESCE(SUM(o.addr_shipped_qty),  0) as total_units,
      COALESCE(SUM(o.addr_requested_qty),0) as total_requested
    FROM removal_orders o
    ${where}
    GROUP BY o.shipping_address
    ORDER BY total_units DESC
  `).all(...params);

  // For each address, get package count + per-order breakdown separately
  const pkgStmt = db.prepare(`
    SELECT COUNT(DISTINCT s.tracking_number) as n
    FROM removal_shipments s
    JOIN removal_orders o ON s.order_id = o.order_id
    WHERE o.shipping_address = ?
  `);

  const ordersStmt = db.prepare(`
    SELECT o.order_id,
      o.addr_shipped_qty   as units,
      o.addr_requested_qty as requested,
      o.order_status,
      o.status,
      o.updated_at         as request_date,
      (SELECT COUNT(DISTINCT tracking_number)
       FROM removal_shipments WHERE order_id = o.order_id) as packages
    FROM removal_orders o
    WHERE o.shipping_address = ? AND o.order_id != ''
    ORDER BY o.updated_at DESC
  `);

  const result = rows.map(r => ({
    ...r,
    package_count: pkgStmt.get(r.shipping_address)?.n || 0,
    orders: ordersStmt.all(r.shipping_address),
  }));

  res.json(result);
});

// ─── GET /api/by-fnsku — grouped by FNSKU ────────────────────────────────────
app.get('/api/by-fnsku', (req, res) => {
  const search  = req.query.search   || '';
  const orderId = req.query.order_id || '';
  const address = req.query.address  || '';

  let where = `WHERE s.fnsku != ''`;
  const params = [];
  if (search) {
    where += ` AND (s.fnsku LIKE ? OR s.sku LIKE ?)`;
    params.push(`%${search}%`, `%${search}%`);
  }
  if (orderId) {
    where += ` AND s.order_id LIKE ?`;
    params.push(`%${orderId}%`);
  }
  if (address) {
    // Try exact match first (from dropdown), fall back to LIKE for text search
    where += ` AND s.order_id IN (SELECT order_id FROM removal_orders WHERE shipping_address = ? OR shipping_address LIKE ?)`;
    params.push(address, `%${address}%`);
  }

  const rows = db.prepare(`
    SELECT
      s.fnsku,
      MAX(s.sku)                              AS sku,
      SUM(s.shipped_quantity)                 AS total_units,
      COUNT(DISTINCT s.order_id)              AS order_count,
      COUNT(DISTINCT s.tracking_number)       AS package_count,
      GROUP_CONCAT(DISTINCT s.disposition)    AS dispositions,
      MAX(s.shipment_date)                    AS last_shipped
    FROM removal_shipments s
    ${where}
    GROUP BY s.fnsku
    ORDER BY total_units DESC
  `).all(...params);

  res.json(rows);
});

// ─── GET /api/fnsku-items — all shipment rows for a given FNSKU ───────────────
app.get('/api/fnsku-items', (req, res) => {
  const fnsku   = req.query.fnsku    || '';
  const orderId = req.query.order_id || '';
  const address = req.query.address  || '';
  if (!fnsku) return res.json([]);

  let where = `WHERE s.fnsku = ?`;
  const params = [fnsku];
  if (orderId) { where += ` AND s.order_id LIKE ?`; params.push(`%${orderId}%`); }
  if (address) {
    where += ` AND s.order_id IN (SELECT order_id FROM removal_orders WHERE shipping_address = ? OR shipping_address LIKE ?)`;
    params.push(address, `%${address}%`);
  }

  const items = db.prepare(`
    SELECT
      s.order_id,
      s.tracking_number,
      s.carrier,
      s.disposition,
      s.shipped_quantity,
      s.shipment_date,
      o.shipping_address
    FROM removal_shipments s
    LEFT JOIN removal_orders o ON s.order_id = o.order_id
    ${where}
    ORDER BY s.order_id, s.shipment_date
  `).all(...params);

  const result = items.map(r => ({
    ...r,
    carrier_name: normalizeCarrier(r.carrier, r.tracking_number),
    tracking_url: trackingUrl(r.carrier, r.tracking_number),
  }));

  res.json(result);
});

// ─── GET /api/address-items — all FNSKUs shipped to a given address ───────────
app.get('/api/address-items', (req, res) => {
  const address = req.query.address  || '';
  const orderId = req.query.order_id || '';
  const fnsku   = req.query.fnsku    || '';
  if (!address) return res.json([]);

  let where = `WHERE o.shipping_address = ?`;
  const params = [address];
  if (orderId) { where += ` AND s.order_id LIKE ?`;  params.push(`%${orderId}%`); }
  if (fnsku)   { where += ` AND s.fnsku    LIKE ?`;  params.push(`%${fnsku}%`);   }

  const items = db.prepare(`
    SELECT
      s.fnsku,
      s.sku,
      s.order_id,
      s.tracking_number,
      s.carrier,
      s.disposition,
      s.shipped_quantity,
      s.shipment_date,
      s.request_date
    FROM removal_shipments s
    JOIN removal_orders o ON s.order_id = o.order_id
    ${where}
    ORDER BY s.order_id, s.tracking_number, s.fnsku
  `).all(...params);

  // Attach carrier name and tracking URL
  const result = items.map(r => ({
    ...r,
    carrier_name: normalizeCarrier(r.carrier, r.tracking_number),
    tracking_url: trackingUrl(r.carrier, r.tracking_number),
  }));

  res.json(result);
});

// ─── GET /api/export/address-items — CSV download of all address+order+FNSKU ──
app.get('/api/export/address-items', (req, res) => {
  const search  = req.query.search   || '';
  const orderId = req.query.order_id || '';
  const fnsku   = req.query.fnsku    || '';

  let where = `WHERE o.shipping_address != ''`;
  const params = [];
  if (search)  { where += ` AND (o.shipping_address LIKE ? OR s.order_id LIKE ?)`; params.push(`%${search}%`, `%${search}%`); }
  if (orderId) { where += ` AND s.order_id LIKE ?`; params.push(`%${orderId}%`); }
  if (fnsku)   { where += ` AND s.fnsku    LIKE ?`; params.push(`%${fnsku}%`); }

  const rows = db.prepare(`
    SELECT
      o.shipping_address,
      s.order_id,
      s.fnsku,
      s.sku,
      s.tracking_number,
      s.carrier,
      s.shipped_quantity,
      s.disposition,
      s.shipment_date,
      s.request_date
    FROM removal_shipments s
    JOIN removal_orders o ON s.order_id = o.order_id
    ${where}
    ORDER BY o.shipping_address, s.order_id, s.fnsku
  `).all(...params);

  // Build CSV
  const csvEsc = v => {
    const s = String(v ?? '').replace(/\r\n/g, ' ').replace(/\n/g, ' ');
    return s.includes(',') || s.includes('"') || s.includes('\n') ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const headers = ['Shipping Address','Order ID','FNSKU','SKU','Tracking Number','Carrier','Shipped Quantity','Disposition','Shipment Date','Request Date'];
  const lines = [
    headers.join(','),
    ...rows.map(r => [
      r.shipping_address, r.order_id, r.fnsku, r.sku,
      r.tracking_number, r.carrier, r.shipped_quantity,
      r.disposition, r.shipment_date, r.request_date
    ].map(csvEsc).join(','))
  ];

  const filename = `address-shipments-${new Date().toISOString().slice(0,10)}.csv`;
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(lines.join('\r\n'));
});

// ─── GET /api/dropdown-options — lightweight lists for filter dropdowns ───────
app.get('/api/dropdown-options', (req, res) => {
  const orderIds = db.prepare(`
    SELECT DISTINCT order_id FROM removal_shipments
    WHERE order_id != '' ORDER BY order_id ASC
  `).all().map(r => r.order_id);

  const addresses = db.prepare(`
    SELECT DISTINCT shipping_address FROM removal_orders
    WHERE shipping_address != '' ORDER BY shipping_address ASC
  `).all().map(r => r.shipping_address);

  res.json({ orderIds, addresses });
});

// ─── GET /api/uploads ─────────────────────────────────────────────────────────
app.get('/api/uploads', (req, res) => {
  const logs = db.prepare('SELECT * FROM upload_log ORDER BY uploaded_at DESC LIMIT 20').all();
  res.json(logs);
});

// ─── GET /api/dedup — find & remove duplicate shipment rows ──────────────────
app.get('/api/dedup', (req, res) => {
  const count = runDedup();
  const total = db.prepare('SELECT COUNT(*) as n FROM removal_shipments').get().n;
  res.send(`
    <html><body style="font-family:sans-serif;padding:32px">
    <h2>Dedup Complete</h2>
    <p>Duplicate rows removed: <strong>${count}</strong></p>
    <p>Rows remaining in DB: <strong>${total.toLocaleString()}</strong></p>
    <p style="color:green">✅ ${count === 0 ? 'No duplicates found — data is clean.' : `${count} duplicate(s) deleted.`}</p>
    <a href="/">← Back to app</a>
    </body></html>
  `);
});

// ─── GET /api/diagnostic — row counts and data health check ──────────────────
app.get('/api/diagnostic', (req, res) => {
  const total     = db.prepare('SELECT COUNT(*) as n FROM removal_shipments').get().n;
  const orders    = db.prepare('SELECT COUNT(DISTINCT order_id) as n FROM removal_shipments WHERE order_id != ""').get().n;
  const tracking  = db.prepare('SELECT COUNT(DISTINCT tracking_number) as n FROM removal_shipments WHERE tracking_number != ""').get().n;
  const noTracking= db.prepare('SELECT COUNT(*) as n FROM removal_shipments WHERE tracking_number = ""').get().n;
  const dupeCheck = db.prepare(`
    SELECT COUNT(*) as n FROM removal_shipments
    WHERE id NOT IN (
      SELECT MIN(id) FROM removal_shipments
      GROUP BY order_id, sku, tracking_number, shipment_date
    )
  `).get().n;
  const dateRange = db.prepare(`
    SELECT MIN(shipment_date) as earliest, MAX(shipment_date) as latest
    FROM removal_shipments WHERE shipment_date != ''
  `).get();
  const uploads   = db.prepare('SELECT COUNT(*) as n FROM upload_log').get().n;

  res.send(`
    <html><body style="font-family:sans-serif;padding:32px;max-width:600px">
    <h2>Database Diagnostic</h2>
    <table style="border-collapse:collapse;width:100%">
      <tr style="background:#f3f4f6"><td style="padding:8px 12px;font-weight:600">Total rows</td><td style="padding:8px 12px">${total.toLocaleString()}</td></tr>
      <tr><td style="padding:8px 12px;font-weight:600">Unique orders</td><td style="padding:8px 12px">${orders.toLocaleString()}</td></tr>
      <tr style="background:#f3f4f6"><td style="padding:8px 12px;font-weight:600">Unique tracking numbers</td><td style="padding:8px 12px">${tracking.toLocaleString()}</td></tr>
      <tr><td style="padding:8px 12px;font-weight:600">Rows without tracking</td><td style="padding:8px 12px">${noTracking.toLocaleString()}</td></tr>
      <tr style="background:#f3f4f6"><td style="padding:8px 12px;font-weight:600">Duplicate rows</td><td style="padding:8px 12px;color:${dupeCheck>0?'red':'green'}">${dupeCheck} ${dupeCheck>0?'⚠️ run /api/dedup to clean':'✅ clean'}</td></tr>
      <tr><td style="padding:8px 12px;font-weight:600">Shipment date range</td><td style="padding:8px 12px">${dateRange.earliest||'—'} → ${dateRange.latest||'—'}</td></tr>
      <tr style="background:#f3f4f6"><td style="padding:8px 12px;font-weight:600">Upload files logged</td><td style="padding:8px 12px">${uploads}</td></tr>
    </table>
    <br/><a href="/api/dedup">Run Dedup →</a> &nbsp; <a href="/">← Back to app</a>
    </body></html>
  `);
});

// ─── DELETE /api/reset ────────────────────────────────────────────────────────
app.delete('/api/reset', (req, res) => {
  db.prepare('DELETE FROM removal_shipments').run();
  db.prepare('DELETE FROM removal_orders').run();
  db.prepare('DELETE FROM upload_log').run();
  res.json({ success: true });
});

// ─── Auto-refresh tracking status every 4 hours ───────────────────────────────
setInterval(() => {
  console.log('[tracking] auto-refresh starting...');
  refreshTracking(100).catch(e => console.error('[tracking] auto-refresh error:', e.message));
}, 4 * 60 * 60 * 1000);

// Run once 30 seconds after startup
setTimeout(() => {
  console.log('[tracking] initial tracking check...');
  refreshTracking(100).catch(e => console.error('[tracking] initial check error:', e.message));
}, 30000);

// ─── start ────────────────────────────────────────────────────────────────────
const server = app.listen(PORT, () => {
  console.log(`\n✅ Removal Tracker running at http://localhost:${PORT}\n`);
});
server.setTimeout(600000);
