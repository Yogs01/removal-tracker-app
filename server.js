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

      try {
        const fh = fileHash(req.file.path);
        db.prepare('INSERT OR IGNORE INTO upload_log (filename,file_hash,rows_added,rows_skipped) VALUES (?,?,?,?)')
          .run(req.file.originalname, fh, added, skipped);
      } catch(_) {}

      job.status = 'done'; job.added = added; job.skipped = skipped; job.progress = 100;
      console.log(`[${jobId}] done — +${added} added, ${skipped} skipped`);
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
        WHEN tracking_number LIKE 'TBA%' THEN 'Amazon'
        WHEN tracking_number LIKE '1Z%'  THEN 'UPS'
        WHEN tracking_number GLOB '9[0-9][0-9][0-9]*' THEN 'USPS'
        WHEN carrier LIKE '%UPS%'   THEN 'UPS'
        WHEN carrier LIKE '%USPS%'  THEN 'USPS'
        WHEN carrier LIKE '%AMZL%'  THEN 'Amazon'
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

  res.json({ totals, byCarrier, byDisposition, byStatus, recentOrders });
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

  let where = `WHERE tracking_number != ''`;
  const params = [];
  if (search)  { where += ` AND (tracking_number LIKE ? OR order_id LIKE ?)`; params.push(`%${search}%`, `%${search}%`); }
  if (carrier) {
    if (carrier === 'UPS')    where += ` AND tracking_number LIKE '1Z%'`;
    if (carrier === 'USPS')   where += ` AND tracking_number GLOB '9[0-9]*'`;
    if (carrier === 'Amazon') where += ` AND tracking_number LIKE 'TBA%'`;
  }

  const total = db.prepare(`SELECT COUNT(DISTINCT tracking_number) as n FROM removal_shipments ${where}`).get(...params).n;

  const rows = db.prepare(`
    SELECT tracking_number, order_id, carrier,
      MIN(shipment_date) as shipment_date,
      SUM(shipped_quantity) as units,
      COUNT(*) as item_count,
      GROUP_CONCAT(DISTINCT disposition) as dispositions
    FROM removal_shipments ${where}
    GROUP BY tracking_number ORDER BY shipment_date DESC
    LIMIT ? OFFSET ?
  `).all(...params, limit, offset);

  const result = rows.map(r => ({
    ...r,
    carrier_name: normalizeCarrier(r.carrier, r.tracking_number),
    tracking_url: trackingUrl(r.carrier, r.tracking_number),
  }));

  res.json({ rows: result, total, page, pages: Math.ceil(total / limit) });
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

// ─── GET /api/uploads ─────────────────────────────────────────────────────────
app.get('/api/uploads', (req, res) => {
  const logs = db.prepare('SELECT * FROM upload_log ORDER BY uploaded_at DESC LIMIT 20').all();
  res.json(logs);
});

// ─── DELETE /api/reset ────────────────────────────────────────────────────────
app.delete('/api/reset', (req, res) => {
  db.prepare('DELETE FROM removal_shipments').run();
  db.prepare('DELETE FROM removal_orders').run();
  db.prepare('DELETE FROM upload_log').run();
  res.json({ success: true });
});

// ─── start ────────────────────────────────────────────────────────────────────
const server = app.listen(PORT, () => {
  console.log(`\n✅ Removal Tracker running at http://localhost:${PORT}\n`);
});
server.setTimeout(600000);
