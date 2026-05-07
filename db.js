const Database = require('better-sqlite3');
const path = require('path');

const db = new Database(process.env.DB_PATH || path.join(__dirname, 'removal.db'));

db.exec(`
  CREATE TABLE IF NOT EXISTS removal_shipments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    request_date TEXT,
    order_id TEXT,
    shipment_date TEXT,
    sku TEXT,
    fnsku TEXT,
    disposition TEXT DEFAULT '',
    shipped_quantity INTEGER DEFAULT 1,
    carrier TEXT DEFAULT '',
    tracking_number TEXT DEFAULT '',
    removal_order_type TEXT DEFAULT '',
    row_hash TEXT UNIQUE
  );

  CREATE TABLE IF NOT EXISTS removal_orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id TEXT UNIQUE,
    destination TEXT DEFAULT '',
    status TEXT DEFAULT 'In Transit',
    notes TEXT DEFAULT '',
    updated_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS upload_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT,
    file_hash TEXT UNIQUE,
    rows_added INTEGER DEFAULT 0,
    rows_skipped INTEGER DEFAULT 0,
    uploaded_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS tracking_status (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tracking_number TEXT UNIQUE NOT NULL,
    status TEXT DEFAULT 'Unknown',
    description TEXT DEFAULT '',
    location TEXT DEFAULT '',
    last_checked TEXT DEFAULT NULL
  );

  CREATE INDEX IF NOT EXISTS idx_ship_order    ON removal_shipments(order_id);
  CREATE INDEX IF NOT EXISTS idx_ship_tracking ON removal_shipments(tracking_number);
  CREATE INDEX IF NOT EXISTS idx_ship_date     ON removal_shipments(shipment_date);
  CREATE INDEX IF NOT EXISTS idx_ship_carrier  ON removal_shipments(carrier);
  CREATE INDEX IF NOT EXISTS idx_ord_order     ON removal_orders(order_id);
  CREATE INDEX IF NOT EXISTS idx_ts_tracking   ON tracking_status(tracking_number);
  CREATE INDEX IF NOT EXISTS idx_ts_status     ON tracking_status(status);
`);

// Migrate: add columns if they don't exist yet
try { db.exec(`ALTER TABLE removal_orders ADD COLUMN shipping_address TEXT DEFAULT ''`); } catch(_) {}
try { db.exec(`ALTER TABLE removal_orders ADD COLUMN order_type TEXT DEFAULT ''`); } catch(_) {}
try { db.exec(`ALTER TABLE removal_orders ADD COLUMN order_status TEXT DEFAULT ''`); } catch(_) {}
try { db.exec(`ALTER TABLE removal_orders ADD COLUMN addr_shipped_qty INTEGER DEFAULT 0`); } catch(_) {}
try { db.exec(`ALTER TABLE removal_orders ADD COLUMN addr_requested_qty INTEGER DEFAULT 0`); } catch(_) {}

module.exports = db;
