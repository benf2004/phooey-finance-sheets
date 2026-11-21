CREATE TABLE IF NOT EXISTS hosted_workbook (
 id INTEGER PRIMARY KEY CHECK(id=1), file_id TEXT, creation_key TEXT NOT NULL, refresh_secret TEXT NOT NULL,
 generation TEXT NOT NULL, grant_hash TEXT, lease TEXT, lease_until INTEGER NOT NULL DEFAULT 0,
 tax_year INTEGER CHECK(tax_year BETWEEN 2000 AND 2100), synced_at INTEGER, error TEXT
);
