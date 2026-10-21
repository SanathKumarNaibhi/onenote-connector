// Local copy of your OneNote: a SQLite database (full-text searchable) plus a folder of
// page files you can open in a browser. Everything lives in DATA_DIR (a persistent volume).

import { DatabaseSync } from "node:sqlite";
import { createCipheriv, createDecipheriv, createHash, randomBytes } from "node:crypto";
import fs from "node:fs";
import path from "node:path";

export function openStore(dataDir) {
  fs.mkdirSync(path.join(dataDir, "pages"), { recursive: true });
  const db = new DatabaseSync(path.join(dataDir, "onenote.db"));
  db.exec(`
    PRAGMA journal_mode = WAL;
    CREATE TABLE IF NOT EXISTS sections (
      id TEXT PRIMARY KEY, name TEXT, notebook TEXT, group_path TEXT,
      graph_modified TEXT, synced_modified TEXT, last_synced TEXT,
      status TEXT, error TEXT, page_count INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS pages (
      id TEXT PRIMARY KEY, section_id TEXT, title TEXT, notebook TEXT, section TEXT,
      created TEXT, graph_modified TEXT,
      first_seen TEXT, changed_at TEXT, fetched_at TEXT, deleted_at TEXT,
      hash TEXT, html TEXT, text TEXT, file_path TEXT, error TEXT
    );
    CREATE INDEX IF NOT EXISTS pages_section ON pages(section_id);
    CREATE INDEX IF NOT EXISTS pages_changed ON pages(changed_at);
    CREATE TABLE IF NOT EXISTS page_versions (
      page_id TEXT, saved_at TEXT, title TEXT, hash TEXT, html TEXT, text TEXT
    );
    CREATE INDEX IF NOT EXISTS versions_page ON page_versions(page_id);
    CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
    CREATE VIRTUAL TABLE IF NOT EXISTS page_fts USING fts5(
      page_id UNINDEXED, title, body, tokenize = 'trigram'
    );
  `);
  return db;
}

export const nowIso = () => new Date().toISOString();
export const sha = (s) => createHash("sha256").update(s).digest("hex");

export function getMeta(db, key) {
  return db.prepare("SELECT value FROM meta WHERE key = ?").get(key)?.value ?? null;
}
export function setMeta(db, key, value) {
  db.prepare("INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value").run(key, value);
}

// The Microsoft refresh token is kept encrypted with a key derived from the app secret,
// so the database file alone is not enough to reach your account.
const keyFrom = (secret) => createHash("sha256").update(`onenote-sync:${secret}`).digest();
export function saveRefreshToken(db, secret, token) {
  const iv = randomBytes(12);
  const c = createCipheriv("aes-256-gcm", keyFrom(secret), iv);
  const enc = Buffer.concat([c.update(token, "utf8"), c.final()]);
  setMeta(db, "refresh_token", [iv, c.getAuthTag(), enc].map((b) => b.toString("base64")).join("."));
  setMeta(db, "refresh_token_saved", nowIso());
}
export function loadRefreshToken(db, secret) {
  const v = getMeta(db, "refresh_token");
  if (!v) return null;
  try {
    const [iv, tag, enc] = v.split(".").map((s) => Buffer.from(s, "base64"));
    const d = createDecipheriv("aes-256-gcm", keyFrom(secret), iv);
    d.setAuthTag(tag);
    return Buffer.concat([d.update(enc), d.final()]).toString("utf8");
  } catch {
    return null; // app secret changed: wait for the next sign-in to re-arm
  }
}

// ----- Readable file mirror: pages/<Notebook>/<Section>/<Title> [id].html -----
const clean = (s, max = 80) =>
  (String(s || "untitled")
    .replace(/[\\/:*?"<>|\x00-\x1f]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max) || "untitled").replace(/^\.+/, "_");

export function writePageFile(dataDir, page, html) {
  const shortId = sha(page.id).slice(0, 8);
  const rel = path.join("pages", clean(page.notebook), clean(page.section), `${clean(page.title)} [${shortId}].html`);
  const abs = path.join(dataDir, rel);
  fs.mkdirSync(path.dirname(abs), { recursive: true });
  fs.writeFileSync(abs, html);
  return rel;
}
export function removeFile(dataDir, rel) {
  if (!rel) return;
  try {
    fs.unlinkSync(path.join(dataDir, rel));
  } catch {
    /* already gone */
  }
}

// ----- Search -----
// Trigram index: matches any 3+ character fragment in any script (English, Kannada, Sanskrit
// transliterations), so "practis" finds practice and practising.
export function searchIndex(db, query, limit) {
  const terms = query
    .split(/\s+/)
    .map((t) => t.trim())
    .filter(Boolean);
  const long = terms.filter((t) => [...t].length >= 3);
  const short = terms.filter((t) => [...t].length < 3);
  const cols = `p.id, p.title, p.notebook, p.section, p.section_id, p.created, p.changed_at, p.deleted_at`;
  let rows;
  if (long.length) {
    const match = long.map((t) => `"${t.replace(/"/g, '""')}"`).join(" ");
    rows = db
      .prepare(
        `SELECT ${cols}, snippet(page_fts, 2, '«', '»', ' … ', 64) AS snippet
         FROM page_fts JOIN pages p ON p.id = page_fts.page_id
         WHERE page_fts MATCH ? ORDER BY bm25(page_fts, 0, 8, 1) LIMIT ?`,
      )
      .all(match, short.length ? 500 : limit);
  } else {
    rows = db.prepare(`SELECT ${cols}, substr(p.text, 1, 200) AS snippet FROM pages p WHERE p.deleted_at IS NULL LIMIT 5000`).all();
  }
  if (short.length) {
    const lc = short.map((t) => t.toLowerCase());
    rows = rows.filter((r) => {
      const hay = `${r.title || ""}\n${db.prepare("SELECT text FROM pages WHERE id = ?").get(r.id)?.text || ""}`.toLowerCase();
      return lc.every((t) => hay.includes(t));
    });
  }
  return rows.slice(0, limit).map((r) => ({ ...r }));
}
