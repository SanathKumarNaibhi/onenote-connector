// Background sync: keeps a full local copy of every OneNote page (HTML + extracted text),
// read-only towards Microsoft. Runs on a timer and on demand.
//
// How it decides what to fetch: on personal accounts a page's lastModifiedDateTime often stays
// at its creation date, but a section's lastModifiedDateTime does move when anything in it changes.
// So a section whose date moved is re-read page by page, and each page's content is compared by hash.

import {
  getMeta, setMeta, nowIso, sha, saveRefreshToken, loadRefreshToken, writePageFile, removeFile,
} from "./store.js";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

class ThrottledError extends Error {
  constructor() {
    super("Microsoft rate-limited this account (limit: 400 requests/hour); the sync paused for 15 minutes and continues at the next scheduled run.");
    this.throttled = true;
  }
}

export function createSync({ db, dataDir, graphBase, tokenUrl, clientId, clientSecret, scopes, allowedEmail, htmlToText, exclude = [], budget, log = console }) {
  let running = null; // the promise of the sync in progress
  let lastRun = null; // summary of the last finished run
  let current = null; // live progress of the run in progress

  // ---- Microsoft access ----
  let cached = { token: null, exp: 0 };

  async function accessToken() {
    if (cached.token && cached.exp > Date.now() + 60_000) return cached.token;
    const refresh = loadRefreshToken(db, clientSecret);
    if (!refresh) throw new Error("Sync is not armed yet: it needs one sign-in through Claude after this version was deployed.");
    const r = await fetch(tokenUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: scopes, grant_type: "refresh_token", refresh_token: refresh }),
    });
    const d = await r.json().catch(() => ({}));
    if (!r.ok || !d.access_token) throw new Error(`Microsoft refused the stored sign-in (${d.error || r.status}). Reconnect OneNote in Claude to re-arm.`);
    if (d.refresh_token) saveRefreshToken(db, clientSecret, d.refresh_token);
    cached = { token: d.access_token, exp: Date.now() + (Number(d.expires_in) || 3600) * 1000 };
    return cached.token;
  }

  // Called by the /token route whenever Claude signs in or refreshes: keeps our own copy fresh.
  async function captureSignIn(tokenResponse) {
    try {
      if (!tokenResponse?.refresh_token || !tokenResponse?.access_token) return;
      const r = await fetch(`${graphBase}/me?$select=mail,userPrincipalName`, { headers: { Authorization: `Bearer ${tokenResponse.access_token}` } });
      if (!r.ok) return;
      const me = await r.json();
      const emails = [me.mail, me.userPrincipalName].filter(Boolean).map((s) => s.toLowerCase());
      if (!emails.includes(allowedEmail)) return; // only ever arm for the owner
      const firstTime = !loadRefreshToken(db, clientSecret);
      saveRefreshToken(db, clientSecret, tokenResponse.refresh_token);
      if (firstTime) {
        log.log("Sync armed by sign-in; starting first full sync.");
        start("first sign-in");
      }
    } catch (e) {
      log.warn("Could not store sign-in for sync:", e.message);
    }
  }

  // GET within the shared budget (see budget.js). Waits while the rolling hour is used up.
  // On a 429 it does not retry (OneNote sends no Retry-After, and retries add to the count):
  // the run stops, the sync pauses 15 minutes, and the next run continues where this one stopped.
  async function get(pathOrUrl, { asText = false } = {}) {
    const url = pathOrUrl.startsWith("http") ? pathOrUrl : graphBase + pathOrUrl;
    for (let attempt = 0; ; attempt++) {
      for (let wait = budget.syncWaitMs(); wait > 0; wait = budget.syncWaitMs()) await sleep(Math.min(wait, 5 * 60_000));
      const token = await accessToken();
      budget.record();
      const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (r.ok) return asText ? r.text() : r.json();
      if (r.status === 429) {
        budget.throttled();
        throw new ThrottledError();
      }
      if (r.status >= 500 && attempt < 2) {
        await sleep(5000 * 2 ** attempt);
        continue;
      }
      if (r.status === 401 && attempt < 1) {
        cached = { token: null, exp: 0 };
        continue;
      }
      let msg = "";
      try {
        msg = (await r.json())?.error?.message || "";
      } catch {
        /* ignore */
      }
      const err = new Error(`Microsoft Graph ${r.status}${msg ? `: ${msg}` : ""}`);
      err.status = r.status;
      throw err;
    }
  }

  async function getAll(first) {
    const out = [];
    for (let url = first; url; ) {
      const d = await get(url);
      out.push(...d.value);
      url = d["@odata.nextLink"];
    }
    return out;
  }

  // Every page of a section. OneNote stops at 100 pages per request and doesn't always send a
  // nextLink, so keep asking with $skip until a short page comes back. Stable order by creation.
  async function sectionPageList(sectionId) {
    const PAGE = 100;
    const seen = new Map();
    for (let skip = 0; skip < 20000; skip += PAGE) {
      let url =
        `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages` +
        `?$select=id,title,createdDateTime,lastModifiedDateTime&$orderby=createdDateTime&$top=${PAGE}&$skip=${skip}`;
      let got = 0;
      let followed = false;
      const before = seen.size;
      while (url) {
        const d = await get(url);
        got += d.value.length;
        for (const p of d.value) seen.set(p.id, p);
        url = d["@odata.nextLink"];
        if (url) followed = true;
      }
      if (followed || got < PAGE) break; // Microsoft paged it for us, or this was the last batch
      if (seen.size === before) break; // Microsoft ignored $skip and repeated itself: stop, don't loop
    }
    return [...seen.values()];
  }

  // SYNC_EXCLUDE entries: "Notebook/Section", "Section", or "Notebook/*" for a whole notebook.
  const norm = (t) => String(t || "").replace(/[\u2018\u2019\u02BC`´]/g, "'").replace(/\s+/g, " ").trim().toLowerCase();
  const isExcluded = (s) =>
    exclude.some((x) => {
      const i = x.lastIndexOf("/");
      const [nb, sec] = i >= 0 ? [norm(x.slice(0, i)), norm(x.slice(i + 1))] : [null, norm(x)];
      const nbOk = !nb || nb === norm(s.notebook);
      return nbOk && (sec === "*" || sec === norm(s.name));
    });

  // An excluded section's pages are removed from the copy (they still live in OneNote).
  function dropSection(id) {
    for (const p of db.prepare("SELECT id, file_path FROM pages WHERE section_id = ?").all(id)) {
      removeFile(dataDir, p.file_path);
      db.prepare("DELETE FROM page_fts WHERE page_id = ?").run(p.id);
      db.prepare("DELETE FROM page_versions WHERE page_id = ?").run(p.id);
    }
    return db.prepare("DELETE FROM pages WHERE section_id = ?").run(id).changes;
  }

  // ---- One section ----
  const upsertSection = db.prepare(`
    INSERT INTO sections(id, name, notebook, group_path, graph_modified) VALUES(?, ?, ?, ?, ?)
    ON CONFLICT(id) DO UPDATE SET name = excluded.name, notebook = excluded.notebook,
      group_path = excluded.group_path, graph_modified = excluded.graph_modified`);

  async function syncSection(s, stats, { onlyNew = false } = {}) {
    const existing = new Map(
      db.prepare("SELECT id, hash, title, file_path, deleted_at, error FROM pages WHERE section_id = ?").all(s.id).map((p) => [p.id, p]),
    );
    let pages;
    try {
      pages = await sectionPageList(s.id);
    } catch (e) {
      if (e.throttled) throw e;
      const status = e.status === 403 ? "encrypted" : "error";
      db.prepare("UPDATE sections SET status = ?, error = ?, last_synced = ? WHERE id = ?").run(status, e.message, nowIso(), s.id);
      if (status === "error") stats.errors.push(`${s.notebook} › ${s.name}: ${e.message}`);
      return status === "encrypted";
    }

    let ok = true;
    const seen = new Set();
    let done = 0;
    if (current) Object.assign(current, { section: `${s.notebook} › ${s.name}`, section_pages: pages.length, section_done: 0, mode: onlyNew ? "new pages only" : "all pages" });
    for (const p of pages) {
      if (current) current.section_done = done++;
      seen.add(p.id);
      const old = existing.get(p.id);
      const meta = { id: p.id, title: p.title || "(untitled)", notebook: s.notebook, section: s.name };
      // Catching up on pages past the first 100 of an unchanged section: the ones already copied are current.
      if (onlyNew && old?.hash && !old.deleted_at) continue;
      let html;
      try {
        html = await get(`/me/onenote/pages/${encodeURIComponent(p.id)}/content`, { asText: true });
      } catch (e) {
        if (e.throttled) throw e;
        ok = false;
        stats.errors.push(`${s.notebook} › ${s.name} › ${meta.title}: ${e.message}`);
        if (!old)
          db.prepare(
            `INSERT INTO pages(id, section_id, title, notebook, section, created, graph_modified, first_seen, error)
             VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          ).run(p.id, s.id, meta.title, s.notebook, s.name, p.createdDateTime, p.lastModifiedDateTime, nowIso(), e.message);
        continue;
      }
      stats.fetched++;
      const text = htmlToText(html);
      const hash = sha(html);
      const now = nowIso();
      if (old && old.hash === hash && old.title === meta.title && !old.deleted_at) {
        db.prepare("UPDATE pages SET fetched_at = ?, notebook = ?, section = ?, section_id = ?, error = NULL WHERE id = ?").run(now, s.notebook, s.name, s.id, p.id);
        continue;
      }
      if (old?.hash && old.hash !== hash) {
        // keep the previous text as a version, so edits and accidental deletions are recoverable
        db.prepare("INSERT INTO page_versions(page_id, saved_at, title, hash, html, text) SELECT id, ?, title, hash, html, text FROM pages WHERE id = ?").run(now, p.id);
      }
      removeFile(dataDir, old?.file_path);
      const filePath = writePageFile(dataDir, meta, html);
      db.prepare(
        `INSERT INTO pages(id, section_id, title, notebook, section, created, graph_modified, first_seen, changed_at, fetched_at, hash, html, text, file_path)
         VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(id) DO UPDATE SET section_id = excluded.section_id, title = excluded.title, notebook = excluded.notebook,
           section = excluded.section, graph_modified = excluded.graph_modified, changed_at = excluded.changed_at,
           fetched_at = excluded.fetched_at, hash = excluded.hash, html = excluded.html, text = excluded.text,
           file_path = excluded.file_path, deleted_at = NULL, error = NULL`,
      ).run(p.id, s.id, meta.title, s.notebook, s.name, p.createdDateTime, p.lastModifiedDateTime, now, now, now, hash, html, text, filePath);
      db.prepare("DELETE FROM page_fts WHERE page_id = ?").run(p.id);
      db.prepare("INSERT INTO page_fts(page_id, title, body) VALUES(?, ?, ?)").run(p.id, meta.title, text);
      old ? stats.updated++ : stats.added++;
    }

    // Pages that vanished from this section: kept in the backup, marked deleted (or moved, if seen elsewhere later).
    for (const [id, old] of existing) {
      if (!seen.has(id) && !old.deleted_at) {
        db.prepare("UPDATE pages SET deleted_at = ? WHERE id = ? AND section_id = ?").run(nowIso(), id, s.id);
        stats.removed++;
      }
    }
    db.prepare("UPDATE sections SET status = ?, error = NULL, last_synced = ?, page_count = ?, synced_modified = CASE WHEN ? THEN graph_modified ELSE synced_modified END WHERE id = ?")
      .run(ok ? "ok" : "partial", nowIso(), pages.length, ok ? 1 : 0, s.id);
    if (ok) setMeta(db, `paged:${s.id}`, nowIso());
    return ok;
  }

  // ---- A full pass ----
  async function run(reason, { full = false } = {}) {
    const started = nowIso();
    const stats = { reason, started, sectionsChecked: 0, sectionsSynced: 0, fetched: 0, added: 0, updated: 0, removed: 0, skipped: [], errors: [] };
    current = { started, stats };
    try {
      const sections = await getAll(
        "/me/onenote/sections?$select=id,displayName,lastModifiedDateTime&$expand=parentNotebook($select=displayName),parentSectionGroup($select=displayName)&$top=100",
      );
      stats.sectionsChecked = sections.length;
      // One notebook at a time, its sections in name order.
      sections.sort((a, b) =>
        (a.parentNotebook?.displayName || "").localeCompare(b.parentNotebook?.displayName || "") ||
        (a.displayName || "").localeCompare(b.displayName || ""));
      for (const raw of sections) {
        const s = {
          id: raw.id,
          name: raw.displayName,
          notebook: raw.parentNotebook?.displayName || "",
          group: raw.parentSectionGroup?.displayName || null,
          modified: raw.lastModifiedDateTime,
        };
        upsertSection.run(s.id, s.name, s.notebook, s.group, s.modified);
        if (isExcluded(s)) {
          db.prepare("UPDATE sections SET status = 'excluded', page_count = 0 WHERE id = ?").run(s.id);
          const dropped = dropSection(s.id);
          if (dropped) log.log(`Excluded ${s.notebook} › ${s.name}: removed ${dropped} pages from the copy.`);
          stats.skipped.push(`${s.notebook} › ${s.name} (excluded)`);
          continue;
        }
        const row = db.prepare("SELECT synced_modified, status, page_count FROM sections WHERE id = ?").get(s.id);
        // Sections copied before paging past 100 was fixed may be cut off at exactly 100: re-read once.
        const maybeCut = row?.page_count > 0 && row.page_count % 100 === 0 && !getMeta(db, `paged:${s.id}`);
        if (!full && !maybeCut && row?.synced_modified === s.modified && row?.status === "ok") continue;
        if (!full && row?.status === "encrypted" && row?.synced_modified === s.modified) continue;
        const unchanged = row?.synced_modified === s.modified && row?.status === "ok";
        const done = await syncSection(s, stats, { onlyNew: !full && maybeCut && unchanged });
        if (db.prepare("SELECT status FROM sections WHERE id = ?").get(s.id)?.status === "encrypted") {
          db.prepare("UPDATE sections SET synced_modified = graph_modified WHERE id = ?").run(s.id);
          stats.skipped.push(`${s.notebook} › ${s.name} (password-protected)`);
        } else if (done) stats.sectionsSynced++;
      }
      // sections that no longer exist in OneNote
      const live = new Set(sections.map((s) => s.id));
      for (const { id } of db.prepare("SELECT id FROM sections WHERE status != 'gone'").all()) {
        if (!live.has(id)) {
          db.prepare("UPDATE sections SET status = 'gone' WHERE id = ?").run(id);
          db.prepare("UPDATE pages SET deleted_at = COALESCE(deleted_at, ?) WHERE section_id = ?").run(nowIso(), id);
        }
      }
      setMeta(db, "last_success", nowIso());
    } catch (e) {
      if (e.throttled) stats.throttled = true;
      stats.errors.unshift(e.message);
    }
    stats.finished = nowIso();
    stats.errors = stats.errors.slice(0, 20);
    lastRun = stats;
    current = null;
    setMeta(db, "last_run", JSON.stringify(stats));
    log.log(`Sync (${reason}) done: ${stats.added} added, ${stats.updated} updated, ${stats.removed} removed, ${stats.errors.length} errors.`);
    return stats;
  }

  function start(reason, opts) {
    if (!running) running = run(reason, opts).finally(() => (running = null));
    return running;
  }

  function status() {
    const counts = db
      .prepare(
        `SELECT COUNT(*) AS pages, SUM(deleted_at IS NOT NULL) AS deleted, SUM(error IS NOT NULL AND hash IS NULL) AS failing,
                (SELECT COUNT(*) FROM page_versions) AS versions FROM pages`,
      )
      .get();
    const sections = db.prepare("SELECT status, COUNT(*) AS n FROM sections GROUP BY status").all();
    let last = lastRun;
    if (!last) {
      try {
        last = JSON.parse(getMeta(db, "last_run") || "null");
      } catch {
        last = null;
      }
    }
    return {
      armed: Boolean(loadRefreshToken(db, clientSecret)),
      microsoft_usage: budget.usage(),
      running: Boolean(running),
      ...(current ? { progress: { section: current.section, section_pages: current.section_pages, section_done: current.section_done, mode: current.mode, this_run: { fetched: current.stats.fetched, added: current.stats.added, updated: current.stats.updated, sections_synced: current.stats.sectionsSynced } } } : {}),
      last_success: getMeta(db, "last_success"),
      pages_in_copy: counts.pages - (counts.deleted || 0),
      pages_deleted_in_onenote_but_kept: counts.deleted || 0,
      pages_failing: counts.failing || 0,
      earlier_versions_kept: counts.versions,
      sections: Object.fromEntries(sections.map((r) => [r.status || "pending", r.n])),
      excluding: exclude,
      last_run: last,
    };
  }

  return { start, status, captureSignIn, isRunning: () => Boolean(running) };
}
