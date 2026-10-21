// OneNote connector for Claude (read-only).
// Exposes your personal OneNote notebooks to Claude through the MCP protocol,
// and signs you in with Microsoft. It also keeps a full local copy of your notes
// (in DATA_DIR) for full-text search and as a backup. It never writes to OneNote.

import express from "express";
import { openStore, searchIndex } from "./store.js";
import { createSync } from "./sync.js";
import { createBudget } from "./budget.js";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { convert } from "html-to-text";
import { z } from "zod";

// ---------- Settings (come from environment variables, never from code) ----------
const env = process.env;
const PUBLIC_URL = (env.PUBLIC_URL || "").replace(/\/+$/, "");
const MS_CLIENT_ID = env.MS_CLIENT_ID || "";
const MS_CLIENT_SECRET = env.MS_CLIENT_SECRET || "";
const ALLOWED_EMAIL = (env.ALLOWED_EMAIL || "").trim().toLowerCase();
const MS_TENANT = env.MS_TENANT || "consumers"; // "consumers" = personal Microsoft accounts
const PORT = Number(env.PORT || 3000);
const GRAPH_BASE = (env.GRAPH_BASE || "https://graph.microsoft.com/v1.0").replace(/\/+$/, "");
const LOGIN_BASE = (env.MS_LOGIN_BASE || "https://login.microsoftonline.com").replace(/\/+$/, "");
const SCOPES = env.SCOPES || "offline_access User.Read Notes.Read";
const ALLOWED_REDIRECTS = [
  "https://claude.ai/api/mcp/auth_callback",
  "https://claude.com/api/mcp/auth_callback",
  ...(env.EXTRA_REDIRECTS ? env.EXTRA_REDIRECTS.split(",").map((s) => s.trim()).filter(Boolean) : []),
];

const missing = ["PUBLIC_URL", "MS_CLIENT_ID", "MS_CLIENT_SECRET", "ALLOWED_EMAIL"].filter((k) => !env[k]);
if (missing.length) {
  console.error(`Missing required settings: ${missing.join(", ")}`);
  process.exit(1);
}

const AUTHORIZE_URL = `${LOGIN_BASE}/${MS_TENANT}/oauth2/v2.0/authorize`;
const TOKEN_URL = `${LOGIN_BASE}/${MS_TENANT}/oauth2/v2.0/token`;

// Local copy + sync settings
const DATA_DIR = env.DATA_DIR || "/data";
const SYNC_INTERVAL_MIN = Math.max(15, Number(env.SYNC_INTERVAL_MIN || 60));
const FULL_RESYNC_DAYS = Math.max(1, Number(env.FULL_RESYNC_DAYS || 30)); // re-read everything monthly, in case a section date didn't move
// Microsoft allows this app 400 OneNote requests per hour for your account; the sync uses at most
// SYNC_HOURLY_BUDGET of them and leaves the rest for Claude's live reads (see budget.js).
const TESTING = env.BUDGET_TEST_MODE === "1"; // only for the automated tests against a fake Microsoft
const SYNC_HOURLY_BUDGET = TESTING ? Number(env.SYNC_HOURLY_BUDGET) : Math.min(380, Math.max(50, Number(env.SYNC_HOURLY_BUDGET || 300)));
const SYNC_PER_MINUTE = TESTING ? Number(env.SYNC_PER_MINUTE) : Math.min(100, Math.max(1, Number(env.SYNC_PER_MINUTE || 30)));
const THROTTLE_PAUSE_MS = TESTING ? Number(env.THROTTLE_PAUSE_MS) : 15 * 60_000;
const SYNC_EXCLUDE =(env.SYNC_EXCLUDE || "").split(",").map((s) => s.trim()).filter(Boolean); // e.g. "Sanath Kumar's Notebook/Archive"

// ---------- Small helpers ----------
const noStore = (res) => res.set("Cache-Control", "no-store");

async function graph(token, pathOrUrl) {
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH_BASE + pathOrUrl;
  budget.record(); // counts toward Microsoft's 400/hour, so the sync yields to live reads
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: "application/json, text/html" } });
  if (!r.ok) {
    let detail = "";
    try {
      const j = await r.json();
      detail = j?.error?.message || JSON.stringify(j).slice(0, 300);
    } catch {
      /* ignore */
    }
    const hint =
      r.status === 429
        ? " (Microsoft allows 400 OneNote requests per hour for this account and that window is used up; it frees up within the hour. search_pages and get_page fall back to the VPS copy.)"
        : r.status === 404
          ? " (Not found: the id may be wrong or the item may have been deleted.)"
          : "";
    throw new Error(`Microsoft Graph error ${r.status}: ${detail}${hint}`);
  }
  const type = r.headers.get("content-type") || "";
  return type.includes("json") ? r.json() : r.text();
}

// Builds a query string. Microsoft's own "$"-options keep their punctuation (only spaces are escaped);
// everything else, including anything a person typed, is fully escaped.
const q = (obj) =>
  Object.entries(obj)
    .filter(([, v]) => v !== undefined && v !== "")
    .map(([k, v]) => `${k}=${k.startsWith("$") ? String(v).replace(/ /g, "%20") : encodeURIComponent(v)}`)
    .join("&");

const clampLimit = (n, def = 20, max = 100) => Math.max(1, Math.min(max, Number.isFinite(n) ? n : def));

const ok = (text) => ({ content: [{ type: "text", text }] });
const asJson = (data) => ok(JSON.stringify(data, null, 1));
const fail = (e) => ({ isError: true, content: [{ type: "text", text: String(e?.message || e) }] });

function htmlToText(html) {
  const title = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(html)?.[1]?.trim();
  const body = convert(html, {
    wordwrap: false,
    selectors: [
      { selector: "img", format: "skip" },
      { selector: "a", options: { ignoreHref: true } },
      ...["h1", "h2", "h3", "h4", "h5", "h6"].map((selector) => ({ selector, options: { uppercase: false } })),
    ],
  });
  return (title ? `# ${title}\n\n` : "") + body.trim();
}

const pageFields = "id,title,createdDateTime,lastModifiedDateTime";
const slim = (p) => ({
  id: p.id,
  title: p.title || "(untitled)",
  created: p.createdDateTime,
  modified: p.lastModifiedDateTime,
  ...(p.parentSection?.displayName ? { section: p.parentSection.displayName } : {}),
});

// Every section in every notebook, with its notebook's name. Follows Microsoft's paging links.
async function allSections(token) {
  let url = `/me/onenote/sections?${q({
    $select: "id,displayName,lastModifiedDateTime",
    $expand: "parentNotebook($select=id,displayName)",
    $top: 100,
  })}`;
  const out = [];
  while (url) {
    const d = await graph(token, url);
    for (const s of d.value)
      out.push({ id: s.id, name: s.displayName, notebook: s.parentNotebook?.displayName, modified: s.lastModifiedDateTime });
    url = d["@odata.nextLink"];
  }
  return out;
}

// Pages of one section (Microsoft refuses account-wide page queries when there are many sections).
const sectionPages = (token, sectionId, top) =>
  graph(
    token,
    `/me/onenote/sections/${encodeURIComponent(sectionId)}/pages?${q({
      $select: pageFields,
      $orderby: "lastModifiedDateTime desc",
      $top: top,
    })}`,
  ).then((d) => d.value);

// Runs fn over items a few at a time, so Microsoft doesn't rate-limit us.
async function mapLimited(items, n, fn) {
  const results = new Array(items.length);
  let i = 0;
  await Promise.all(
    Array.from({ length: Math.min(n, items.length) }, async () => {
      while (i < items.length) {
        const k = i++;
        results[k] = await fn(items[k]);
      }
    }),
  );
  return results;
}

// ---------- The tools Claude can use ----------
function buildServer(token) {
  const server = new McpServer(
    { name: "onenote-readonly", version: "1.0.0" },
    {
      instructions:
        "Read-only access to the user's personal OneNote. Start with list_notebooks or recent_pages, " +
        "then use list_sections / list_pages to find a page id, and get_page to read it. " +
        "Text is extracted from each page; images, drawings and handwriting are not included.",
    },
  );
  const readOnly = { readOnlyHint: true, openWorldHint: false };

  server.registerTool(
    "list_notebooks",
    { title: "List notebooks", description: "List all OneNote notebooks (id, name, last modified).", inputSchema: {}, annotations: readOnly },
    async () => {
      try {
        const d = await graph(token, `/me/onenote/notebooks?${q({ $select: "id,displayName,lastModifiedDateTime" })}`);
        return asJson(d.value.map((n) => ({ id: n.id, name: n.displayName, modified: n.lastModifiedDateTime })));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "list_sections",
    {
      title: "List sections",
      description:
        "List sections. With no arguments: every section in every notebook. With notebook_id or section_group_id: the sections and section groups directly inside it.",
      inputSchema: {
        notebook_id: z.string().optional().describe("Notebook id from list_notebooks"),
        section_group_id: z.string().optional().describe("Section group id from a previous list_sections call"),
      },
      annotations: readOnly,
    },
    async ({ notebook_id, section_group_id }) => {
      try {
        const sel = { $select: "id,displayName,lastModifiedDateTime" };
        const base = section_group_id
          ? `/me/onenote/sectionGroups/${encodeURIComponent(section_group_id)}`
          : notebook_id
            ? `/me/onenote/notebooks/${encodeURIComponent(notebook_id)}`
            : null;
        if (!base) return asJson(await allSections(token));
        const [secs, groups] = await Promise.all([
          graph(token, `${base}/sections?${q(sel)}`),
          graph(token, `${base}/sectionGroups?${q(sel)}`),
        ]);
        return asJson({
          sections: secs.value.map((s) => ({ id: s.id, name: s.displayName, modified: s.lastModifiedDateTime })),
          section_groups: groups.value.map((g) => ({ id: g.id, name: g.displayName, modified: g.lastModifiedDateTime })),
        });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "list_pages",
    {
      title: "List pages in a section",
      description: "List pages in a section, newest-modified first (id, title, dates).",
      inputSchema: {
        section_id: z.string().describe("Section id from list_sections"),
        limit: z.number().int().min(1).max(100).optional().describe("How many pages (default 30)"),
      },
      annotations: readOnly,
    },
    async ({ section_id, limit }) => {
      try {
        const d = await graph(
          token,
          `/me/onenote/sections/${encodeURIComponent(section_id)}/pages?${q({
            $select: pageFields,
            $orderby: "lastModifiedDateTime desc",
            $top: clampLimit(limit, 30),
          })}`,
        );
        return asJson(d.value.map(slim));
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "recent_pages",
    {
      title: "Recently modified pages",
      description:
        "Pages that were added or changed recently, across all notebooks, with notebook and section. Once the local copy exists, " +
        "'changed_at' is when the sync actually saw the page's content change (reliable even though Microsoft's own page dates are not); " +
        "pass since to get everything changed after a date. Before the first sync it falls back to the most recently edited sections.",
      inputSchema: {
        limit: z.number().int().min(1).max(100).optional().describe("How many pages (default 20)"),
        since: z.string().optional().describe("Only pages changed after this date (YYYY-MM-DD or ISO time); local copy only"),
        sections: z.number().int().min(1).max(30).optional().describe("Fallback mode: how many recently edited sections to scan (default 8)"),
      },
      annotations: readOnly,
    },
    async ({ limit, since, sections }) => {
      try {
        if (hasCopy()) {
          const rows = db
            .prepare(
              `SELECT id, title, notebook, section, section_id, created, changed_at, deleted_at FROM pages
               WHERE changed_at IS NOT NULL AND changed_at > ? ORDER BY changed_at DESC LIMIT ?`,
            )
            .all(since || "", clampLimit(limit, 20));
          return asJson({
            source: "local copy",
            note: "The first full sync marks every page as changed on the day it ran; later dates are real edits.",
            last_sync: sync.status().last_success,
            pages: rows.map((r) => ({ ...r })),
          });
        }
        const recent = (await allSections(token))
          .sort((a, b) => String(b.modified).localeCompare(String(a.modified)))
          .slice(0, clampLimit(sections, 8, 30));
        const perSection = await mapLimited(recent, 4, async (s) =>
          (await sectionPages(token, s.id, 10)).map((p) => ({
            ...slim(p),
            section: s.name,
            section_id: s.id,
            notebook: s.notebook,
            section_modified: s.modified,
          })),
        );
        const pages = perSection
          .flat()
          .sort((a, b) => String(b.modified).localeCompare(String(a.modified)))
          .slice(0, clampLimit(limit, 20));
        return asJson({ recently_edited_sections: recent, pages });
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "search_pages",
    {
      title: "Search pages",
      description:
        "Full-text search across every notebook: page titles AND page text, in any script (English, Kannada, transliterated Sanskrit). " +
        "All words must appear; any 3+ letter fragment matches (\"practis\" finds practice, practising). Results carry notebook, section " +
        "and a snippet with the match in «». Searches the local copy kept by the sync; before the first sync finishes it falls back " +
        "to matching page titles only. Password-protected sections are never included.",
      inputSchema: {
        query: z.string().min(1).describe("Words to search for"),
        limit: z.number().int().min(1).max(100).optional().describe("How many results (default 20)"),
        include_deleted: z.boolean().optional().describe("Also return pages since deleted in OneNote but kept in the backup"),
      },
      annotations: readOnly,
    },
    async ({ query, limit, include_deleted }) => {
      try {
        if (hasCopy()) {
          const hits = searchIndex(db, query, 200)
            .filter((r) => include_deleted || !r.deleted_at)
            .slice(0, clampLimit(limit, 20));
          return hits.length
            ? asJson({
                source: "local copy (full text)",
                last_sync: sync.status().last_success,
                ...(sync.status().last_success ? {} : { note: "The first sync is still running, so some notebooks may be missing." }),
                results: hits,
              })
            : ok(`No pages contain all of: ${query}. Try fewer or shorter words.`);
        }
        const needle = query.toLowerCase();
        const secs = await allSections(token);
        const perSection = await mapLimited(secs, 4, async (s) => {
          try {
            return (await sectionPages(token, s.id, 100))
              .filter((p) => (p.title || "").toLowerCase().includes(needle))
              .map((p) => ({ ...slim(p), section: s.name, section_id: s.id, notebook: s.notebook }));
          } catch {
            return []; // encrypted (password-protected) sections can't be read; skip them
          }
        });
        const hits = perSection.flat().slice(0, clampLimit(limit, 20));
        return hits.length ? asJson(hits) : ok("No page titles matched. Try fewer words, or browse with list_sections / list_pages.");
      } catch (e) {
        return fail(e);
      }
    },
  );

  server.registerTool(
    "get_page",
    {
      title: "Read a page",
      description:
        "Read the text of one OneNote page, headed by its notebook and section (read every page in that context: the same section " +
        "names mean different things in different notebooks). Reads live from OneNote; if that fails (or the page was deleted) it " +
        "returns the local backup copy. Pass version_saved_at (from page_history) to read an earlier version. Long pages come in " +
        "chunks: use offset to continue.",
      inputSchema: {
        page_id: z.string().describe("Page id from list_pages, recent_pages or search_pages"),
        offset: z.number().int().min(0).optional().describe("Character position to start from (default 0)"),
        max_chars: z.number().int().min(1000).max(100000).optional().describe("Chunk size (default 30000)"),
        version_saved_at: z.string().optional().describe("Read an earlier version kept by the backup (a saved_at from page_history)"),
      },
      annotations: readOnly,
    },
    async ({ page_id, offset = 0, max_chars = 30000, version_saved_at }) => {
      const local = db.prepare("SELECT title, notebook, section, text, fetched_at, deleted_at FROM pages WHERE id = ?").get(page_id);
      const where = local ? `[Notebook: ${local.notebook} › Section: ${local.section}]\n\n` : "";
      const chunked = (text, note = "") => {
        const chunk = text.slice(offset, offset + max_chars);
        const end = offset + chunk.length;
        const more = end < text.length ? `\n\n[Chunk ends at character ${end} of ${text.length}. Call get_page again with offset=${end} to continue.]` : "";
        return ok(note + where + (chunk || "(This page has no readable text.)") + more);
      };
      if (version_saved_at) {
        const v = db.prepare("SELECT text FROM page_versions WHERE page_id = ? AND saved_at = ?").get(page_id, version_saved_at);
        return v ? chunked(v.text || "", `[Earlier version, replaced on ${version_saved_at}]\n`) : fail("No version saved at that time; see page_history.");
      }
      try {
        const html = await graph(token, `/me/onenote/pages/${encodeURIComponent(page_id)}/content`);
        return chunked(htmlToText(typeof html === "string" ? html : JSON.stringify(html)));
      } catch (e) {
        if (local?.text != null)
          return chunked(
            local.text,
            `[From the local backup, saved ${local.fetched_at}${local.deleted_at ? `; deleted in OneNote around ${local.deleted_at}` : ""}. Live read failed: ${e.message}]\n`,
          );
        return fail(e);
      }
    },
  );

  server.registerTool(
    "page_history",
    {
      title: "Earlier versions of a page",
      description: "List the earlier versions of a page that the backup kept each time its content changed (newest first).",
      inputSchema: { page_id: z.string().describe("Page id") },
      annotations: readOnly,
    },
    async ({ page_id }) => {
      const rows = db
        .prepare("SELECT saved_at, title, length(text) AS chars FROM page_versions WHERE page_id = ? ORDER BY saved_at DESC")
        .all(page_id)
        .map((r) => ({ ...r }));
      return rows.length ? asJson(rows) : ok("No earlier versions kept for this page (it hasn't changed since the backup started).");
    },
  );

  server.registerTool(
    "sync_status",
    {
      title: "Local copy status",
      description:
        "State of the local copy/backup: whether sync is armed, last successful sync, page counts, kept versions, sections by status " +
        "(ok, encrypted, excluded, error) and the last run's errors.",
      inputSchema: {},
      annotations: readOnly,
    },
    async () => asJson(sync.status()),
  );

  server.registerTool(
    "sync_now",
    {
      title: "Refresh the local copy now",
      description:
        "Start a sync now (it normally runs every hour) and wait up to 50 seconds for it. Only changed sections are re-read unless full is true. " +
        "Reads OneNote; writes only to the local copy on this server.",
      inputSchema: { full: z.boolean().optional().describe("Re-read every section, not just changed ones") },
      annotations: { readOnlyHint: true, openWorldHint: false },
    },
    async ({ full }) => {
      try {
        const p = sync.start(full ? "manual full" : "manual", { full: Boolean(full) });
        const done = await Promise.race([p, new Promise((r) => setTimeout(() => r(null), 50_000))]);
        return asJson(done ? { finished: true, ...done } : { finished: false, note: "Still running in the background; check sync_status later." });
      } catch (e) {
        return fail(e);
      }
    },
  );

  return server;
}

// ---------- Local copy ----------
const db = openStore(DATA_DIR);
const budget = createBudget(db, { syncHourly: SYNC_HOURLY_BUDGET, syncPerMinute: SYNC_PER_MINUTE, pauseMs: THROTTLE_PAUSE_MS });
const hasCopy = () => Boolean(db.prepare("SELECT 1 FROM pages WHERE hash IS NOT NULL LIMIT 1").get());
const sync = createSync({
  db,
  dataDir: DATA_DIR,
  graphBase: GRAPH_BASE,
  tokenUrl: TOKEN_URL,
  clientId: MS_CLIENT_ID,
  clientSecret: MS_CLIENT_SECRET,
  scopes: SCOPES,
  allowedEmail: ALLOWED_EMAIL,
  htmlToText,
  exclude: SYNC_EXCLUDE,
  budget,
});

// Normal runs re-read only sections whose date moved (or that didn't finish last time), so the
// first copy simply continues run after run. A full re-read happens once every FULL_RESYNC_DAYS,
// counted from when the first copy completed.
function scheduledSync() {
  const lastFull = db.prepare("SELECT value FROM meta WHERE key = 'last_full'").get()?.value;
  const full = Boolean(lastFull) && Date.now() - Date.parse(lastFull) > FULL_RESYNC_DAYS * 86_400_000;
  sync
    .start(full ? "scheduled full" : "scheduled", { full })
    .then((s) => {
      const complete = !s.throttled && !db.prepare("SELECT 1 FROM sections WHERE status IN ('error','partial') OR (status IS NULL)").get();
      if (complete && (full || !lastFull)) db.prepare("INSERT OR REPLACE INTO meta(key, value) VALUES('last_full', ?)").run(new Date().toISOString());
    })
    .catch((e) => console.warn("Scheduled sync failed:", e.message));
}
setTimeout(scheduledSync, 20_000); // shortly after start (no-op until armed)
setInterval(scheduledSync, SYNC_INTERVAL_MIN * 60_000);

// ---------- Web server ----------
const app = express();
app.set("trust proxy", true);
app.disable("x-powered-by");
app.use(express.json({ limit: "1mb" }));
app.use(express.urlencoded({ extended: false, limit: "100kb" }));

app.get("/health", (_req, res) => res.json({ ok: true }));
app.get("/", (_req, res) => res.type("text/plain").send("OneNote connector for Claude is running."));

// --- Discovery documents Claude reads to learn how to sign in ---
const protectedResource = (_req, res) =>
  res.json({
    resource: `${PUBLIC_URL}/mcp`,
    authorization_servers: [PUBLIC_URL],
    bearer_methods_supported: ["header"],
    scopes_supported: SCOPES.split(" "),
  });
app.get("/.well-known/oauth-protected-resource", protectedResource);
app.get("/.well-known/oauth-protected-resource/mcp", protectedResource);

app.get("/.well-known/oauth-authorization-server", (_req, res) =>
  res.json({
    issuer: PUBLIC_URL,
    authorization_endpoint: `${PUBLIC_URL}/authorize`,
    token_endpoint: `${PUBLIC_URL}/token`,
    registration_endpoint: `${PUBLIC_URL}/register`,
    response_types_supported: ["code"],
    grant_types_supported: ["authorization_code", "refresh_token"],
    token_endpoint_auth_methods_supported: ["none"],
    code_challenge_methods_supported: ["S256"],
    scopes_supported: SCOPES.split(" "),
  }),
);

// Claude registers itself here. We hand back our Microsoft app's id (the secret never leaves this server).
app.post("/register", (req, res) => {
  const uris = Array.isArray(req.body?.redirect_uris) ? req.body.redirect_uris : [];
  if (!uris.length || !uris.every((u) => ALLOWED_REDIRECTS.includes(u))) {
    return res.status(400).json({ error: "invalid_redirect_uri", error_description: "Only Claude's callback addresses are allowed." });
  }
  noStore(res)
    .status(201)
    .json({
      client_id: MS_CLIENT_ID,
      client_id_issued_at: Math.floor(Date.now() / 1000),
      client_name: req.body?.client_name || "Claude",
      redirect_uris: uris,
      grant_types: ["authorization_code", "refresh_token"],
      response_types: ["code"],
      token_endpoint_auth_method: "none",
    });
});

// Send the person to Microsoft's sign-in page.
app.get("/authorize", (req, res) => {
  const { response_type, redirect_uri, state, code_challenge, code_challenge_method } = req.query;
  if (response_type !== "code") return res.status(400).send("Unsupported response_type.");
  if (!ALLOWED_REDIRECTS.includes(String(redirect_uri))) return res.status(400).send("Unknown redirect address.");
  if (!code_challenge || code_challenge_method !== "S256") return res.status(400).send("PKCE (S256) is required.");
  const target = new URL(AUTHORIZE_URL);
  target.search = q({
    client_id: MS_CLIENT_ID,
    response_type: "code",
    redirect_uri: String(redirect_uri),
    scope: SCOPES,
    state: state ? String(state) : "",
    code_challenge: String(code_challenge),
    code_challenge_method: "S256",
    prompt: "select_account",
  });
  noStore(res).redirect(302, target.toString());
});

// Swap the sign-in code (or a refresh token) for an access token, adding our secret.
app.post("/token", async (req, res) => {
  const b = req.body || {};
  const form = new URLSearchParams({ client_id: MS_CLIENT_ID, client_secret: MS_CLIENT_SECRET, scope: SCOPES });
  if (b.grant_type === "authorization_code") {
    if (!ALLOWED_REDIRECTS.includes(String(b.redirect_uri))) return res.status(400).json({ error: "invalid_request" });
    form.set("grant_type", "authorization_code");
    form.set("code", String(b.code || ""));
    form.set("redirect_uri", String(b.redirect_uri));
    if (b.code_verifier) form.set("code_verifier", String(b.code_verifier));
  } else if (b.grant_type === "refresh_token") {
    form.set("grant_type", "refresh_token");
    form.set("refresh_token", String(b.refresh_token || ""));
  } else {
    return res.status(400).json({ error: "unsupported_grant_type" });
  }
  try {
    const r = await fetch(TOKEN_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: form,
    });
    const data = await r.json().catch(() => ({ error: "server_error" }));
    if (!r.ok) console.warn(`Token request failed: ${r.status} ${data.error || ""}`);
    else sync.captureSignIn(data); // keeps the background sync's own sign-in fresh (owner only)
    noStore(res).status(r.status).json(data);
  } catch (e) {
    console.error("Token request error:", e.message);
    res.status(502).json({ error: "server_error", error_description: "Could not reach Microsoft." });
  }
});

// --- Checking who is calling /mcp ---
const verified = new Map(); // token -> { exp }
function challenge(res, status, error) {
  res
    .status(status)
    .set("WWW-Authenticate", `Bearer resource_metadata="${PUBLIC_URL}/.well-known/oauth-protected-resource"${error ? `, error="${error}"` : ""}`)
    .json({ error: error || "unauthorized" });
}

async function requireOwner(req, res, next) {
  const m = /^Bearer (.+)$/i.exec(req.get("authorization") || "");
  if (!m) return challenge(res, 401);
  const token = m[1];
  const hit = verified.get(token);
  if (hit && hit.exp > Date.now()) {
    req.msToken = token;
    return next();
  }
  try {
    const r = await fetch(`${GRAPH_BASE}/me?${q({ $select: "mail,userPrincipalName" })}`, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) return challenge(res, 401, "invalid_token");
    const me = await r.json();
    const emails = [me.mail, me.userPrincipalName].filter(Boolean).map((s) => s.toLowerCase());
    if (!emails.includes(ALLOWED_EMAIL)) {
      console.warn(`Rejected sign-in from ${emails.join(" / ") || "unknown account"}`);
      return res.status(403).json({ error: "forbidden", error_description: "This connector is private." });
    }
    if (verified.size > 200) verified.clear();
    verified.set(token, { exp: Date.now() + 5 * 60 * 1000 });
    req.msToken = token;
    next();
  } catch (e) {
    console.error("Verification error:", e.message);
    res.status(502).json({ error: "server_error" });
  }
}

app.post("/mcp", requireOwner, async (req, res) => {
  const server = buildServer(req.msToken);
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  res.on("close", () => {
    transport.close();
    server.close();
  });
  try {
    await server.connect(transport);
    await transport.handleRequest(req, res, req.body);
  } catch (e) {
    console.error("MCP error:", e.message);
    if (!res.headersSent) res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal error" }, id: null });
  }
});
const methodNotAllowed = (_req, res) =>
  res.status(405).set("Allow", "POST").json({ jsonrpc: "2.0", error: { code: -32000, message: "Method not allowed." }, id: null });
app.get("/mcp", methodNotAllowed);
app.delete("/mcp", methodNotAllowed);

app.listen(PORT, () => console.log(`OneNote connector listening on port ${PORT}`));
