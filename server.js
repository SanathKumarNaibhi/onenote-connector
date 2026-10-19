// OneNote connector for Claude (read-only).
// Exposes your personal OneNote notebooks to Claude through the MCP protocol,
// and signs you in with Microsoft. Nothing is stored on this server.

import express from "express";
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

// ---------- Small helpers ----------
const noStore = (res) => res.set("Cache-Control", "no-store");

async function graph(token, pathOrUrl) {
  const url = pathOrUrl.startsWith("http") ? pathOrUrl : GRAPH_BASE + pathOrUrl;
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
        ? " (Microsoft is rate-limiting requests; try again in a minute.)"
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
        if (!base) {
          const d = await graph(token, `/me/onenote/sections?${q({ ...sel, $top: 100 })}`);
          return asJson(d.value.map((s) => ({ id: s.id, name: s.displayName, modified: s.lastModifiedDateTime })));
        }
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
      description: "The most recently modified pages across all notebooks, with the section each belongs to.",
      inputSchema: { limit: z.number().int().min(1).max(100).optional().describe("How many pages (default 20)") },
      annotations: readOnly,
    },
    async ({ limit }) => {
      try {
        const d = await graph(
          token,
          `/me/onenote/pages?${q({
            $select: pageFields,
            $expand: "parentSection($select=id,displayName)",
            $orderby: "lastModifiedDateTime desc",
            $top: clampLimit(limit, 20),
          })}`,
        );
        return asJson(d.value.map(slim));
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
        "Search OneNote pages for a word or phrase (best effort: Microsoft's search can miss recent edits). If it returns nothing, browse with list_sections and list_pages.",
      inputSchema: {
        query: z.string().min(1).describe("Text to search for"),
        limit: z.number().int().min(1).max(100).optional().describe("How many results (default 20)"),
      },
      annotations: readOnly,
    },
    async ({ query, limit }) => {
      try {
        const d = await graph(
          token,
          `/me/onenote/pages?${q({
            search: query,
            $select: pageFields,
            $expand: "parentSection($select=id,displayName)",
            $top: clampLimit(limit, 20),
          })}`,
        );
        return d.value.length ? asJson(d.value.map(slim)) : ok("No pages matched. Try fewer words, or browse with list_sections / list_pages.");
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
        "Read the text of one OneNote page. Long pages are returned in chunks: use offset to continue where the previous chunk ended.",
      inputSchema: {
        page_id: z.string().describe("Page id from list_pages, recent_pages or search_pages"),
        offset: z.number().int().min(0).optional().describe("Character position to start from (default 0)"),
        max_chars: z.number().int().min(1000).max(100000).optional().describe("Chunk size (default 30000)"),
      },
      annotations: readOnly,
    },
    async ({ page_id, offset = 0, max_chars = 30000 }) => {
      try {
        const html = await graph(token, `/me/onenote/pages/${encodeURIComponent(page_id)}/content`);
        const text = htmlToText(typeof html === "string" ? html : JSON.stringify(html));
        const chunk = text.slice(offset, offset + max_chars);
        const end = offset + chunk.length;
        const more = end < text.length ? `\n\n[Chunk ends at character ${end} of ${text.length}. Call get_page again with offset=${end} to continue.]` : "";
        return ok((chunk || "(This page has no readable text.)") + more);
      } catch (e) {
        return fail(e);
      }
    },
  );

  return server;
}

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
