# OneNote connector for Claude

A private, read-only MCP connector for one personal Microsoft account's OneNote. It also keeps a
**full local copy** of every page on this server, for full-text search and as a backup.

## What it does

- **Live tools**: list notebooks, sections and pages; read any page (headed with its notebook and section).
- **Local copy** (in `/data`): every page's HTML and text, refreshed every hour; only sections whose
  date moved are re-read, plus a full pass once a week.
- **Full-text search** over titles and page text, in any script (English, Kannada, transliteration).
  Any 3+ letter fragment matches.
- **Backup behaviour**: pages deleted in OneNote are kept and marked deleted. Each time a page's
  content changes, the previous version is kept (`page_history`, `get_page` with `version_saved_at`).
- **Readable mirror**: `/data/pages/<Notebook>/<Section>/<Title> [id].html` can be opened in a browser.
- It never writes to OneNote (scope `Notes.Read`).

Not copied: password-protected sections (Microsoft doesn't serve them), and images or attachments
(the HTML keeps their links, but they need a Microsoft sign-in to load).

## Deploy (Coolify)

1. In the app's **Persistent Storage**, add a volume mounted at **`/data`**. Without it, the copy is
   lost on every redeploy.
2. Redeploy. No new settings are required.
3. Arming: the background sync needs one Microsoft sign-in after deploy. Claude refreshes its sign-in
   about every hour, which arms it automatically. To arm at once, disconnect and reconnect the OneNote
   connector in Claude. The first full sync then starts by itself; ask Claude for `sync_status` to
   watch it.

Optional environment variables:

| Variable | Default | Meaning |
| --- | --- | --- |
| `SYNC_INTERVAL_MIN` | `60` | Minutes between syncs (minimum 15) |
| `FULL_RESYNC_DAYS` | `7` | Days between full re-reads of every section |
| `SYNC_EXCLUDE` | (none) | Comma-separated sections to leave out of the copy, as `Notebook/Section` or `Section` |
| `DATA_DIR` | `/data` | Where the copy lives |

## Security notes

- Only `ALLOWED_EMAIL` can use the connector, and only that account's sign-in arms the sync.
- The sync's Microsoft refresh token is stored encrypted (AES-256-GCM) with a key derived from
  `MS_CLIENT_SECRET`. When you renew the secret, the next sign-in through Claude re-arms the sync.
- The copy holds everything in your notes, readable by anyone with access to the server's volume.
  Back up the `/data` folder somewhere private if you want an off-server copy.
