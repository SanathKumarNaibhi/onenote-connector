// Microsoft's documented OneNote limits for one app acting for one user (delegated):
//   120 requests per minute, 400 requests per hour, 5 concurrent requests.
//   https://learn.microsoft.com/en-us/graph/throttling-limits#onenote-service-limits
// 429 responses from OneNote carry no Retry-After header; the windows simply roll over.
//
// Every request this server makes to OneNote (Claude's live reads and the background sync)
// is recorded here. The sync only spends up to SYNC_HOURLY_BUDGET per rolling hour and waits
// when the window is full, which leaves the rest of the 400 for Claude's live reads.
// Timestamps are persisted so a restart doesn't reset the count.

import { getMeta, setMeta } from "./store.js";

export const HOURLY_LIMIT = 400;
export const MINUTE_LIMIT = 120;
const HOUR = 3_600_000;
const MINUTE = 60_000;

export function createBudget(db, { syncHourly = 300, syncPerMinute = 30, pauseMs = 15 * 60_000 } = {}) {
  let stamps = [];
  try {
    stamps = JSON.parse(getMeta(db, "graph_request_times") || "[]").filter((t) => Date.now() - t < HOUR);
  } catch {
    stamps = [];
  }
  let throttledUntil = Number(getMeta(db, "throttled_until") || 0);
  let unsaved = 0;

  const prune = () => {
    const cut = Date.now() - HOUR;
    while (stamps.length && stamps[0] < cut) stamps.shift();
  };
  const save = () => {
    setMeta(db, "graph_request_times", JSON.stringify(stamps));
    unsaved = 0;
  };

  function record() {
    stamps.push(Date.now());
    if (++unsaved >= 10) save();
  }

  // How long the sync must wait before its next request (0 = go now).
  function syncWaitMs() {
    prune();
    const now = Date.now();
    if (throttledUntil > now) return throttledUntil - now;
    let wait = 0;
    if (stamps.length >= syncHourly) wait = stamps[stamps.length - syncHourly] + HOUR - now;
    const lastMinute = stamps.filter((t) => now - t < MINUTE);
    if (lastMinute.length >= syncPerMinute) wait = Math.max(wait, lastMinute[lastMinute.length - syncPerMinute] + MINUTE - now);
    return Math.max(0, wait) + (wait > 0 ? 250 : 0);
  }

  // Microsoft said "too many": nothing more from the sync for a while.
  function throttled(ms = pauseMs) {
    throttledUntil = Date.now() + ms;
    setMeta(db, "throttled_until", String(throttledUntil));
    save();
  }

  function usage() {
    prune();
    const now = Date.now();
    return {
      requests_last_hour: stamps.length,
      requests_last_minute: stamps.filter((t) => now - t < MINUTE).length,
      microsoft_limits: `${HOURLY_LIMIT}/hour, ${MINUTE_LIMIT}/minute`,
      sync_allowance: `${syncHourly}/hour, ${syncPerMinute}/minute (the rest is kept for Claude's live reads)`,
      paused_until: throttledUntil > now ? new Date(throttledUntil).toISOString() : null,
    };
  }

  return { record, syncWaitMs, throttled, usage, save };
}
