// Playwright transport: reuse the persistent profile, lift chatsvc auth headers
// from a request Teams web makes itself, then call the endpoints via context.request.
import { chatsvcBase } from './core.js';

const TEAMS_URL = 'https://teams.microsoft.com/v2/';
const LOGIN_RE = /login\.microsoftonline|login\.live|login\.windows/;
const DROP = new Set(['content-length', 'host', 'cookie', 'connection', 'content-type', 'origin', 'referer']);

async function chromium() {
  try {
    return (await import('playwright')).chromium;
  } catch {
    throw new Error('playwright not installed: run `npm install && npx playwright install chromium`');
  }
}

// Headless Chromium advertises "HeadlessChrome", which Teams may reject.
async function launch(profileDir, headless) {
  const pw = await chromium();
  const executablePath = process.env.TEAMS_EXPORT_BROWSER || undefined;
  let userAgent;
  if (headless) {
    const probe = await pw.launchPersistentContext(profileDir, { headless: true, executablePath });
    userAgent = (await probe.pages()[0]?.evaluate(() => navigator.userAgent)) ||
      (await (await probe.newPage()).evaluate(() => navigator.userAgent));
    await probe.close();
    userAgent = userAgent.replace(/HeadlessChrome/g, 'Chrome');
  }
  return pw.launchPersistentContext(profileDir, { headless, userAgent, executablePath, viewport: { width: 1280, height: 800 } });
}

function watchChatsvc(page) {
  let resolve;
  const found = new Promise((r) => (resolve = r));
  page.on('request', async (req) => {
    const base = chatsvcBase(req.url());
    if (!base) return;
    const h = await req.allHeaders().catch(() => req.headers());
    if (!h.authentication && !h.authorization && !h['x-skypetoken']) return;
    const headers = {};
    for (const [k, v] of Object.entries(h)) if (!k.startsWith(':') && !DROP.has(k)) headers[k] = v;
    resolve({ base, headers });
  });
  return found;
}

const timeout = (ms, v) => new Promise((r) => setTimeout(() => r(v), ms));

export async function login({ profileDir, waitMs = 10 * 60_000 }) {
  const ctx = await launch(profileDir, false);
  const page = ctx.pages()[0] || (await ctx.newPage());
  const found = watchChatsvc(page);
  const closed = new Promise((r) => ctx.on('close', () => r('closed')));
  await page.goto(TEAMS_URL).catch(() => {});
  const r = await Promise.race([found, closed, timeout(waitMs, 'timeout')]);
  if (typeof r === 'object') {
    await timeout(5000); // let Teams persist its tokens to the profile
    await ctx.close().catch(() => {});
    return { ok: true, base: r.base };
  }
  await ctx.close().catch(() => {});
  return { ok: false, reason: r };
}

export async function openSession({ profileDir, headless = true, waitMs = 90_000 }) {
  const ctx = await launch(profileDir, headless);
  const page = ctx.pages()[0] || (await ctx.newPage());
  const found = watchChatsvc(page);
  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' }).catch(() => {});
  const stuckOnLogin = (async () => {
    for (let i = 0; i < waitMs / 1000; i++) {
      await timeout(1000);
      if (i >= 10 && LOGIN_RE.test(page.url())) return null; // no interactive login in export
    }
    return null;
  })();
  const r = await Promise.race([found, stuckOnLogin, timeout(waitMs, null)]);
  if (!r) {
    const url = page.url();
    await ctx.close().catch(() => {});
    const err = new Error(LOGIN_RE.test(url)
      ? 'not logged in: run `teams-export login`'
      : 'no chatsvc request seen (Teams web changed or slow load); try --headful or --wait');
    err.url = url;
    throw err;
  }
  const base = process.env.TEAMS_CHATSVC_BASE || r.base;
  const fetchJson = async (url) => {
    for (let attempt = 0; ; attempt++) {
      const resp = await ctx.request.get(url, { headers: r.headers, failOnStatusCode: false, timeout: 60_000 });
      const status = resp.status();
      if ((status === 429 || status >= 500) && attempt < 3) {
        const ra = Number(resp.headers()['retry-after']) || 2 ** attempt * 2;
        await timeout(Math.min(ra, 60) * 1000);
        continue;
      }
      const text = await resp.text();
      if (status >= 200 && status < 300) {
        try { return { status, body: JSON.parse(text) }; } catch { return { status, body: text }; }
      }
      return { status, body: text };
    }
  };
  return { base, fetchJson, close: () => ctx.close().catch(() => {}) };
}
