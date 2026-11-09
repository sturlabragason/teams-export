// Transport-agnostic export logic. No Playwright here so it can be tested offline.
import fs from 'node:fs';
import path from 'node:path';

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', '#39': "'" };

export function stripHtml(html) {
  if (!html) return '';
  return String(html)
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|blockquote|h[1-6]|tr)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&(#x[0-9a-f]+|#\d+|[a-z]+\d*|#39);/gi, (m, e) => {
      const k = e.toLowerCase();
      if (k.startsWith('#x')) return String.fromCodePoint(parseInt(k.slice(2), 16));
      if (k.startsWith('#') && k !== '#39') return String.fromCodePoint(parseInt(k.slice(1), 10));
      return ENTITIES[k] ?? m;
    })
    .replace(/[ \t\u00a0]+/g, ' ')
    .replace(/ *\n */g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// Message ids are numeric strings (epoch ms); compare without float precision loss.
export function cmpId(a, b) {
  a = String(a); b = String(b);
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) return a.length - b.length || (a < b ? -1 : a > b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

// Base = everything before "/v1/users/ME/", e.g.
//   https://teams.microsoft.com/api/chatsvc/emea   or   https://amer.ng.msg.teams.microsoft.com
export function chatsvcBase(url) {
  const m = /^(https:\/\/[^?#]+?)\/v1\/users\/ME\//i.exec(url || '');
  if (!m) return null;
  return /chatsvc|\.msg\.|ng\.msg/i.test(m[1]) ? m[1] : null;
}

// Pagination links may point at another host/region alias; keep our authenticated base.
export function rebase(link, base) {
  if (!link) return null;
  const i = link.indexOf('/v1/users/ME/');
  return i >= 0 ? base + link.slice(i) : new URL(link, base + '/').toString();
}

export const conversationsUrl = (base) =>
  `${base}/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=500&startTime=1`;

export const messagesUrl = (base, id, pageSize = 200) =>
  `${base}/v1/users/ME/conversations/${encodeURIComponent(id)}/messages?view=msnp24Equivalent|supportsMessageProperties&pageSize=${pageSize}&startTime=1`;

const CHANNEL_TYPES = new Set(['space', 'topic']);

export function isChat(conv, { includeChannels = false } = {}) {
  const id = conv?.id || '';
  if (!id.startsWith('19:')) return false; // 48:notifications, 48:calllogs, bots...
  const tt = conv.threadProperties?.threadType;
  return includeChannels || !CHANNEL_TYPES.has(tt);
}

export function normalizeChat(conv) {
  const tp = conv.threadProperties || {};
  return {
    id: conv.id,
    topic: tp.topic || tp.spaceThreadTopic || null,
    type: tp.threadType || (conv.id.endsWith('@unq.gbl.spaces') ? 'oneOnOne' : 'chat'),
    lastMessageId: conv.lastMessage?.id ?? null,
    lastMessageTime: conv.lastMessage?.originalarrivaltime || conv.lastMessage?.composetime || null,
  };
}

const CONTENT_TYPES = /^(Text|RichText)/i;

export function isContentMessage(m) {
  if (!CONTENT_TYPES.test(m.messagetype || '')) return false;
  if (m.properties?.deletetime) return false;
  return true;
}

function parseMaybeJson(v) {
  if (typeof v !== 'string') return v;
  try { return JSON.parse(v); } catch { return null; }
}

export function extractAttachments(m) {
  const out = [];
  const files = parseMaybeJson(m.properties?.files);
  if (Array.isArray(files)) {
    for (const f of files) {
      out.push({
        type: 'file',
        name: f.fileName || f.title || null,
        url: f.objectUrl || f.fileInfo?.fileUrl || f.fileInfo?.shareUrl || null,
        fileType: f.fileType || null,
      });
    }
  }
  const re = /<img\b[^>]*?\bsrc="([^"]+)"[^>]*>/gi;
  let r;
  while ((r = re.exec(m.content || ''))) {
    if (/emoji|\/assets\//i.test(r[0]) && !/AMSImage/i.test(r[0])) continue;
    out.push({ type: 'image', url: r[1].replace(/&amp;/g, '&') });
  }
  return out;
}

export function normalizeMessage(m) {
  const fromUrl = m.from || '';
  return {
    id: m.id,
    conversationId: m.conversationid || m.conversationId || null,
    from: fromUrl.split('/').pop() || null,
    displayName: m.imdisplayname || m.fromDisplayNameInToken || null,
    time: m.originalarrivaltime || m.composetime || null,
    content: stripHtml(m.content),
    attachments: extractAttachments(m),
  };
}

export const safeName = (id) => id.replace(/[^A-Za-z0-9._-]/g, '_').slice(0, 180);

export class HttpError extends Error {
  constructor(url, status, body) {
    super(`HTTP ${status}`);
    Object.assign(this, { url, status, body: String(body ?? '').slice(0, 500) });
  }
}

async function getJson(fetchJson, url) {
  const { status, body } = await fetchJson(url);
  if (status < 200 || status >= 300) throw new HttpError(url, status, typeof body === 'string' ? body : JSON.stringify(body));
  if (typeof body === 'string') {
    try { return JSON.parse(body); } catch { throw new HttpError(url, status, body); }
  }
  return body;
}

async function* pages(fetchJson, firstUrl, base, maxPages) {
  const seen = new Set();
  let url = firstUrl;
  for (let i = 0; url && i < maxPages && !seen.has(url); i++) {
    seen.add(url);
    const json = await getJson(fetchJson, url);
    yield json;
    url = rebase(json?._metadata?.backwardLink, base);
  }
}

export async function listChats(fetchJson, base, opts = {}) {
  const chats = [];
  for await (const page of pages(fetchJson, conversationsUrl(base), base, opts.maxPages ?? 50)) {
    for (const c of page.conversations || []) if (isChat(c, opts)) chats.push(c);
  }
  return chats;
}

// Messages arrive newest-first; walk backward until we hit what we already have.
export async function fetchNewMessages(fetchJson, base, chatId, lastId, opts = {}) {
  const raw = [];
  outer: for await (const page of pages(fetchJson, messagesUrl(base, chatId, opts.pageSize), base, opts.maxPages ?? 1000)) {
    const msgs = page.messages || [];
    if (!msgs.length) break;
    for (const m of msgs) {
      if (lastId && cmpId(m.id, lastId) <= 0) break outer;
      raw.push(m);
    }
  }
  raw.sort((a, b) => cmpId(a.id, b.id));
  const seen = new Set();
  return raw.filter((m) => !seen.has(m.id) && seen.add(m.id));
}

export function loadState(file) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return { chats: {} }; }
}

function writeJsonAtomic(file, data) {
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(data, null, 1));
  fs.renameSync(tmp, file);
}

export function writeDebug(outDir, err, extra = {}) {
  fs.mkdirSync(outDir, { recursive: true });
  const file = path.join(outDir, `debug-${new Date().toISOString().replace(/[:.]/g, '-')}.json`);
  fs.writeFileSync(file, JSON.stringify({
    error: err?.message || String(err),
    url: err?.url ?? null,
    status: err?.status ?? null,
    body: err?.body ?? null,
    stack: (err?.stack || '').split('\n').slice(0, 4).join('\n'),
    ...extra,
  }, null, 1));
  return file;
}

// Main entry. Returns counts only; never message contents.
export async function runExport({ fetchJson, base, outDir, stateFile, includeChannels = false, chatFilter = null, pageSize = 200 }) {
  fs.mkdirSync(outDir, { recursive: true });
  const state = loadState(stateFile);
  state.chats ||= {};
  const res = { chats: 0, newMsgs: 0, skipped: 0, errors: 0, debugFile: null };

  let convs;
  try {
    convs = await listChats(fetchJson, base, { includeChannels });
  } catch (e) {
    res.errors++;
    res.debugFile = writeDebug(outDir, e, { stage: 'listChats' });
    return res;
  }
  if (chatFilter) convs = convs.filter((c) => c.id.includes(chatFilter));

  const index = [];
  for (const conv of convs) {
    const chat = normalizeChat(conv);
    const st = state.chats[chat.id] || {};
    const file = safeName(chat.id) + '.jsonl';
    index.push({ ...chat, file });
    res.chats++;
    if (st.lastId && chat.lastMessageId && cmpId(chat.lastMessageId, st.lastId) <= 0) { res.skipped++; continue; }
    try {
      const raw = await fetchNewMessages(fetchJson, base, chat.id, st.lastId, { pageSize });
      if (raw.length) {
        const lines = raw.filter(isContentMessage).map((m) => JSON.stringify({ ...normalizeMessage(m), conversationId: chat.id }));
        if (lines.length) fs.appendFileSync(path.join(outDir, file), lines.join('\n') + '\n');
        res.newMsgs += lines.length;
        const last = raw[raw.length - 1];
        state.chats[chat.id] = { lastId: last.id, lastTime: last.originalarrivaltime || last.composetime || null, file };
        writeJsonAtomic(stateFile, state); // checkpoint per chat so a crash doesn't duplicate lines
      }
    } catch (e) {
      res.errors++;
      if (!res.debugFile) res.debugFile = writeDebug(outDir, e, { stage: 'messages', chatId: chat.id });
      if (e.status === 401 || e.status === 403) break; // token dead; no point hammering
    }
  }
  writeJsonAtomic(path.join(outDir, 'chats.json'), index);
  state.lastRun = new Date().toISOString();
  writeJsonAtomic(stateFile, state);
  return res;
}

export function totalMessages(outDir) {
  let n = 0;
  for (const f of fs.existsSync(outDir) ? fs.readdirSync(outDir) : []) {
    if (!f.endsWith('.jsonl')) continue;
    const buf = fs.readFileSync(path.join(outDir, f));
    for (let i = 0; i < buf.length; i++) if (buf[i] === 10) n++;
  }
  return n;
}

export function summaryLine(res, outDir) {
  return `${res.chats} chats (${res.skipped} unchanged), ${res.newMsgs} new msgs, ${totalMessages(outDir)} total msgs, ${res.errors} errors` +
    (res.debugFile ? ` -> debug: ${res.debugFile}` : '');
}

// Offline transport: fixture = { base, responses: { "<url>": <json body> | {status, body} } }
export function fixtureFetcher(fixture) {
  return async (url) => {
    const hit = fixture.responses[url];
    if (hit === undefined) return { status: 404, body: `no fixture for ${url}` };
    if (hit && typeof hit === 'object' && 'status' in hit && 'body' in hit) return hit;
    return { status: 200, body: hit };
  };
}
