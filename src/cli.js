#!/usr/bin/env node
// stdout contract: exactly one summary line per command. Never message contents.
import fs from 'node:fs';
import path from 'node:path';
import { parseArgs } from 'node:util';
import { runExport, summaryLine, writeDebug, fixtureFetcher } from './core.js';

const HELP = `teams-export - export your own MS Teams chats without Graph (Playwright + chatsvc)

Usage:
  teams-export login  [--profile DIR]             open headful Teams; sign in (SSO/MFA) once
  teams-export export [options]                   incremental export to OUT/<chat>.jsonl + chats.json
  teams-export --help

Options:
  --profile DIR        persistent browser profile         (default ./profile)
  --out DIR            export directory                   (default ./out)
  --state FILE         incremental state file             (default ./state.json)
  --headful            show the browser during export
  --wait SEC           max seconds to wait for Teams auth (default 90)
  --chat SUBSTR        only export chats whose id contains SUBSTR
  --include-channels   also export team channels (default: chats only)
  --offline FILE       use a recorded fixture instead of the network (testing)

Output: one line, e.g. "42 chats (40 unchanged), 12 new msgs, 3100 total msgs, 0 errors".
On failure: "... -> debug: out/debug-<ts>.json" (URL, status, first 500 chars). Exit 1 if errors.
Env: TEAMS_CHATSVC_BASE overrides the detected chatsvc base URL;
     TEAMS_EXPORT_BROWSER uses a specific Chromium/Chrome executable.`;

async function main() {
  const { values: o, positionals } = parseArgs({
    allowPositionals: true,
    options: {
      help: { type: 'boolean', short: 'h' },
      profile: { type: 'string', default: './profile' },
      out: { type: 'string', default: './out' },
      state: { type: 'string', default: './state.json' },
      headful: { type: 'boolean' },
      wait: { type: 'string', default: '90' },
      chat: { type: 'string' },
      'include-channels': { type: 'boolean' },
      offline: { type: 'string' },
    },
  });
  const cmd = positionals[0];
  if (o.help || !cmd) { console.log(HELP); return 0; }

  if (cmd === 'login') {
    const { login } = await import('./live.js');
    console.log('Sign in to Teams in the opened window; it closes itself once chats load.');
    const r = await login({ profileDir: path.resolve(o.profile) });
    console.log(r.ok ? `login ok (chatsvc ${new URL(r.base).host})` : `login not confirmed: ${r.reason}`);
    return r.ok ? 0 : 1;
  }

  if (cmd === 'export') {
    const outDir = path.resolve(o.out);
    let session;
    try {
      if (o.offline) {
        const fx = JSON.parse(fs.readFileSync(o.offline, 'utf8'));
        session = { base: fx.base, fetchJson: fixtureFetcher(fx), close: async () => {} };
      } else {
        const { openSession } = await import('./live.js');
        session = await openSession({ profileDir: path.resolve(o.profile), headless: !o.headful, waitMs: Number(o.wait) * 1000 });
      }
    } catch (e) {
      console.log(`0 chats, 0 msgs, 1 errors (${e.message}) -> debug: ${writeDebug(outDir, e, { stage: 'auth' })}`);
      return 1;
    }
    try {
      const res = await runExport({
        fetchJson: session.fetchJson, base: session.base, outDir, stateFile: path.resolve(o.state),
        includeChannels: !!o['include-channels'], chatFilter: o.chat || null,
      });
      console.log(summaryLine(res, outDir));
      return res.errors ? 1 : 0;
    } finally {
      await session.close();
    }
  }

  console.log(`unknown command "${cmd}"; see --help`);
  return 2;
}

main().then((code) => process.exit(code), (e) => {
  const file = writeDebug(path.resolve('./out'), e, { stage: 'fatal' });
  console.log(`fatal: ${e.message} -> debug: ${file}`);
  process.exit(1);
});
