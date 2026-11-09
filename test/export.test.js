import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { runExport, fixtureFetcher, stripHtml, chatsvcBase, rebase, cmpId, messagesUrl } from '../src/core.js';

const here = path.dirname(fileURLToPath(import.meta.url));
const FX = path.join(here, 'fixtures', 'sample.json');
const fixture = () => JSON.parse(fs.readFileSync(FX, 'utf8'));
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'teams-export-'));
const readJsonl = (f) => fs.readFileSync(f, 'utf8').trim().split('\n').map((l) => JSON.parse(l));
const ONE = '19:11111111-1111-1111-1111-111111111111_22222222-2222-2222-2222-222222222222@unq.gbl.spaces';

test('helpers', () => {
  assert.equal(stripHtml('<p>a&nbsp;&amp; <b>b</b></p><div>c</div>x<br/>y &#8364;&#x41;'), 'a & b\nc\nx\ny €A');
  assert.equal(chatsvcBase('https://teams.microsoft.com/api/chatsvc/amer/v1/users/ME/conversations?x'), 'https://teams.microsoft.com/api/chatsvc/amer');
  assert.equal(chatsvcBase('https://apac.ng.msg.teams.microsoft.com/v1/users/ME/properties'), 'https://apac.ng.msg.teams.microsoft.com');
  assert.equal(chatsvcBase('https://teams.microsoft.com/api/mt/emea/beta/users/ME/x'), null);
  assert.equal(rebase('https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations?a=1', 'https://b'), 'https://b/v1/users/ME/conversations?a=1');
  assert.ok(cmpId('999', '1000') < 0 && cmpId('1762000004000', '1762000003000') > 0);
});

test('fixture export -> JSONL, index, state; then incremental', async () => {
  const dir = tmp(), out = path.join(dir, 'out'), stateFile = path.join(dir, 'state.json');
  const fx = fixture();
  const res = await runExport({ fetchJson: fixtureFetcher(fx), base: fx.base, outDir: out, stateFile });
  assert.deepEqual({ ...res, debugFile: undefined }, { chats: 3, newMsgs: 6, skipped: 0, errors: 0, debugFile: undefined });

  const index = JSON.parse(fs.readFileSync(path.join(out, 'chats.json'), 'utf8'));
  assert.deepEqual(index.map((c) => c.type), ['chat', 'chat', 'meeting']); // channel + 48:notifications excluded
  assert.equal(index[1].topic, 'Project &quot;X&quot; sync');

  const one = readJsonl(path.join(out, index[0].file));
  assert.deepEqual(one.map((m) => m.content), [
    'Hi Alice\nline two <tag> €5', 'Here is the doc', 'ok 👍', 'See you at 10 & bring the slides',
  ]); // both pages, oldest first
  assert.deepEqual(Object.keys(one[0]), ['id', 'conversationId', 'from', 'displayName', 'time', 'content', 'attachments']);
  assert.equal(one[0].from, '8:orgid:11111111-1111-1111-1111-111111111111');
  assert.equal(one[0].conversationId, ONE);
  assert.equal(one[0].time, '2025-11-01T12:26:41.000Z');
  assert.deepEqual(one[1].attachments, [{ type: 'file', name: 'Q3-plan.docx', url: 'https://contoso-my.sharepoint.com/personal/alice/Documents/Q3-plan.docx', fileType: 'docx' }]);

  const grp = readJsonl(path.join(out, index[1].file));
  assert.deepEqual(grp.map((m) => m.content), ['Kickoff for Bob Example', 'Screenshot:']); // control + deleted dropped
  assert.equal(grp[1].attachments[0].type, 'image');
  assert.ok(!fs.existsSync(path.join(out, index[2].file))); // empty meeting chat -> no file

  const state = JSON.parse(fs.readFileSync(stateFile, 'utf8'));
  assert.equal(state.chats[ONE].lastId, '1762000004000');

  // Rerun with unchanged data: zero message requests.
  const calls = [];
  const fetch2 = fixtureFetcher(fx);
  const res2 = await runExport({ fetchJson: (u) => (calls.push(u), fetch2(u)), base: fx.base, outDir: out, stateFile });
  assert.equal(res2.newMsgs, 0);
  assert.equal(res2.skipped, 2);
  assert.ok(calls.every((u) => !u.includes('/messages') || u.includes('meeting')));

  // New message arrives in the 1:1: only it is appended, page 2 not fetched.
  const t = '1762000005000';
  fx.responses[fx.base + '/v1/users/ME/conversations?view=msnp24Equivalent&pageSize=500&startTime=1'].conversations[0].lastMessage.id = t;
  const p1 = fx.responses[messagesUrl(fx.base, ONE)];
  p1.messages.unshift({ ...p1.messages[0], id: t, content: '<p>new one</p>', originalarrivaltime: '2025-11-01T12:26:45.000Z' });
  const res3 = await runExport({ fetchJson: fixtureFetcher(fx), base: fx.base, outDir: out, stateFile });
  assert.equal(res3.newMsgs, 1);
  const one2 = readJsonl(path.join(out, index[0].file));
  assert.equal(one2.length, 5);
  assert.equal(one2.at(-1).content, 'new one');
});

test('HTTP failure writes small debug file', async () => {
  const dir = tmp(), out = path.join(dir, 'out');
  const fx = fixture();
  fx.responses[messagesUrl(fx.base, ONE)] = { status: 401, body: 'x'.repeat(2000) };
  const res = await runExport({ fetchJson: fixtureFetcher(fx), base: fx.base, outDir: out, stateFile: path.join(dir, 's.json') });
  assert.equal(res.errors, 1);
  const dbg = JSON.parse(fs.readFileSync(res.debugFile, 'utf8'));
  assert.equal(dbg.status, 401);
  assert.equal(dbg.body.length, 500);
  assert.match(dbg.url, /messages/);
});

test('CLI: --help and offline export print one line, no content', () => {
  const cli = path.join(here, '..', 'src', 'cli.js');
  assert.match(execFileSync('node', [cli, '--help'], { encoding: 'utf8' }), /teams-export login/);
  const dir = tmp();
  const stdout = execFileSync('node', [cli, 'export', '--offline', FX, '--out', path.join(dir, 'out'), '--state', path.join(dir, 's.json')], { encoding: 'utf8' });
  assert.equal(stdout, '3 chats (0 unchanged), 6 new msgs, 6 total msgs, 0 errors\n');
});
