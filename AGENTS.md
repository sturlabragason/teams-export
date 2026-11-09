# AGENTS.md: using teams-export as an AI agent

teams-export exports the user's own Microsoft Teams chats to JSONL. Keep runs cheap: read the one
stdout line, not the data.

## Install

```bash
npm install
npx playwright install chromium     # or set TEAMS_EXPORT_BROWSER=/path/to/chrome
npm test                            # offline, about 1 s
```

Node 18 or later.

## Log in: the human does this, not you

```bash
node src/cli.js login
```

This opens a **visible** browser on teams.microsoft.com. The human signs in and completes MFA in that
window. Tell them to do so, then wait. The window closes by itself once chats load and the command
prints `login ok (chatsvc <host>)`. It gives up after 10 minutes, or when the window is closed, with
`login not confirmed: timeout|closed`. The session is saved in `./profile/`.

Never type, ask for, store or log credentials, MFA codes, tokens or headers. Login needs a desktop
session (a display), so it can't run on a headless server.

## Export

```bash
node src/cli.js export                      # headless, incremental
node src/cli.js export --headful            # show the browser (use if headless is blocked)
node src/cli.js export --wait 180           # allow a slow Teams load (default 90 s)
node src/cli.js export --chat <substr>      # only chats whose id contains <substr>
node src/cli.js export --include-channels   # also team channels (default: chats only)
node src/cli.js --help
```

Every run is incremental. `state.json` records the last message per chat, so a rerun fetches only
new messages and is safe to repeat after any failure.

It prints exactly one line, for example `42 chats (40 unchanged), 12 new msgs, 3100 total msgs, 0 errors`.
Exit code 0 means success. Report that line to the user and stop.

## Output

- `out/chats.json`: index of `{id, topic, type, lastMessageId, lastMessageTime, file}`.
- `out/<chat-id>.jsonl`: one message per line, oldest first:
  `{id, conversationId, from, displayName, time, content, attachments}`. `content` is plain text;
  attachments are links only.
- `out/debug-<ts>.json` on failure: URL, HTTP status, first 500 characters of the response.

## Failures (exit 1)

The line ends with `-> debug: <path>`. Read only that file.

- `not logged in` or HTTP 401/403: the session expired. Ask the human to run
  `node src/cli.js login`, wait for `login ok`, then rerun `export`.
- `no chatsvc request seen`: try `--headful --wait 180`. If that fails, Teams changed its URLs;
  compare `chatsvcBase()` in `src/core.js` with the URL in the debug file.
- 404 on `/conversations` or `/messages`: the endpoint shape changed; fix `conversationsUrl` or
  `messagesUrl` in `src/core.js`.
- 429/5xx after retries: wait, then rerun.

After a code fix, run `npm test` before you export again. Update `test/fixtures/sample.json` with
**fake** data only.

## Safety rules

- Never commit `profile/` (it is the login session, treat it like a password), `out/` or `state.json`.
  They are git-ignored; keep them that way.
- Never read `out/*.jsonl`, `out/chats.json`, `state.json` or `profile/` into your context. To count
  messages use `wc -l out/*.jsonl | tail -1`.
- Never paste, upload, summarise to third parties or otherwise share exported chats. They are private
  messages of the user and the people they talked to.
- No screenshots, page snapshots or DOM dumps of Teams.

## Layout

- `src/cli.js`: arguments and the one-line output.
- `src/live.js`: Playwright: persistent profile, visible login, header capture, requests.
- `src/core.js`: URLs, pagination, normalising, incremental state. Unit-tested.
