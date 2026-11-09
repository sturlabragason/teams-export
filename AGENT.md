# AGENT.md: running and repairing teams-export cheaply

Budget: about 1–3k tokens per run. Follow these rules exactly.

## Run

```bash
cd <this dir> && node src/cli.js export
```

Read the **single stdout line** and nothing else. Exit code 0 means done. Report the line to the user.

## Never

- Never `cat`, `head`, `grep` or otherwise read `out/*.jsonl`, `out/chats.json`, `state.json`
  or `profile/` into context. They contain private messages and are large. To count, use
  `wc -l out/*.jsonl | tail -1`.
- Never take page snapshots or screenshots, and never dump HTML or the DOM.
- Never print tokens or headers. Never commit `profile/`, `out/` or `state.json`.

## On failure (exit 1)

1. Read **only** the debug file named in the line. It is small: url, status, 500-char body.
2. Diagnose from the status:
   - `not logged in` / 401 / 403: ask the user to run `node src/cli.js login` (it needs their
     MFA, so you can't do it). Then rerun.
   - `no chatsvc request seen`: Teams web changed its URLs, or headless mode is blocked. Try
     `--headful --wait 180`. If that fails too, check `chatsvcBase()` in `src/core.js`
     against the URL in the debug file.
   - 404 on `/messages` or `/conversations`: the endpoint shape changed. Update `messagesUrl`
     or `conversationsUrl` in `src/core.js`.
   - 429/5xx after retries: wait and rerun. The run is incremental, so this is safe.
3. Edit the smallest function you can, run `npm test` (offline, about 1 s), then rerun the export.
   When a response shape changes, update `test/fixtures/sample.json` with **fake** data only.

## Layout

- `src/core.js`: pure logic (URLs, pagination, normalize, incremental state). Fully unit-tested.
- `src/live.js`: Playwright: persistent profile, header capture, `context.request` fetcher.
- `src/cli.js`: argument parsing and the one-line output.
