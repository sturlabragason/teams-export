<p align="right">
  <a href="https://github.com/sturlabragason">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="assets/brand/logo-reversed.svg">
      <img src="assets/brand/logo.svg" alt="Sturla Bragason" width="120">
    </picture>
  </a>
</p>

# teams-export

Export **your own** Microsoft Teams chats to JSONL without Microsoft Graph permissions or an
app registration. A Playwright browser profile holds your normal Teams web sign-in. The tool
takes the chat-service auth headers from a request that Teams web sends itself, then calls the
same internal JSON endpoints directly (`…/chatsvc/<region>/v1/users/ME/conversations` and
`…/conversations/{id}/messages`). It doesn't scrape the DOM.

## Setup

```bash
npm install
npx playwright install chromium     # or set TEAMS_EXPORT_BROWSER=/path/to/chrome
npm test                            # offline test against test/fixtures/sample.json
```

Requires Node 18+ (tested on 22).

## Log in (once, by hand)

```bash
node src/cli.js login
```

A browser window opens on teams.microsoft.com. Complete SSO/MFA. When Teams loads your chats,
the tool detects a chat-service request, waits 5 s so the profile is saved, closes the window
and prints `login ok`. The session is stored in `./profile/`. **Treat that folder like a
password.** It is git-ignored.

Run `login` again when `export` reports `not logged in` (for example when the refresh token expires
or after a password change or Conditional Access prompt).

## Export

```bash
node src/cli.js export              # headless, incremental
node src/cli.js export --headful    # watch it (useful if headless is blocked)
node src/cli.js --help
```

Output:

- `out/chats.json`: an index of `{id, topic, type, lastMessageId, lastMessageTime, file}`
- `out/<chat-id>.jsonl`: one message per line, oldest first:
  `{id, conversationId, from, displayName, time, content, attachments}`. `content` has the HTML
  removed. `attachments` holds `{type:"file", name, url, fileType}` and `{type:"image", url}`.
  The tool records links only and doesn't download attachments.
- `state.json`: the last message id and time for each chat. A rerun skips unchanged chats with
  no message requests and fetches only new messages for the others, paging backward through
  `backwardLink` until it reaches a message it already has.

The command prints one line, such as
`42 chats (40 unchanged), 12 new msgs, 3100 total msgs, 0 errors`. When a step fails, it
writes `out/debug-<ts>.json` with the failing URL, HTTP status and the first 500 characters of
the response, appends `-> debug: <path>` to that line, and exits with code 1.

Default scope: 1:1, group and meeting chats. Team channels are excluded; add
`--include-channels` to include them. Use `--chat <substr>` to limit the run to particular chats.

## How it works / robustness

- The regional base URL comes from the captured request (`teams.microsoft.com/api/chatsvc/emea`,
  `amer.ng.msg.teams.microsoft.com`, and so on). Pagination links that point to a different
  host alias are rewritten to the captured base. You can override it with `TEAMS_CHATSVC_BASE`.
- Headless Chromium has its `HeadlessChrome` user agent changed to `Chrome`.
- HTTP 429/5xx responses are retried up to 3 times and follow `Retry-After`. On 401/403 the run
  stops early.
- `state.json` is saved after each chat, so a crash doesn't duplicate lines.
- Edited messages keep the version they had when exported. Edits made after export are not
  picked up again.

## Caveats

- **These are undocumented internal endpoints.** Microsoft can change them without notice. If
  that happens, look at the debug file and adjust `src/core.js` (URL builders and
  `normalizeMessage`) or `src/live.js` (header capture).
- **Check your organisation's policies first.** Exporting chat history can fall under DLP,
  retention, eDiscovery or acceptable-use rules. Only export data you're allowed to keep
  outside Teams, and store `out/` and `profile/` securely.
- This exports only what your own Teams web session can already see.

## License

[MIT](LICENSE). This project is not affiliated with or endorsed by Microsoft. "Microsoft Teams"
is a trademark of Microsoft. The Sturla Bragason logo in `assets/brand/` is not covered by the
MIT licence; please do not reuse it.

---

<a href="https://github.com/sturlabragason">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="assets/brand/lockup-reversed.svg">
    <img src="assets/brand/lockup.svg" alt="Sturla Bragason, platform and AI engineering" width="240">
  </picture>
</a>

Made by [Sturla Bragason](https://github.com/sturlabragason). Platform and AI engineering.
