# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

A static, no-build PWA for studying English vocabulary and Chinese idioms/articles. There is no
framework and no bundler — every page is a single self-contained `.html` file with inline `<style>`
and `<script>`. Data comes from Excel workbooks in `docs/` that get converted to JSON once and served
as static files.

Pages:
- `index.html` — landing hub. Links to `english.html` / `chinese.html`, manages the service worker
  (update check/apply, enable/disable), holds app settings (pronunciation, special pronunciation
  overrides), and provides a raw `localStorage` inspector ("紀錄") plus GitHub Gist export/import for
  cross-device sync.
- `english.html` — English vocabulary app: word list ("normal" mode), spaced-style review mode, quiz
  mode, custom word add/edit (with auto-lookup via external dictionary/translation APIs), favorites,
  and pronunciation (Web Speech API).
- `chinese.html` — superset of `english.html`. It embeds essentially the same vocabulary engine
  (normal/review/quiz modes, storage helpers, custom words, favorites — copy-pasted, not shared) and
  additionally implements idiom (成語) study, article (文章) reading with idiom highlighting + TTS
  playback, and an error-correction (改錯) quiz, with zhuyin rendering.

## Commands

```bash
npm start                 # serve the static site (http-server) on http://localhost:8000
node script/run.js        # regenerate json/english.json and json/chinese.json from docs/*.xlsx
node script/convert_en.js # regenerate only json/english.json
node script/convert_ch.js # regenerate only json/chinese.json (成語/文章/改錯/特殊發音 + json/pronunciation.js data)
node script/encrypt-token-simple.js  # interactive encrypt/decrypt for the GitHub PAT baked into index.html
```

There is no test suite and no linter configured. `npm run convert` in `package.json` is stale (points
at a nonexistent `convert.js` at the repo root) — use `node script/run.js` instead.

## Architecture

### Data pipeline
`docs/*.xlsx` (成語.xlsx, 文章.xlsx, 改錯.xlsx, 特殊發音.xlsx, english.xlsx for English sheets) are
converted by `script/convert_en.js` / `script/convert_ch.js` into `json/english.json` and
`json/chinese.json`, which the HTML pages `fetch()` at runtime. `script/run.js` just runs both
converters in sequence. Re-run the converters after editing the source workbooks; the JSON files are
committed and are what the deployed app actually reads.

### Duplicated app logic (english.html vs chinese.html)
`chinese.html` was forked from `english.html` rather than importing shared code — the storage
helpers, quiz engine, word modal, custom-word CRUD, and pronunciation logic exist almost identically
in both files (same function names: `initQuizPanel`, `startQuiz`, `saveCustomWord`,
`getFavoriteFromStorage`, `initializeLocalStorage`, `migrateFromLegacy`, etc.). **Bug fixes or
behavior changes to the shared vocabulary engine generally need to be applied in both files.**

### localStorage schema & versioning
Each of the three HTML files independently defines `STORAGE_CONFIG.VERSION` (currently `'1.0'`) and
runs `initializeLocalStorage()` on load, which checks the `_storageVersion` marker key and calls
`migrateFromLegacy()` (first run) or `migrateVersion(from, to)` (version bump) to reshape old data.
Top-level `localStorage` keys:
- `setting` — global app settings (pronunciation rate, etc.)
- `english` — `{ favorite, quiz, customWords, quizSessions, ... }` for the English app
- `chinese` — Chinese app data
- `_gistId`, `_storageVersion` — internal, prefixed with `_` and excluded from user-facing
  import/export "reset" operations

Because the version constant is duplicated per file, a schema change must bump/update the migration
logic in `index.html`, `english.html`, and `chinese.html` together, or pages will disagree about the
current shape of `english`/`chinese` data.

`index.html`'s "紀錄" viewer (`showStorageRecords()` / `formatJSON()`) is a generic recursive
localStorage JSON viewer — it renders whatever is under each top-level key, including nested data
like `english.quizSessions`.

### PWA / service worker
`service-worker.js` implements a cache-first strategy keyed by `CACHE_VERSION`; bump that constant to
invalidate old caches on deploy. `self.skipWaiting()` is intentionally commented out so a new worker
sits in "waiting" state until the user confirms via the update UI in `index.html`
(`checkForUpdates()` / `updateApp()`, driven by a `SKIP_WAITING` postMessage). `manifest.json` sets
`start_url`/`scope` to `/study/`, so the app is expected to be served from a `/study/` subpath, not
the domain root.

### Cross-device sync
`index.html` can export/import the entire `localStorage` contents as a GitHub Gist
(`showExportDialog()` / `showImportDialog()`, `api.github.com/gists`). The PAT used for this is not
stored in plaintext — it's obfuscated in `ENCRYPTED_TOKEN` using a trivial char-code ±125 hex cipher
(`decryptToken()` in `index.html`, `script/encrypt-token-simple.js` is the paired CLI for
generating/reading it). This is obfuscation, not real security — treat `ENCRYPTED_TOKEN` as sensitive
and regenerate it with `script/encrypt-token-simple.js` when rotating the token.
