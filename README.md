# Phrasal Verbs — Spaced-Repetition Flashcards

A web app for studying the 150 most common English phrasal verbs with a spaced-repetition system (SRS). Built on **Google Sheets + Google Apps Script**, so your progress is stored per-user in the cloud and accessible from any device.

## Features

- **SM-2 spaced repetition** — cards resurface based on how well you know them (Again / Hard / Good / Easy).
- **Three study modes** — Verb → Meaning, Meaning → Verb, or Mixed (switchable mid-session).
- **Daily new-card limit** — configurable; new cards only count against the daily quota once you actually rate them.
- **Category filter** — focus on one of 10 categories (Daily Life, Communication, Emotions, …).
- **Randomized sessions** — due and new cards are shuffled each session (1 new interleaved per 5 reviews).
- **Study again / Study more** — replay the same session, or pull extra new cards when you're caught up.
- **Add new phrasal verbs** — a maintenance form in Settings (see below) adds cards to the shared list, with a category dropdown and duplicate protection.
- **3D flip cards**, progress bar, live counters (🔴 again / 📖 reviews / ✨ new), and keyboard shortcuts.

### Keyboard shortcuts

| Key | Action |
| --- | --- |
| `Space` / `Enter` | Flip card |
| `1` | Again |
| `2` | Hard |
| `3` | Good |
| `4` | Easy |

### Maintaining the word list

Open **Settings → "+ Add a phrasal verb"** to add a card without touching the spreadsheet:

- Four fields — **Phrasal Verb, Meaning, Example, Category**. The card is appended to the shared
  list and appears immediately (no reload).
- **Category** is a dropdown sourced from a dedicated `Categories` tab (column A), with a
  **"+ New category…"** option. A brand-new category is written back to the `Categories` tab
  automatically, so it shows up in the dropdown next time.
- **Duplicate protection** — a warning appears inline as soon as you leave the Phrasal Verb field
  if it already exists, and the server rejects duplicates authoritatively (case-insensitive on the
  verb), so a duplicate can never be saved.

## Project layout

```
.
├── flashcards.html                 # Standalone localStorage version (reference build, no backend)
├── 150 Most Common Phrasal Verbs - Untitled.csv   # Source data
└── gas/                            # Google Apps Script web app
    ├── Code.gs                     # Server: doGet, getInitialData, saveSRS, saveNewToday, saveSettings, resetProgress, addCard
    ├── index.html                  # Client: full UI + google.script.run async saves
    ├── appsscript.json             # Web app manifest (executeAs USER_DEPLOYING, access ANYONE_ANONYMOUS)
    └── deploy.sh                   # One-command release: git push + clasp push --force + clasp deploy
```

## Architecture

- **Data** lives in a Google Sheet ("Common Phrasal Verbs"). The cards tab holds
  `Phrasal Verb | Meaning | Example | Category`; a `Categories` tab (column A) holds the canonical
  category list used by the add-card form; a `Progress` tab (auto-created) stores one row per user:
  `User | SRS Data (JSON) | New Today (JSON) | Settings (JSON)`.
- **Server (`Code.gs`)** is a standalone Apps Script web app that reads the sheet via
  `SpreadsheetApp.openById(...)` and exposes data through a single `getInitialData()` round-trip.
- **Client (`index.html`)** keeps everything in memory and writes back asynchronously with
  `google.script.run` (fire-and-forget; last write wins — fine for a single-user study app).

## Deploying

Requires [`clasp`](https://github.com/google/clasp) (logged in) and `gh`/git for the GitHub side.

```bash
cd gas
./deploy.sh "your message"
```

This commits and pushes to GitHub, then `clasp push --force`es the code and redeploys the web app.
The app runs as the deploying owner (`executeAs: USER_DEPLOYING`, `access: ANYONE_ANONYMOUS`) and
identifies users by a **name they type** (stored in their browser, sent with every call) — so
external users never hit the "unverified app" authorization wall and no Google sign-in is required.

## The standalone version

`flashcards.html` is a self-contained build that uses `localStorage` instead of a backend —
open it directly in a browser to try the SRS logic with no setup. Progress is per-browser only.
