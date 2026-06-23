# Phrasal Verbs — Spaced-Repetition Flashcards

A web app for studying the 150 most common English phrasal verbs with a spaced-repetition system (SRS). Built on **Google Sheets + Google Apps Script**, so your progress is stored per-user in the cloud and accessible from any device.

## Features

- **SM-2 spaced repetition** — cards resurface based on how well you know them (Again / Hard / Good / Easy).
- **Three study modes** — Verb → Meaning, Meaning → Verb, or Mixed (switchable mid-session).
- **Daily new-card limit** — configurable; new cards only count against the daily quota once you actually rate them.
- **Category filter** — focus on one of 10 categories (Daily Life, Communication, Emotions, …).
- **Randomized sessions** — due and new cards are shuffled each session (1 new interleaved per 5 reviews).
- **Study again / Study more** — replay the same session, or pull extra new cards when you're caught up.
- **3D flip cards**, progress bar, live counters (🔴 again / 📖 reviews / ✨ new), and keyboard shortcuts.

### Keyboard shortcuts

| Key | Action |
| --- | --- |
| `Space` / `Enter` | Flip card |
| `1` | Again |
| `2` | Hard |
| `3` | Good |
| `4` | Easy |

## Project layout

```
.
├── flashcards.html                 # Standalone localStorage version (reference build, no backend)
├── 150 Most Common Phrasal Verbs - Untitled.csv   # Source data
└── gas/                            # Google Apps Script web app
    ├── Code.gs                     # Server: doGet, getInitialData, saveSRS, saveNewToday, saveSettings, resetProgress
    ├── index.html                  # Client: full UI + google.script.run async saves
    ├── appsscript.json             # Web app manifest (executeAs USER_ACCESSING, access ANYONE)
    └── deploy.sh                   # One-command release: git push + clasp push + clasp deploy
```

## Architecture

- **Data** lives in a Google Sheet ("150 Most Common Phrasal Verbs"). The cards tab holds
  `Phrasal Verb | Meaning | Example | Category`; a `Progress` tab (auto-created) stores one row
  per user: `Email | SRS Data (JSON) | New Today (JSON) | Settings (JSON)`.
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

This commits and pushes to GitHub, then `clasp push`es the code and redeploys the web app
(`executeAs: USER_ACCESSING` is required so each user gets their own progress row).

## The standalone version

`flashcards.html` is a self-contained build that uses `localStorage` instead of a backend —
open it directly in a browser to try the SRS logic with no setup. Progress is per-browser only.
