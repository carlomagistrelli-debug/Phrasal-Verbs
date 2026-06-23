#!/bin/bash
# Deploy: commit + push to GitHub, then push + deploy the Apps Script web app.
# Usage: ./deploy.sh "commit / deployment message"
set -e

MSG="${1:-update}"
DIR="$(cd "$(dirname "$0")" && pwd)"   # gas/ — holds .clasp.json
ROOT="$(cd "$DIR/.." && pwd)"          # repo root

# 1. Commit & push to GitHub (skip cleanly when there's nothing to commit)
git -C "$ROOT" add -A
if git -C "$ROOT" diff --cached --quiet; then
  echo "No changes to commit."
else
  git -C "$ROOT" commit -m "$MSG"
  git -C "$ROOT" push
fi

# 2. Push code to Apps Script and redeploy the web app
# --force skips the interactive "manifest changed, overwrite?" prompt, which
# otherwise auto-declines in this non-TTY environment and silently skips the push.
cd "$DIR"
clasp push --force
clasp deploy --deploymentId AKfycbyOMjOs25VRPpNdIUvpvYuDpZ-5j6joMayfhgXQiDaGlrXnEmCm51NlFwBFmjvLnw6a --description "$MSG"
