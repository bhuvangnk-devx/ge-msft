#!/usr/bin/env bash
# Source-level brand guard: fails if a merge changed the brand files, dropped the brand wiring, or
# hard-coded the upstream product name in the UI. Needs no build, so it runs first on every PR.
# (scripts/check-brand.sh then checks the built output.)
#
#   scripts/brand-guard.sh            check
#   scripts/brand-guard.sh --update   re-record brands/<brand>.lock after an approved brand change
set -euo pipefail
cd "$(dirname "$0")/.."

fail() { echo "brand guard failed: $*" >&2; exit 1; }
[ -f brands/current ] || fail "brands/current is missing, so the build would fall back to the upstream look"
brand="$(tr -d '[:space:]' < brands/current)"
dir="brands/$brand"
lock="brands/$brand.lock"
[ "$brand" != default ] || fail "brands/current says 'default'; it must name this repository's brand"
[ -f "$dir/brand.json" ] || fail "$dir/brand.json is missing"

fingerprint() { find "$dir" brands/current -type f ! -name '.DS_Store' | LC_ALL=C sort | xargs shasum -a 256; }

if [ "${1:-}" = --update ]; then
  fingerprint > "$lock"
  echo "recorded $(wc -l < "$lock" | tr -d ' ') brand files in $lock; commit it with the brand change"
  exit 0
fi

# 1. The brand files are exactly the approved ones.
[ -f "$lock" ] || fail "$lock is missing; run scripts/brand-guard.sh --update once and commit it"
if ! diff -u "$lock" <(fingerprint) >/tmp/brand-guard.diff; then
  cat /tmp/brand-guard.diff >&2
  fail "files in $dir changed. If this is an approved brand change, run scripts/brand-guard.sh --update and commit the lock"
fi

# 2. The code that applies the brand is still in place.
wired=(
  "tools/brand/brand.mjs:export"
  "packages/web-shell/brand-assets.ts:brand"
  "packages/web-shell/vite.config.ts:brand"
  "packages/web-shell/src/brand.ts:brand"
  "packages/web-shell/src/taskpane/components/Toolbar.tsx:brand."
  "packages/web-shell/src/taskpane/components/Composer.tsx:brand."
  "packages/web-shell/src/taskpane/main.tsx:brand."
  "tools/release/common.mjs:loadBrand"
)
for entry in "${wired[@]}"; do
  file="${entry%%:*}" needle="${entry#*:}"
  [ -f "$file" ] || fail "$file is gone, so the brand is no longer applied there"
  grep -qF "$needle" "$file" || fail "$file no longer reads the brand (expected '$needle')"
done

# 3. No upstream product name hard-coded where users see it (comments and tests are ignored).
hits="$(grep -rnE "Gemini Enterprise|Ask Gemini|Let Gemini|Open Gemini" \
  packages/web-shell/src --include='*.tsx' --include='*.html' 2>/dev/null \
  | grep -vE '\.test\.|\.spec\.' \
  | grep -vE ':[0-9]+:[[:space:]]*(//|\*|/\*|\{/\*)' || true)"
hits+="$(grep -nE "Gemini Enterprise" packages/web-shell/*.html 2>/dev/null || true)"
[ -z "$hits" ] || fail "the upstream name is hard-coded in the UI; read it from brand instead:
$hits"

echo "brand guard passed: $brand ($(wc -l < "$lock" | tr -d ' ') files locked, wiring intact)"
