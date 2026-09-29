#!/usr/bin/env bash
# Bring another repo's branch into this one without losing the brand. Merges on a new sync branch,
# keeps this repo's brand files on any conflict, then runs the brand guard, build and brand check.
# It never pushes: review the result, then push the sync branch and open a pull request.
#
#   scripts/sync-upstream.sh                                   # upstream main (vamsiramakrishnan/ge-msft)
#   scripts/sync-upstream.sh https://github.com/harshitajakiya/ge-msft.git merged-ge-fixes
#   SKIP_TESTS=1 scripts/sync-upstream.sh <repo> <branch>       # skip the full test run
set -euo pipefail
cd "$(dirname "$0")/.."

fail() { echo "sync stopped: $*" >&2; exit 1; }

verify() {
  scripts/brand-guard.sh
  bun install --silent
  bun run build >/dev/null
  scripts/check-brand.sh
  [ -n "${SKIP_TESTS:-}" ] || { bun run typecheck && bun run test; }
}

if [ "${1:-}" = --verify ]; then # after resolving conflicts by hand
  verify
  echo "done: the brand is intact. Push this branch and open a pull request."
  exit 0
fi

repo="${1:-https://github.com/vamsiramakrishnan/ge-msft.git}"
ref="${2:-main}"

[ -z "$(git status --porcelain --untracked-files=no)" ] || fail "commit or stash your changes first"
scripts/brand-guard.sh >/dev/null || fail "the brand guard fails before the merge; fix that first"

base="$(git branch --show-current)"
name="$(basename "${repo%.git}")-${ref//\//-}"
sync="sync/${name}-$(date +%Y%m%d-%H%M)"

git config merge.ours.driver true # used by .gitattributes for the brand files
git fetch --quiet "$repo" "$ref"
incoming="$(git rev-parse FETCH_HEAD)"
if git merge-base --is-ancestor "$incoming" HEAD; then echo "nothing new in $repo $ref"; exit 0; fi

git switch --quiet -c "$sync"
echo "merging $repo $ref (${incoming:0:7}) into $sync (from $base)"
if ! git merge --no-edit --no-ff -m "Merge $repo $ref into $base" "$incoming"; then
  # Brand files belong to this repo: always keep ours.
  while IFS= read -r f; do
    case "$f" in brands/*) git checkout --ours -- "$f" && git add -- "$f" && echo "kept our brand file: $f" ;; esac
  done < <(git diff --name-only --diff-filter=U)
  left="$(git diff --name-only --diff-filter=U)"
  if [ -n "$left" ]; then
    echo >&2
    echo "Resolve these conflicts by hand, keeping every brand.* reference:" >&2
    while IFS= read -r f; do echo "  $f" >&2; done <<<"$left"
    echo >&2
    echo "Then: git add <files> && git commit --no-edit && scripts/sync-upstream.sh --verify" >&2
    exit 1
  fi
  git commit --quiet --no-edit
fi

verify

echo
echo "done: $sync is merged and the brand is intact. Next:"
echo "  git push -u origin $sync   # then open a pull request into $base"
