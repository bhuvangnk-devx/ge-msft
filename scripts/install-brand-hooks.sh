#!/usr/bin/env bash
# One-time setup per clone: every manual merge, merge commit and push then runs the brand guard.
set -euo pipefail
cd "$(dirname "$0")/.."
chmod +x scripts/hooks/* scripts/brand-guard.sh scripts/sync-upstream.sh
git config core.hooksPath scripts/hooks
git config merge.ours.driver true # lets .gitattributes keep our brand files on conflicts
echo "brand hooks installed: merges, merge commits and pushes now run scripts/brand-guard.sh"
