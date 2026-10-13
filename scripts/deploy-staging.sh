#!/usr/bin/env bash
# Deploy the add-in web build to the Cloud Run service the client uses, ONLY for a commit that
# passed the quality gate (bun run quality:check). Run from a clean checkout of that commit.
#
#   bash scripts/deploy-staging.sh
set -euo pipefail
cd "$(dirname "$0")/.."

PROJECT="${CLOUD_RUN_PROJECT:-yash-test-495112}"
REGION="${CLOUD_RUN_REGION:-asia-southeast2}"
SERVICE="${CLOUD_RUN_SERVICE:-cngpt-web}"
REPORT="scenarios/.results/quality-check.json"

fail() { echo "deploy refused: $*" >&2; exit 1; }
[ -z "$(git status --porcelain --untracked-files=no)" ] || fail "uncommitted changes; commit, then run bun run quality:check"
# Untracked files under the built folders would ship without being in the checked commit.
[ -z "$(git status --porcelain --untracked-files=all -- packages deploy scripts scenarios)" ] || fail "untracked files in packages/, deploy/, scripts/ or scenarios/; commit or remove them"
[ -f "$REPORT" ] || fail "no quality report; run bun run quality:check"
head="$(git rev-parse HEAD)"
python3 - "$REPORT" "$head" <<'PY' || fail "the quality report is not a pass for this commit; run bun run quality:check"
import json, sys
r = json.load(open(sys.argv[1]))
sys.exit(0 if r.get("ok") and r.get("clean") and r.get("commit") == sys.argv[2] else 1)
PY

bun run build
bun run release:web
rm -rf deploy/cloudrun/public
cp -R packages/web-shell/dist-web deploy/cloudrun/public
(cd deploy/cloudrun && gcloud run deploy "$SERVICE" --source . --project "$PROJECT" --region "$REGION" \
  --allow-unauthenticated --port 8080 --quiet)
echo "deployed $head (quality check passed)"
