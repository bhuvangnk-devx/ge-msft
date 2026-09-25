#!/usr/bin/env bash
# CI-agnostic release steps for the CIMB deployment. Any CI (Jenkins, Bamboo, a shell) calls these
# in order; the CI itself only provides the environment variables and a signed-in gcloud.
#
#   scripts/cimb-release.sh check                 typecheck, lint, test, build
#   scripts/cimb-release.sh build                 web build with this environment's VITE_* baked in
#   scripts/cimb-release.sh deploy                image → Artifact Registry → Cloud Run (by digest) → smoke test
#   scripts/cimb-release.sh manifests             Office manifests + upload zips for this environment
#   scripts/cimb-release.sh rollback <revision>   move all traffic to an earlier Cloud Run revision
#
# ENVIRONMENT=staging builds development-profile manifests; ENVIRONMENT=production builds the
# cimb-production profile ("CNGPT", real app ids). See deploy/cloudrun/README.md for every variable.
set -euo pipefail
cd "$(dirname "$0")/.."

die() { echo "error: $*" >&2; exit 1; }
need() { for v in "$@"; do [ -n "${!v:-}" ] || die "$v is not set"; done; }

ENVIRONMENT="${ENVIRONMENT:-}"
[ "$ENVIRONMENT" = staging ] || [ "$ENVIRONMENT" = production ] || die "ENVIRONMENT must be staging or production"

# Browser config is public: only these VITE_* keys may reach the bundle (a secret named VITE_* would ship).
ALLOWED_VITE="VITE_GCP_PROJECT VITE_GCP_LOCATION VITE_GE_ENGINE VITE_GE_COLLECTION VITE_GE_ASSISTANT
VITE_WIF_POOL_ID VITE_WIF_PROVIDER_ID VITE_ENTRA_TENANT_ID VITE_ENTRA_CLIENT_ID VITE_ENTRA_AUTHORITY
VITE_WIF_ID_TOKEN_SCOPES VITE_GRAPH_SCOPES VITE_GE_MODEL_ID VITE_GE_SKILL_IDS VITE_NOTEBOOK_ID
VITE_GE_COMMAND_PLANNER_SKILL VITE_GE_SURFACE_COMMANDER_SKILL VITE_PROXY_URL"

check_vite_env() {
  local key
  for key in $(compgen -e | grep '^VITE_' || true); do
    grep -qw "$key" <<<"$ALLOWED_VITE" || die "$key is not an allowed browser setting; remove it from the CI environment"
  done
}

check_region() {
  need GCP_REGION
  if [ -n "${ALLOWED_REGIONS:-}" ]; then
    grep -qw "$GCP_REGION" <<<"${ALLOWED_REGIONS//,/ }" || die "GCP_REGION $GCP_REGION is outside ALLOWED_REGIONS"
  fi
}

image() { echo "$GCP_REGION-docker.pkg.dev/$GCP_PROJECT/$ARTIFACT_REPO/web"; }

origin() {
  if [ -n "${WEB_ORIGIN:-}" ]; then echo "$WEB_ORIGIN"; return; fi
  gcloud run services describe "$CLOUD_RUN_SERVICE" --project "$GCP_PROJECT" --region "$GCP_REGION" \
    --format='value(status.url)'
}

cmd_check() {
  bun install --frozen-lockfile
  bun run typecheck
  bun run lint
  bun run test
  bun run build
  bun run skills:check
}

cmd_build() {
  need VITE_GCP_PROJECT VITE_GCP_LOCATION VITE_GE_ENGINE VITE_WIF_POOL_ID VITE_WIF_PROVIDER_ID \
    VITE_ENTRA_TENANT_ID VITE_ENTRA_CLIENT_ID
  check_vite_env
  bun install --frozen-lockfile
  bun run build
  bun run release:web # fails on placeholders, localhost origins or secrets in the bundle
  rm -rf deploy/cloudrun/public
  cp -R packages/web-shell/dist-web deploy/cloudrun/public
}

cmd_deploy() {
  need GCP_PROJECT ARTIFACT_REPO CLOUD_RUN_SERVICE RUNTIME_SERVICE_ACCOUNT
  check_region
  [ -d deploy/cloudrun/public ] || die "run '$0 build' first"
  local tag img digest url
  tag="$(git rev-parse HEAD)"
  img="$(image)"
  gcloud auth configure-docker "$GCP_REGION-docker.pkg.dev" --quiet
  docker build -t "$img:$tag" deploy/cloudrun
  docker push "$img:$tag"
  digest="$(gcloud artifacts docker images describe "$img:$tag" --format='value(image_summary.digest)')"
  # A dedicated, role-less runtime identity: the service only serves static files.
  gcloud run deploy "$CLOUD_RUN_SERVICE" --image "$img@$digest" \
    --service-account "$RUNTIME_SERVICE_ACCOUNT" \
    --project "$GCP_PROJECT" --region "$GCP_REGION" --port 8080 --quiet
  url="$(origin)"
  echo "deployed $ENVIRONMENT: $img@$digest (commit $tag) at $url"

  local headers
  headers="$(curl -fsSI "$url/taskpane.html")"
  grep -qi '^cache-control: no-cache' <<<"$headers" || die "taskpane.html is cacheable"
  if grep -qi '^x-frame-options' <<<"$headers"; then die "X-Frame-Options breaks Office"; fi
  curl -fsS -o /dev/null "$url/auth-redirect.html"
  curl -fsS -o /dev/null "$url/commands.html"
  echo "smoke test passed"
}

cmd_manifests() {
  need GCP_PROJECT CLOUD_RUN_SERVICE
  check_region
  local url
  url="$(origin)"
  if [ "$ENVIRONMENT" = production ]; then
    need GE_PROD_WEB_DOMAIN # must match the Entra Application ID URI; never the *.run.app fallback
    export GE_PROD_ENTRA_CLIENT_ID="${GE_PROD_ENTRA_CLIENT_ID:-${VITE_ENTRA_CLIENT_ID:-}}"
    bun tools/release/manifests-generate.mjs --profile cimb-production
    bun tools/release/manifests-validate.mjs --profile cimb-production
    bun run package:cimb
    echo "upload: dist/release/cimb-production-m365-v*.zip (or dist/package/cimb-production/centralized/)"
  else
    export GE_DEV_WEB_ORIGIN="$url"
    bun tools/release/manifests-generate.mjs --profile development
    bun tools/release/manifests-validate.mjs --profile development
    bun run package:dev
    echo "upload: dist/release/development-m365-v*.zip (or dist/package/development/centralized/)"
  fi
}

cmd_rollback() {
  local revision="${1:-}"
  [ -n "$revision" ] || die "usage: $0 rollback <revision>"
  need GCP_PROJECT CLOUD_RUN_SERVICE
  check_region
  gcloud run revisions list --service "$CLOUD_RUN_SERVICE" --project "$GCP_PROJECT" \
    --region "$GCP_REGION" --format='value(metadata.name)' | grep -qx "$revision" \
    || die "$revision is not a revision of $CLOUD_RUN_SERVICE"
  gcloud run services update-traffic "$CLOUD_RUN_SERVICE" --to-revisions="$revision=100" \
    --project "$GCP_PROJECT" --region "$GCP_REGION"
  echo "traffic moved to $revision on $ENVIRONMENT"
}

case "${1:-}" in
  check) cmd_check ;;
  build) cmd_build ;;
  deploy) cmd_deploy ;;
  manifests) cmd_manifests ;;
  rollback) shift; cmd_rollback "$@" ;;
  *) die "usage: $0 {check|build|deploy|manifests|rollback <revision>}" ;;
esac
