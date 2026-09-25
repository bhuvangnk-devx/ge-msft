# CIMB release: Cloud Run hosting and production manifests

`scripts/cimb-release.sh` holds every release step, so any CI can run them. The GitHub Actions files in
`.github/workflows/` show the same stages as a reference. Your CI only has to set the environment
variables below and sign gcloud in (keyless, see [Security checklist](#security-checklist)).

## Stages

| Trigger | Run |
| --- | --- |
| Every push and pull request | `ENVIRONMENT=staging scripts/cimb-release.sh check` |
| Merge to `main` | `ENVIRONMENT=staging scripts/cimb-release.sh build`, then `deploy`, then `manifests` |
| Release tag `v*`, after a named approver signs off | `ENVIRONMENT=production scripts/cimb-release.sh build`, then `deploy`, then `manifests` |
| Rollback, approver only | `ENVIRONMENT=production scripts/cimb-release.sh rollback <revision>` |

Staging and production each need their own build: the `VITE_*` values are baked into the JavaScript,
so a staging image can't be promoted to production.

The Cloud Run release and the Office add-in are separate steps. A new manifest only needs uploading to
the Microsoft 365 admin center when the name, icons, commands, permissions, domain or apps change.

## Variables

**Both environments** (values differ per environment):

| Variable | Example | Notes |
| --- | --- | --- |
| `ENVIRONMENT` | `staging` / `production` | Picks the manifest profile |
| `GCP_PROJECT`, `GCP_REGION` | `cimb-cngpt-prod`, `asia-southeast2` | Region must match data residency |
| `ALLOWED_REGIONS` | `asia-southeast2` | Optional; the build fails outside this list |
| `ARTIFACT_REPO`, `CLOUD_RUN_SERVICE` | `cngpt`, `cngpt-web` | |
| `RUNTIME_SERVICE_ACCOUNT` | `cngpt-web-runtime@…` | Dedicated account with no roles |
| `WEB_ORIGIN` | `https://cngpt.cimbniaga.co.id` | Optional; defaults to the Cloud Run URL |
| `VITE_GCP_PROJECT`, `VITE_GCP_LOCATION`, `VITE_GE_ENGINE` | | Required browser settings |
| `VITE_WIF_POOL_ID`, `VITE_WIF_PROVIDER_ID` | | The *workforce* pool users sign in through |
| `VITE_ENTRA_TENANT_ID`, `VITE_ENTRA_CLIENT_ID` | | Entra app registration |
| `VITE_GE_COLLECTION`, `VITE_GE_ASSISTANT`, `VITE_GRAPH_SCOPES`, … | | Optional; only the keys listed in the script are allowed |

**Staging only** (development-profile manifests, all optional):
`GE_DEV_APP_ID`, `GE_DEV_OFFICE_XML_APP_ID`, `GE_DEV_OUTLOOK_APP_ID`, `GE_DEV_DEVELOPER_NAME`. Set the
three ids once so staging installs keep the same identity.

**Production only** (`cimb-production` manifest profile):

| Variable | Notes |
| --- | --- |
| `GE_PROD_WEB_DOMAIN` | Required. Must match the Entra Application ID URI host |
| `GE_PROD_APP_ID`, `GE_PROD_OFFICE_XML_APP_ID`, `GE_PROD_OUTLOOK_APP_ID` | Generate once with `uuidgen`, never change them, all different |
| `GE_PROD_DEVELOPER_NAME`, `GE_PROD_WEBSITE_URL`, `GE_PROD_PRIVACY_URL`, `GE_PROD_TERMS_URL`, `GE_PROD_SUPPORT_URL` | HTTPS only |
| `GE_PROD_ENTRA_CLIENT_ID` | Optional; defaults to `VITE_ENTRA_CLIENT_ID` |
| `GE_PROD_SURFACES` | Optional; default `word,excel,powerpoint,outlook` |

**Adding OneNote later:** set `GE_PROD_SURFACES=word,excel,powerpoint,outlook,onenote` and
`GE_PROD_ONENOTE_APP_ID`. The build then also writes `onenote/onenote.manifest.xml`. OneNote only
accepts the classic XML format, so it's uploaded separately from the unified package.

## Security checklist

These come from a security review of the pipeline. Check them in your CI before the first
production deploy.

1. **Keep public upstream code away from secrets.** Pull GitHub updates into a branch through a
   reviewed pull request. Don't let a pipeline that holds deploy credentials run code straight from
   upstream, and drop any CI file upstream adds before it lands.
2. **Pin the CI's Google identity.** CI signs in to Google Cloud through a *workload* identity pool
   (keyless, no service-account keys). Give staging and production separate providers or deploy
   accounts, each with an attribute condition that pins your repository and environment, so no
   other branch or repo can deploy.
3. **Gate production.** Only named approvers can run the production `deploy` and `rollback`, and only
   from `main` or a `v*` tag. Protect `v*` tags.
4. **Keep the deploy credential out of the install step.** Sign gcloud in after `build` (after
   `bun install` and its install scripts have run), or run `build` and `deploy` as separate jobs.
5. **Pin images.** Pin the CI images and `nginx` in `deploy/cloudrun/Dockerfile` to `@sha256:` digests.
6. **Review permissions before production.** The Outlook manifest asks for `ReadWriteMailbox`, and the
   Graph scopes include `Files.Read.All` and `Sites.Read.All`. Narrow them if the pilot doesn't need them.
