# Deployment Guide — Outlook Add-in Test Harness

Full step-by-step guide for provisioning infrastructure, deploying the Azure Functions backend, and deploying the add-in to Azure Static Web Apps.

**Read order:** do the sections in order on first deploy. On subsequent deploys, jump to the relevant part (§3 for function changes, §4 for add-in changes).

---

## Prerequisites checklist

Before you start, confirm you have:

- [ ] Azure CLI installed and logged in: `az login && az account set --subscription <your-subscription-id>`
- [ ] Terraform ≥ 1.6: `terraform -version`
- [ ] Azure Functions Core Tools v4: `func --version`
- [ ] uv + Python 3.11+: `uv --version && python3 --version`
- [ ] Node.js v20+: `node --version`
- [ ] An App Registration created per `SETUP.md` Step 1 — you need the **Tenant ID** and **Client ID** before starting Terraform

---

## § 1 — Provision infrastructure with Terraform

### 1.1 Fill in the env-specific tfvars

Two environments, two files — pick the one you're targeting:

| File | When to use |
|---|---|
| `envs/local.tfvars` | Add-in on localhost:3000, functions via `func start` |
| `envs/dev.tfvars` | Add-in deployed to SWA, functions deployed to Azure |

> The two environments differ only in **where the add-in and Function App run**. The Service Bus namespace and the Storage Account are **always** provisioned in Azure in both cases — there are no local/emulated equivalents, so even a `local` run's `func start` host talks to the same cloud Service Bus and Storage Account. (Both files also share one `name_prefix`, so they target the same resources and state — switching env only changes the live `API_AUDIENCE`/CORS.)

```bash
cd infra

# Local:
cp envs/local.tfvars.example envs/local.tfvars

# Dev (SWA deploy) — fill in SWA placeholders after first apply (see §1.2):
cp envs/dev.tfvars.example envs/dev.tfvars
```

Open the file and fill in every `REPLACE_WITH_*` value:

| Variable | Where to find it |
|---|---|
| `subscription_id` | `az account show --query id -o tsv` |
| `tenant_id` | `az account show --query tenantId -o tsv` |
| `client_id` | App Registration → Application (client) ID |
| `api_audience` | `api://localhost:3000/<client-id>` for local; `api://<swa-hostname>/<client-id>` for dev (fill after first apply) |
| `owner_email` | Your `firstname.lastname@mesh-ai.com` address |
| `location` | Azure region for all resources (e.g. `uksouth`) |
| `name_prefix` | Short prefix, ≤ 12 chars, e.g. `oaddintest` |
| `swa_location` | SWA region — must be one of `westeurope`, `eastus2`, `centralus`, `eastasia`, `westus2` |
| `end_date` | Leave default `170826` unless you need a different end date |

### 1.2 Run Terraform

```bash
cd infra
terraform init

# First apply — use local.tfvars (api_audience and add_in_origin for localhost):
terraform plan -var-file=envs/local.tfvars -out tfplan
terraform apply tfplan
```

After this first apply you'll have the SWA hostname. To switch to the dev environment:
1. Fill in the `REPLACE_WITH_SWA_HOSTNAME` placeholders in `envs/dev.tfvars`
2. Re-apply: `terraform plan -var-file=envs/dev.tfvars -out tfplan && terraform apply tfplan`

Terraform creates (in dependency order): resource group → Service Bus namespace + queue → storage account → Log Analytics → Application Insights → App Service Plan → Function App + Managed Identity → Service Bus role assignments → Static Web App.

### 1.3 Record all outputs — you need these in every step that follows

```bash
terraform output
# Sensitive outputs require the -raw flag:
terraform output -raw addin_deploy_token   # SWA deployment token (treat as secret)
```

Save these values somewhere accessible (a password manager or notes file — NOT in the repo):

```
function_app_hostname  = https://<name>-func.azurewebsites.net
function_app_name      = <name>-func
service_bus_fqdn       = <name>-sbns.servicebus.windows.net
service_bus_namespace_id = /subscriptions/.../servicebus/namespaces/<name>-sbns
addin_url              = https://<random-slug>.azurestaticapps.net
addin_deploy_token     = <token>   ← sensitive, keep private
```

---

## § 2 — Configure the App Registration for the dev deployment (Entra ID)

The SWA hostname is now known. You must update the App Registration before the deployed add-in can authenticate.

### 2.1 Add the SWA URL as a second App ID URI

The `api://` URI must match the domain the add-in is served from. The local URI (`api://localhost:3000/<client-id>`) stays valid alongside the dev one.

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → Expose an API:**

1. Click **Add** next to the Application ID URI field (or edit the existing one).
2. Add a second URI: `api://<swa-hostname>/<client-id>`
   - `<swa-hostname>` is the full hostname from `terraform output addin_url`, without `https://`
   - Example: `api://proud-pond-0123456789.azurestaticapps.net/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
3. Save.

> The two App ID URIs co-exist. Local uses the `localhost:3000` one; dev deployment uses the SWA one.

### 2.2 Add dev redirect URIs

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → Authentication → Single-page application:**

Add these two URIs (keep the existing `brk-multihub://localhost:3000` entry):

```
https://<swa-hostname>/commands.html
brk-multihub://<swa-hostname>
```

- The first is the SPA redirect for standard MSAL flows.
- The second is the **broker redirect** required for NAA (`createNestablePublicClientApplication`) in Outlook. It must be type **Single-page application**, use the `brk-multihub://` scheme, and contain **only the origin** — the bare hostname, no `https://` and no path. Example: `brk-multihub://proud-pond-0123456789.azurestaticapps.net`.

> **If you skip the `brk-multihub://<swa-hostname>` entry**, the deployed add-in fails token acquisition with `AADSTS700046: Invalid Reply Address … must have scheme brk-<broker-id>:// and be of Single Page Application type`. The error names a *specific* broker client ID, but the `brk-multihub://` group already covers Outlook (plus Word/Excel/PowerPoint/Teams) — registering `brk-multihub://<your-domain>` is what resolves it. Do **not** register a broker-specific `brk-<client-id>://…` URI or append a path like `/auth`; NAA redirect URIs are origin-only. See [Microsoft's NAA guide](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#add-a-trusted-broker-through-spa-redirect).

### 2.3 Update dev.tfvars with the SWA audience

Now that you know the SWA URL, fill the placeholder in `infra/envs/dev.tfvars`:

```hcl
api_audience  = "api://<swa-hostname>/<client-id>"
add_in_origin = "https://<swa-hostname>"
```

Re-apply Terraform so the Function App picks up the new `API_AUDIENCE` app setting:

```bash
cd infra
terraform plan -var-file=envs/dev.tfvars -out tfplan
terraform apply tfplan
```

### 2.4 Grant admin consent again (if needed)

If you added new scopes or changed the App ID URI, re-grant admin consent:

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → API permissions → Grant admin consent for `<your-tenant>`**

---

## § 3 — Deploy Azure Functions

### 3.1 Install dependencies and export requirements

```bash
cd src/functions
uv sync                                     # creates/updates .venv
uv export --no-dev --no-hashes -o requirements.txt  # Oryx remote build installs from this
```

> `requirements.txt` is the file Azure's Oryx build reads at publish time. Always regenerate it from `pyproject.toml` before publishing if you changed any dependencies.

### 3.2 Publish to Azure

```bash
func azure functionapp publish <function_app_name>
# e.g.: func azure functionapp publish oaddintest-func
```

The CLI packages your code, uploads it, and triggers an Oryx remote build that installs the Python dependencies from `requirements.txt`. This takes 1–3 minutes.

You should see output ending with:
```
Syncing triggers...
Functions in <name>-func:
    submission_prepare - [httpTrigger]
        Invoke url: https://<name>-func.azurewebsites.net/api/submissions/prepare
    submission_receiver - [httpTrigger]
        Invoke url: https://<name>-func.azurewebsites.net/api/submissions
    download_generator - [httpTrigger]
        Invoke url: https://<name>-func.azurewebsites.net/api/downloads
    health - [httpTrigger]
        Invoke url: https://<name>-func.azurewebsites.net/api/health
```

### 3.3 Verify the deployment

Check the health endpoint responds:

```bash
curl https://<function_app_name>.azurewebsites.net/api/health
# Expected: {"status": "ok"}
```

If this returns a 500 or times out, check the deployment logs:

```bash
func azure functionapp logstream <function_app_name>
```

Common causes of startup failures:
- **Missing app settings** — run `az functionapp config appsettings list --name <name> --resource-group <rg>` and check all expected keys are present.
- **Import error on startup** — the `shared/` package failed to load. Check `requirements.txt` was exported and includes all deps.
- **Wrong Python version** — `main.tf` sets `python_version = "3.13"`. Confirm the Functions runtime supports it; otherwise change to `"3.11"`.

### 3.4 Confirm Service Bus connectivity

The Function App authenticates to Service Bus using a connection string stored in app settings. Terraform provisions this automatically. Verify the setting is present:

```bash
az functionapp config appsettings list \
  --name <function_app_name> --resource-group <rg_name> \
  --query "[?name=='ServiceBusConnection']"
```

If the `ServiceBusConnection` setting is missing, re-run `terraform apply`.

---

## § 4 — Deploy the add-in to Static Web Apps

### 4.1 Generate placeholder icons (first time only)

```bash
cd <repo-root>
node scripts/create-icons.js
```

Replace the placeholder PNGs in `src/addin/public/assets/` with real branded icons before any wider distribution.

### 4.2 Create the dev environment file

Create `src/addin/.env.dev` (this file is gitignored):

```bash
# src/addin/.env.dev
VITE_CLIENT_ID=<your-client-id>
VITE_TENANT_ID=<your-tenant-id>
VITE_API_BASE_URL=https://<function_app_name>.azurewebsites.net
VITE_API_SCOPE=api://<swa-hostname>/<client-id>/access_as_user
```

Fill in:
- `VITE_CLIENT_ID` and `VITE_TENANT_ID` — from your App Registration
- `VITE_API_BASE_URL` — from `terraform output function_app_hostname`
- `VITE_API_SCOPE` — constructed as `api://<swa-hostname>/<client-id>/access_as_user`
  - `<swa-hostname>` from `terraform output addin_url` (strip the `https://`)
  - Example: `api://proud-pond-0123456789.azurestaticapps.net/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx/access_as_user`

> **Why `access_as_user`?** That is the scope name you defined on the App Registration in `SETUP.md` Step 1a. The full scope URI is `<App-ID-URI>/access_as_user`. MSAL's `ssoSilent` call in `commands.ts` requests this scope, and the Function's JWT validator checks the token audience against the App ID URI.

### 4.3 Install add-in dependencies

```bash
cd src/addin
npm install
```

### 4.4 Build the add-in for dev

```bash
npm run build
# Runs: tsc && vite build --mode dev
# Output goes to src/addin/dist/
```

Vite picks up `.env.dev` because the build script passes `--mode dev`. Confirm the build succeeds and `dist/commands.html` exists.

### 4.5 Deploy to Static Web Apps

Get the deployment token from Terraform and deploy:

```bash
cd src/addin

DEPLOY_TOKEN=$(cd ../../infra && terraform output -raw addin_deploy_token)

npx @azure/static-web-apps-cli deploy dist \
  --deployment-token "$DEPLOY_TOKEN"
# Note: --env production is the SWA slot name (Free tier has one slot); omitting it uses the same default.
```

The SWA CLI uploads the contents of `dist/` and propagates them to the CDN. This takes ~30–60 seconds.

> If `@azure/static-web-apps-cli` isn't installed, the `npx` call installs it temporarily. To install globally: `npm install -g @azure/static-web-apps-cli`.

### 4.6 Verify the SWA deployment

```bash
curl https://<swa-hostname>/commands.html
# Should return the HTML content (200 OK, not a 404 or redirect)
```

Also check that the icon assets are served:

```bash
curl -I https://<swa-hostname>/assets/icon-80.png
# Expect: HTTP/2 200, content-type: image/png
```

---

## § 5 — Update the manifest and re-sideload

### 5.1 Generate the dev manifest

`src/addin/manifest.xml` is the localhost source template. Don't hand-edit hosts — generate the environment-specific manifest with the build script, which substitutes the add-in host into every URL, `<AppDomain>`, and the `WebApplicationInfo` App ID URI:

```bash
cd src/addin
npm run build:manifest:dev     # → src/addin/manifest.dev.xml  (SWA host)
# local equivalent (no substitution; identical to manifest.xml):
npm run build:manifest:local   # → src/addin/manifest.local.xml
# or both at once:
npm run build:manifest
```

The host for each env is read from the matching `src/addin/.env.<env>` (`VITE_API_SCOPE = api://<host>/<client-id>/…`), so `.env.dev` must have the SWA hostname filled in first (§2.3). Override with `node ../../scripts/build-manifest.js dev --host <hostname>` if needed.

> Both manifests keep the same `<Id>`, so sideload **one at a time** — remove the other first. (If you need the local and dev add-ins installed simultaneously, give each a distinct `<Id>`.)

### 5.2 Sideload the dev manifest

Follow `SETUP.md` Step 7:

- **Outlook on the web / new Outlook on Windows**: Settings → search "add-in" → My add-ins → Add a custom add-in → Add from file → upload `src/addin/manifest.dev.xml`.
- If the old version (localhost or a previous deploy) was already sideloaded, remove it first, then re-add.

### 5.3 End-to-end test

1. Open an email in your test mailbox.
2. Click **Sync to Test API** in the ribbon.
3. The notification should show `Submitted — reference: ENQ-XXXXXXXX`.
4. Confirm both functions logged the same `correlationId`:

```bash
func azure functionapp logstream <function_app_name>
# Or in Application Insights:
# Logs → KQL:
# traces
# | where message contains "correlationId"
# | order by timestamp desc
```

You should see `enquiry_received` and `enquiry_processing` events with the same `correlationId`. That is the proof the event-driven loop works end to end.

---

## § 6 — Subsequent deployments (re-deploy workflow)

### Re-deploy the function only

```bash
cd src/functions
uv export --no-dev --no-hashes -o requirements.txt
func azure functionapp publish <function_app_name>
```

### Re-deploy the add-in only

```bash
cd src/addin
npm run build
DEPLOY_TOKEN=$(cd ../../infra && terraform output -raw addin_deploy_token)
npx @azure/static-web-apps-cli deploy dist --deployment-token "$DEPLOY_TOKEN"
```

No manifest re-sideload needed unless you changed `manifest.xml`.

### Re-apply Terraform (infra changes only)

```bash
cd infra
terraform plan -var-file=envs/local.tfvars -out tfplan   # or dev.tfvars
terraform apply tfplan
```

---

## § 7 — Troubleshooting

### "Server rejected token (401)"

The Function App's `API_AUDIENCE` setting doesn't match what the add-in is requesting.

Check:
1. `terraform output` — what `api_audience` was Terraform given? Run `terraform apply` after updating `terraform.tfvars` `api_audience`.
2. `VITE_API_SCOPE` in `src/addin/.env.dev` — the scope must be `<api_audience>/access_as_user`.
3. The App Registration's App ID URI matches both.

```bash
az functionapp config appsettings list \
  --name <function_app_name> --resource-group <rg_name> \
  --query "[?name=='API_AUDIENCE']"
```

### "Auth failed: NAA token timed out"

The MSAL NAA broker didn't respond within 10 seconds.

Check:
1. The App Registration has the correct **Authorized client applications** (the 6 Office host GUIDs from `SETUP.md` Step 1b).
2. Admin consent is granted for the `access_as_user` scope (`SETUP.md` Step 1c).
3. The SPA redirect URI `brk-multihub://<swa-hostname>` is registered (§2.2 above).
4. The `VITE_API_SCOPE` in the built add-in matches the App ID URI registered on the App Registration.

### "Sign-in required — open the add-in task pane once"

This is `InteractionRequiredAuthError` from MSAL. The user needs to grant interactive consent — but this add-in has no task pane. Resolution: re-grant admin consent from the portal (§2.4) so the user never sees an interactive prompt.

### Function App health endpoint returns 500

```bash
func azure functionapp logstream <function_app_name>
```

Common causes: missing app settings (re-apply Terraform), import error in `shared/` (check `requirements.txt` was exported correctly), Python version mismatch.

### SWA returns 404 for /commands.html

The `dist/` directory didn't contain `commands.html` when you deployed. Confirm `npm run build` completed without errors and `src/addin/dist/commands.html` exists before running `swa deploy`.

### Service Bus messages not being processed

```bash
az servicebus queue show \
  --name submission-queue \
  --namespace-name <namespace-name> \
  --resource-group <rg_name> \
  --query "{active:countDetails.activeMessageCount, dlq:countDetails.deadLetterMessageCount}"
```

If `deadLetterMessageCount` is climbing, check the function logs for processing errors. If `activeMessageCount` is climbing, the trigger binding isn't connecting — confirm the `ServiceBusConnection` app setting is present and correct (§3.4).

---

## § 8 — Tear down

```bash
cd infra
terraform destroy
```

Then delete the `outlook-addin-test` App Registration manually in Entra ID (not managed by Terraform). Remember to do this before `_end_date` (`170826` = 17 Aug 2026) to stay within the tagging policy commitment.
