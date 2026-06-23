# Runbook — Outlook Add-in Test Harness

Covers everything from first-time setup through ongoing deployments and troubleshooting. Work through in order on first run; jump to the relevant section for subsequent operations.

> **Assumption:** you have access to an Azure subscription and a Microsoft 365 tenant you control (or a [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) sandbox tenant with test mailboxes).

---

## Contents

- [Prerequisites](#prerequisites)
- [§ 1 — App Registration (Entra ID)](#-1--app-registration-entra-id)
- [§ 2 — Provision infrastructure](#-2--provision-infrastructure)
- [§ 3 — Configure local environment](#-3--configure-local-environment)
- [§ 4 — Run locally](#-4--run-locally)
- [§ 5 — Deploy to Azure](#-5--deploy-to-azure)
- [§ 6 — Re-deploy workflow](#-6--re-deploy-workflow)
- [§ 7 — Troubleshooting](#-7--troubleshooting)
- [§ 8 — Tear down](#-8--tear-down)

---

## Prerequisites

Install the following tools before doing anything else. These are one-time installs on your machine.

**Azure CLI**

```bash
# macOS
brew install azure-cli
az login
```

**Terraform ≥ 1.6**

```bash
brew tap hashicorp/tap && brew install hashicorp/tap/terraform
```

**Azure Functions Core Tools v4**

```bash
npm i -g azure-functions-core-tools@4
```

**Node.js v20+ and npm**

Used for the add-in (TypeScript + Vite). Use [nvm](https://github.com/nvm-sh/nvm) or install directly from nodejs.org.

**uv + Python 3.13**

uv manages the Python version and virtual environment for the functions.

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Python 3.13 is declared in the functions project; uv will install it automatically on `uv sync`.

**Office add-in dev certs (one-time)**

Office add-ins require HTTPS even when running locally.

```bash
npx office-addin-dev-certs install
```

---

## § 1 — App Registration (Entra ID)

This is the only step that can't be automated. The App Registration provides the identity for NAA SSO (Office.js `getAccessToken`) and for token validation in the functions (PyJWT).

### 1.1 Register the application

**Azure Portal → Microsoft Entra ID → App registrations → New registration**

- **Name**: `outlook-addin-test`
- **Supported account types**: "Accounts in this organizational directory only" (single-tenant)
- **Redirect URI**: leave blank
- Click **Register**

Copy the **Application (client) ID** and **Directory (tenant) ID** — you will need both throughout the rest of this guide.

### 1.2 Expose an API

**Expose an API → Application ID URI → Set**

For NAA SSO, the URI format is `api://<host>/<client-id>`, where `<host>` must match the domain your add-in files are served from.

- **Local dev**: `api://localhost:3000/<client-id>`
- **Deployed (SWA)**: `api://<swa-hostname>/<client-id>` — add this as a second URI after your first Terraform apply and SWA deployment

**Add a scope:**

| Field | Value |
|---|---|
| Scope name | `access_as_user` |
| Who can consent | Admins and users |
| Admin consent display name | Access the add-in as the signed-in user |
| Admin consent description | Allows Office to call the add-in API on behalf of the signed-in user |
| State | Enabled |

Click **Add scope**.

### 1.3 API permissions

**API permissions**

- The default `Microsoft Graph → User.Read` (delegated) is correct — leave it.
- Click **Add a permission → My APIs → outlook-addin-test → Delegated permissions → access_as_user → Add**.
- Click **Grant admin consent for \<your tenant\>** and confirm.

### 1.4 SPA redirect URIs (NAA broker)

**Authentication → Add a platform → Single-page application**

Add the following redirect URIs:

| URI | When |
|---|---|
| `brk-multihub://localhost:3000` | Local dev |
| `brk-multihub://<swa-hostname>` | After SWA deploy |
| `https://<swa-hostname>/commands.html` | After SWA deploy |

The `brk-multihub://` entries are the NAA broker redirects required by `createNestablePublicClientApplication`. They must be origin-only — no `https://`, no path. Without them, token acquisition fails with `AADSTS700046`.

See [§ 5.1](#51-add-the-swa-url-as-a-second-app-id-uri) for the SWA redirect URIs once the hostname is known.

### 1.5 Values to record

You will use these in every configuration file that follows:

| Variable | Where to find it |
|---|---|
| `TENANT_ID` | Directory (tenant) ID on the App Registration overview |
| `CLIENT_ID` | Application (client) ID on the App Registration overview |
| `API_AUDIENCE` | The Application ID URI you set in step 1.2, e.g. `api://localhost:3000/<client-id>` |

---

## § 2 — Provision infrastructure

### 2.1 Fill in the env-specific tfvars

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

# Dev (SWA deploy) — fill in SWA placeholders after first apply (see §2.2):
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
| `end_date` | Default `170826` (17 Aug 2026, ddmmyy format) |

Both tfvars files are gitignored — never commit them.

> **Indicium mandatory tagging policy.** Every resource carries six `_`-prefixed tags. `_purpose = Testing` and `_business_criticality = Low` are hardcoded. The remaining four (`_owner_email`, `_end_date`, `_project`, `_description`) come from tfvars. Terraform plan fails immediately if `owner_email` is not `@mesh-ai.com` or `end_date` is not in ddmmyy format. The `end_date` represents the date you committed to under the tagging policy — either tear down by then or bump it in tfvars and re-apply (within the 2-month Testing cap).

### 2.2 Run Terraform

```bash
cd infra
terraform init

# First apply — use local.tfvars (api_audience for localhost):
terraform plan -var-file=envs/local.tfvars -out tfplan
terraform apply tfplan
```

After this first apply you'll have the SWA hostname. To switch to the dev environment:
1. Fill in the `REPLACE_WITH_SWA_HOSTNAME` placeholders in `envs/dev.tfvars`
2. Re-apply: `terraform plan -var-file=envs/dev.tfvars -out tfplan && terraform apply tfplan`

Terraform creates (in dependency order): resource group → Service Bus namespace + queue → storage account → Log Analytics → Application Insights → App Service Plan → Function App + Managed Identity → Static Web App.

### 2.3 Record all outputs — you need these in every step that follows

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

## § 3 — Configure local environment

### 3.1 Configure the functions

```bash
cd src/functions
cp local.settings.json.example local.settings.json
```

Fill `local.settings.json` using the values from `terraform output`:

| Setting | Value |
|---|---|
| `AzureWebJobsStorage` | `"UseDevelopmentStorage=true"` when using Azurite locally, or the real connection string from `terraform output storage_account_key` |
| `TENANT_ID` | From App Registration |
| `API_AUDIENCE` | `api://localhost:3000/<client-id>` |
| `ALLOWED_CORS_ORIGIN` | `https://localhost:3000` |
| `STORAGE_ACCOUNT_NAME` | From `terraform output storage_account_name` |
| `STORAGE_ACCOUNT_KEY` | From `terraform output storage_account_key` |
| `BLOB_CONTAINER_NAME` | From Terraform output |
| `SERVICEBUS_CONNECTION_STRING` | From `terraform output service_bus_connection_string` |
| `ServiceBusConnection` | Same value as `SERVICEBUS_CONNECTION_STRING` |
| `SERVICEBUS_FQDN` | From `terraform output service_bus_fqdn` |
| `SERVICEBUS_QUEUE_NAME` | `submission-queue` |

`local.settings.json` is gitignored — never commit it.

Install dependencies:

```bash
uv sync
```

uv will install Python 3.13 (if not present) and create the virtual environment.

### 3.2 Configure the add-in

```bash
cd src/addin
npm install
node ../../scripts/create-icons.js   # generates placeholder icons (one-time)
```

The dev cert install from the Prerequisites step is sufficient — no need to run it again here.

Create `src/addin/.env.local`:

```
VITE_CLIENT_ID=<your-client-id>
VITE_TENANT_ID=<your-tenant-id>
VITE_API_BASE_URL=http://localhost:7071
VITE_API_SCOPE=api://localhost:3000/<client-id>/access_as_user
```

`.env.local` is gitignored — never commit it.

### 3.3 Build and sideload the local manifest

```bash
cd src/addin
npm run build:manifest:local   # generates src/addin/manifest.local.xml
```

Sideload into Outlook:

**Outlook on the web or new Outlook on Windows:** Settings → search "add-in" → My add-ins → Add a custom add-in → Add from file → upload `manifest.local.xml`.

The manifest needs to be re-sideloaded whenever its content changes.

---

## § 4 — Run locally

Start the functions host:

```bash
cd src/functions
uv run func start
```

Start the add-in dev server:

```bash
cd src/addin
npm run dev
```

The add-in is served at `https://localhost:3000`. The functions host listens at `http://localhost:7071`.

**API endpoints reference**

All functions use anonymous auth at the platform level; tokens are validated in code.

| Function | Method | Path |
|---|---|---|
| `health` | GET | `/api/health` |
| `submission_prepare` | POST | `/api/submissions/prepare` |
| `submission_receiver` | POST | `/api/submissions` |
| `download_generator` | POST | `/api/downloads` |

Queue name: `submission-queue`

---

## § 5 — Deploy to Azure

Work through these subsections in order on first deploy. On subsequent deploys, jump to the relevant part (§5.3 for function changes, §5.4 for add-in changes).

### 5.1 Add the SWA URL as a second App ID URI

The SWA hostname is now known from `terraform output addin_url`. You must update the App Registration before the deployed add-in can authenticate.

The `api://` URI must match the domain the add-in is served from. The local URI (`api://localhost:3000/<client-id>`) stays valid alongside the dev one.

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → Expose an API:**

1. Click **Add** next to the Application ID URI field (or edit the existing one).
2. Add a second URI: `api://<swa-hostname>/<client-id>`
   - `<swa-hostname>` is the full hostname from `terraform output addin_url`, without `https://`
   - Example: `api://proud-pond-0123456789.azurestaticapps.net/xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
3. Save.

> The two App ID URIs co-exist. Local uses the `localhost:3000` one; dev deployment uses the SWA one.

### 5.2 Add dev redirect URIs

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → Authentication → Single-page application:**

Add these two URIs (keep the existing `brk-multihub://localhost:3000` entry):

```
https://<swa-hostname>/commands.html
brk-multihub://<swa-hostname>
```

- The first is the SPA redirect for standard MSAL flows.
- The second is the **broker redirect** required for NAA (`createNestablePublicClientApplication`) in Outlook. It must be type **Single-page application**, use the `brk-multihub://` scheme, and contain **only the origin** — the bare hostname, no `https://` and no path. Example: `brk-multihub://proud-pond-0123456789.azurestaticapps.net`.

> **If you skip the `brk-multihub://<swa-hostname>` entry**, the deployed add-in fails token acquisition with `AADSTS700046: Invalid Reply Address … must have scheme brk-<broker-id>:// and be of Single Page Application type`. The error names a *specific* broker client ID, but the `brk-multihub://` group already covers Outlook (plus Word/Excel/PowerPoint/Teams) — registering `brk-multihub://<your-domain>` is what resolves it. Do **not** register a broker-specific `brk-<client-id>://…` URI or append a path like `/auth`; NAA redirect URIs are origin-only. See [Microsoft's NAA guide](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/enable-nested-app-authentication-in-your-add-in#add-a-trusted-broker-through-spa-redirect).

### 5.3 Deploy Azure Functions

Install dependencies and export requirements:

```bash
cd src/functions
uv sync                                     # creates/updates .venv
uv export --no-dev --no-hashes -o requirements.txt  # Oryx remote build installs from this
```

> `requirements.txt` is the file Azure's Oryx build reads at publish time. Always regenerate it from `pyproject.toml` before publishing if you changed any dependencies.

Update `dev.tfvars` with the SWA audience and re-apply so the Function App picks up the new `API_AUDIENCE` app setting:

```bash
# infra/envs/dev.tfvars:
#   api_audience = "api://<swa-hostname>/<client-id>"

cd infra
terraform plan -var-file=envs/dev.tfvars -out tfplan
terraform apply tfplan
```

If you added new scopes or changed the App ID URI, re-grant admin consent:

**Azure Portal → Entra ID → App registrations → `outlook-addin-test` → API permissions → Grant admin consent for `<your-tenant>`**

Publish the functions:

```bash
cd src/functions
func azure functionapp publish <function_app_name> --python
# e.g.: func azure functionapp publish oaddintest-func --python
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

Verify the health endpoint responds:

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

Confirm Service Bus connectivity — Terraform provisions this automatically, but verify the setting is present:

```bash
az functionapp config appsettings list \
  --name <function_app_name> --resource-group <rg_name> \
  --query "[?name=='ServiceBusConnection']"
```

If the `ServiceBusConnection` setting is missing, re-run `terraform apply`.

### 5.4 Deploy the add-in to Static Web Apps

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

> **Why `access_as_user`?** That is the scope name you defined on the App Registration. The full scope URI is `<App-ID-URI>/access_as_user`. MSAL's `ssoSilent` call in `commands.ts` requests this scope, and the Function's JWT validator checks the token audience against the App ID URI.

Install dependencies and build:

```bash
cd src/addin
npm install
npm run build
# Runs: tsc && vite build --mode dev
# Output goes to src/addin/dist/
```

Vite picks up `.env.dev` because the build script passes `--mode dev`. Confirm the build succeeds and `dist/commands.html` exists.

Deploy to Static Web Apps:

```bash
cd src/addin

DEPLOY_TOKEN=$(cd ../../infra && terraform output -raw addin_deploy_token)

npx @azure/static-web-apps-cli deploy dist \
  --deployment-token "$DEPLOY_TOKEN"
```

The SWA CLI uploads the contents of `dist/` and propagates them to the CDN. This takes ~30–60 seconds.

> If `@azure/static-web-apps-cli` isn't installed, the `npx` call installs it temporarily. To install globally: `npm install -g @azure/static-web-apps-cli`.

Verify the SWA deployment:

```bash
curl https://<swa-hostname>/commands.html
# Should return the HTML content (200 OK, not a 404 or redirect)

curl -I https://<swa-hostname>/assets/icon-80.png
# Expect: HTTP/2 200, content-type: image/png
```

### 5.5 Generate the dev manifest and re-sideload

`src/addin/manifest.xml` is the localhost source template. Generate the environment-specific manifest with the build script, which substitutes the add-in host into every URL, `<AppDomain>`, and the `WebApplicationInfo` App ID URI:

```bash
cd src/addin
npm run build:manifest:dev     # → src/addin/manifest.dev.xml  (SWA host)
# local equivalent (no substitution; identical to manifest.xml):
npm run build:manifest:local   # → src/addin/manifest.local.xml
# or both at once:
npm run build:manifest
```

The host for each env is read from the matching `src/addin/.env.<env>` (`VITE_API_SCOPE = api://<host>/<client-id>/…`), so `.env.dev` must have the SWA hostname filled in first (§5.4). Override with `node ../../scripts/build-manifest.js dev --host <hostname>` if needed.

> Both manifests keep the same `<Id>`, so sideload **one at a time** — remove the other first. (If you need the local and dev add-ins installed simultaneously, give each a distinct `<Id>`.)

Sideload the dev manifest:

**Outlook on the web or new Outlook on Windows:** Settings → search "add-in" → My add-ins → Add a custom add-in → Add from file → upload `src/addin/manifest.dev.xml`.

If the old version (localhost or a previous deploy) was already sideloaded, remove it first, then re-add.

### 5.6 End-to-end test

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

## § 6 — Re-deploy workflow

### Functions only

```bash
cd src/functions
uv export --no-dev --no-hashes -o requirements.txt
func azure functionapp publish <function_app_name> --python
```

### Add-in only

```bash
cd src/addin
npm run build
DEPLOY_TOKEN=$(cd ../../infra && terraform output -raw addin_deploy_token)
npx @azure/static-web-apps-cli deploy dist --deployment-token "$DEPLOY_TOKEN"
```

No manifest re-sideload needed unless `manifest.xml` changed.

### Infra only

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
1. `terraform output` — what `api_audience` was Terraform given? Run `terraform apply` after updating `api_audience` in your tfvars.
2. `VITE_API_SCOPE` in `src/addin/.env.dev` — must be `<api_audience>/access_as_user`.
3. The App Registration's App ID URI matches both.

```bash
az functionapp config appsettings list \
  --name <function_app_name> --resource-group <rg_name> \
  --query "[?name=='API_AUDIENCE']"
```

### "Auth failed: NAA token timed out"

The MSAL NAA broker didn't respond within 10 seconds.

Check:
1. Admin consent is granted for the `access_as_user` scope.
2. The SPA redirect URI `brk-multihub://<swa-hostname>` is registered (§5.2).
3. `VITE_API_SCOPE` in the built add-in matches the App ID URI on the App Registration.

### "Sign-in required — open the add-in task pane once"

This is `InteractionRequiredAuthError` from MSAL. The user needs to grant interactive consent — but this add-in has no task pane. Resolution: re-grant admin consent from the portal (§5.3) so the user never sees an interactive prompt.

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

If `deadLetterMessageCount` is climbing, check the function logs for processing errors. If `activeMessageCount` is climbing, the trigger binding isn't connecting — confirm the `ServiceBusConnection` app setting is present and correct (§5.3).

---

## § 8 — Tear down

```bash
cd infra
terraform destroy -var-file=envs/local.tfvars
```

Then delete the `outlook-addin-test` App Registration manually in Entra ID (not managed by Terraform). Do this before `_end_date` (`170826` = 17 Aug 2026) to stay within the tagging policy commitment.
