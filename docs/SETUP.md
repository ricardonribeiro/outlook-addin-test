# SETUP.md — First-time Setup Guide

The code is ready. This guide covers the one-time steps you need to run it locally: installing prerequisites, creating the Entra ID App Registration (the only step that can't be automated), provisioning infrastructure with Terraform, and configuring each component.

Work through the sections in order.

> **Assumption:** you have access to an Azure subscription and a Microsoft 365 tenant you control (or a [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program) sandbox tenant with test mailboxes).

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

## Step 1 — Create the App Registration (Entra ID)

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

### 1.3 Authorise Office host applications

**Expose an API → Authorized client applications → Add a client application**

Add each of the following GUIDs and tick the `access_as_user` scope. These are the Office hosts that are allowed to request tokens silently — without them `getAccessToken()` fails.

```
d3590ed6-52b3-4102-aeff-aad2292ab01c   Microsoft Office (desktop)
ea5a67f6-b6f3-4338-b240-c655ddc3cc8e   Microsoft Office (alt)
57fb890c-0dab-4253-a5e0-7188c88b2bb4   Office on the web
08e18876-6177-487e-b8b5-cf950c1e598c   Office on the web / SharePoint
bc59ab01-8403-45c6-8796-ac3ef710b3e3   Outlook on the web
93d53678-613d-4013-afc1-62e9e444a0a5   Office on the web (other)
```

### 1.4 API permissions

**API permissions**

- The default `Microsoft Graph → User.Read` (delegated) is correct — leave it.
- Click **Add a permission → My APIs → outlook-addin-test → Delegated permissions → access_as_user → Add**.
- Click **Grant admin consent for \<your tenant\>** and confirm.

### 1.5 SPA redirect URIs (NAA broker)

**Authentication → Add a platform → Single-page application**

Add the following redirect URIs:

- `https://localhost:3000/commands.html`
- `brk-multihub://localhost:3000`

The `brk-multihub://` entry is the NAA broker redirect required by `createNestablePublicClientApplication`. It must be origin-only — no `https://`, no path. Without it, token acquisition fails with `AADSTS700046`.

After deploying to Azure Static Web Apps, add the corresponding SWA URIs here as well:

- `https://<swa-hostname>/commands.html`
- `brk-multihub://<swa-hostname>`

See `docs/deploy.md` for the full deployment walkthrough.

### 1.6 Values to record

You will use these in every configuration file that follows:

| Variable | Where to find it |
|---|---|
| `TENANT_ID` | Directory (tenant) ID on the App Registration overview |
| `CLIENT_ID` | Application (client) ID on the App Registration overview |
| `API_AUDIENCE` | The Application ID URI you set in step 1.2, e.g. `api://localhost:3000/<client-id>` |

---

## Step 2 — Provision infrastructure with Terraform

### 2.1 Fill in tfvars

```bash
cd infra
cp envs/local.tfvars.example envs/local.tfvars
```

Open `envs/local.tfvars` and fill in every value:

| Variable | How to get it |
|---|---|
| `subscription_id` | `az account show --query id -o tsv` |
| `tenant_id` | From App Registration (step 1.6) |
| `client_id` | From App Registration (step 1.6) |
| `api_audience` | `api://localhost:3000/<client-id>` (local) |
| `owner_email` | Your `firstname.lastname@mesh-ai.com` address |
| `location` | Azure region, e.g. `uksouth` |
| `name_prefix` | Short prefix ≤ 12 chars, e.g. `oaddintest` |
| `swa_location` | Must be one of: `westeurope`, `eastus2`, `centralus`, `eastasia`, `westus2` |
| `end_date` | Default `170826` (17 Aug 2026, ddmmyy format) |

`local.tfvars` is gitignored — never commit it.

> **Indicium mandatory tagging policy.** Every resource carries six `_`-prefixed tags. `_purpose = Testing` and `_business_criticality = Low` are hardcoded. The remaining four (`_owner_email`, `_end_date`, `_project`, `_description`) come from tfvars. Terraform plan fails immediately if `owner_email` is not `@mesh-ai.com` or `end_date` is not in ddmmyy format. The `end_date` represents the date you committed to under the tagging policy — either tear down by then or bump it in tfvars and re-apply (within the 2-month Testing cap).

### 2.2 Run Terraform

```bash
cd infra
terraform init
terraform plan -var-file=envs/local.tfvars -out tfplan
terraform apply tfplan
terraform output   # note all output values
```

This provisions: resource group, Storage Account, Service Bus Standard namespace + `submission-queue`, Azure Static Web Apps (Free), Function App on Flex Consumption (FC1) with system-assigned Managed Identity, Log Analytics, and Application Insights.

---

## Step 3 — Configure the functions

### 3.1 Copy and fill local settings

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

### 3.2 Install dependencies

```bash
cd src/functions
uv sync
```

uv will install Python 3.13 (if not present) and create the virtual environment.

---

## Step 4 — Configure the add-in

### 4.1 Install dependencies and run one-time setup

```bash
cd src/addin
npm install
node ../../scripts/create-icons.js   # generates placeholder icons (one-time)
```

The dev cert install from the Prerequisites step is sufficient — no need to run it again here.

### 4.2 Create `.env.local`

Create `src/addin/.env.local` with the following content:

```
VITE_CLIENT_ID=<your-client-id>
VITE_TENANT_ID=<your-tenant-id>
VITE_API_BASE_URL=http://localhost:7071
VITE_API_SCOPE=api://localhost:3000/<client-id>/access_as_user
```

`.env.local` is gitignored — never commit it.

### 4.3 Build and sideload the local manifest

```bash
cd src/addin
npm run build:manifest:local   # generates src/addin/manifest.local.xml
```

Sideload into Outlook:

**Outlook on the web or new Outlook on Windows:** Settings → search "add-in" → My add-ins → Add a custom add-in → Add from file → upload `manifest.local.xml`.

The manifest needs to be re-sideloaded whenever its content changes.

---

## API endpoints reference

All functions use anonymous auth at the platform level; tokens are validated in code.

| Function | Method | Path |
|---|---|---|
| `health` | GET | `/api/health` |
| `submission_prepare` | POST | `/api/submissions/prepare` |
| `submission_receiver` | POST | `/api/submissions` |
| `download_generator` | POST | `/api/downloads` |

Queue name: `submission-queue`

---

## Running locally

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

---

## Tear down

```bash
cd infra
terraform destroy -var-file=envs/local.tfvars
```

Then delete the `outlook-addin-test` App Registration manually in Entra ID — it is not managed by Terraform.

---

## Setup checklist

| # | Task | Where | One-time? |
|---|---|---|---|
| 0 | Install tooling + dev certs | Local machine | Yes |
| 1 | Create and configure App Registration | Entra ID portal | Yes |
| 2 | Fill in `envs/local.tfvars` | Local edit | Yes |
| 3 | `terraform init/plan/apply` | Terraform | Per infra change |
| 4 | Fill in `local.settings.json` | Local edit | Yes |
| 5 | `uv sync` in `src/functions` | Terminal | Per dependency change |
| 6 | Fill in `src/addin/.env.local` | Local edit | Yes |
| 7 | `npm install` + generate icons | Terminal | Yes |
| 8 | `npm run build:manifest:local` + sideload | Terminal + Outlook | Per manifest change |

For full Azure deployment instructions, see `docs/deploy.md`.
