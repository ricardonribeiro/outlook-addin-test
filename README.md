# Outlook Add-in Test Harness

Validates an event-driven submission pattern end to end: Outlook add-in → Azure Function → Service Bus → downstream processor. Not production code.

## What it does

Two ribbon buttons appear on any selected email:

- **Sync to Test API** — acquires an Entra ID SSO token silently, POSTs the email metadata and attachments to the `submission_receiver` function, which validates the JWT, writes `payload.json` to blob storage, and enqueues a message to the `submission-queue` Service Bus queue. The add-in shows a success or error notification directly on the email.
- **Download Payload** — POSTs to `download_generator`, which zips the blobs written during the submission step and returns a short-lived read SAS URL. The add-in triggers the download.

There is no task pane. All feedback is delivered via Outlook notification messages.

---

## Architecture overview

The add-in is a TypeScript/Vite function command hosted on an Azure Static Web App. On button click it calls `Office.auth.getAccessToken()` to obtain a bearer token, then calls the relevant Azure Function endpoint over HTTPS.

The four active Azure Functions (Python v2 model, 3.13, Flex Consumption plan) are:

| Function | Method + path | Auth | Purpose |
|---|---|---|---|
| `submission_prepare` | POST `/api/submissions/prepare` | JWT | Validates token; issues write SAS URLs for attachment upload |
| `submission_receiver` | POST `/api/submissions` | JWT | Validates token; writes `payload.json` to blob; enqueues to `submission-queue` |
| `download_generator` | POST `/api/downloads` | JWT | Validates token; zips blobs; returns read SAS URL |
| `health` | GET `/api/health` | Anonymous | Liveness check |

`submission_processor` exists in the codebase as a stub for the downstream consumer but is commented out.

Infrastructure (Terraform): Flex Consumption Function App (FC1), Azure Static Web App (Free tier), Service Bus Standard namespace, Storage Account, Application Insights.

Two environments:

| Environment | Add-in host | Functions host |
|---|---|---|
| `local` | Vite dev server at `https://localhost:3000` | `func start` at `http://localhost:7071` |
| `dev` | Azure Static Web App | Azure Function App |

In both environments the Service Bus namespace and Storage Account are real Azure resources. Only `AzureWebJobsStorage` (the Functions host's internal bookkeeping) can use Azurite locally.

---

## Repo structure

```
/
├── README.md
├── scripts/
│   ├── build-manifest.js        ← generates manifest.local.xml / manifest.dev.xml
│   └── create-icons.js          ← generates placeholder PNG icons
├── src/
│   ├── addin/                   ← TypeScript + Vite add-in
│   │   ├── manifest.xml                ← source template (localhost:3000)
│   │   ├── manifest.local.xml          ← generated; gitignored
│   │   ├── manifest.dev.xml            ← generated; gitignored
│   │   ├── package.json
│   │   ├── vite.config.ts
│   │   ├── tsconfig.json
│   │   ├── commands.html               ← function runtime host (never shown to user)
│   │   ├── commands.ts                 ← SSO + extract + POST + notify
│   │   ├── types.ts
│   │   └── public/
│   │       ├── index.html              ← placeholder page for the SWA root
│   │       └── assets/                 ← icon PNGs
│   └── functions/               ← Azure Functions Python v2, managed by uv
│       ├── function_app.py             ← FRAMEWORK EXCEPTION: must be at project root
│       ├── .funcignore                 ← excludes dev files from Azure deploy
│       ├── pyproject.toml
│       ├── requirements.txt            ← generated via uv export; needed by Oryx
│       ├── host.json
│       ├── local.settings.json.example
│       └── shared/
│           ├── auth.py                 ← JWT validation (PyJWT)
│           ├── blob.py                 ← SAS URL generation + ZIP upload
│           ├── service_bus.py          ← ServiceBusClient + connection string auth
│           └── models.py               ← payload TypedDicts + validation
└── infra/                       ← Terraform (flat, no /src)
    ├── main.tf
    ├── variables.tf
    ├── outputs.tf
    ├── providers.tf
    ├── envs/
    │   ├── local.tfvars.example
    │   ├── local.tfvars                ← gitignored; fill from example
    │   ├── dev.tfvars.example
    │   └── dev.tfvars                  ← gitignored; fill from example
    └── README.md
```

**Framework exception:** `function_app.py` lives at `src/functions/` root because the Azure Functions v2 model host discovers functions via `function_app.py` at the project root — it cannot be in a subfolder.

**Manifests are generated — never edit them by hand.** Run `npm run build:manifest:local` or `npm run build:manifest:dev` to regenerate `manifest.local.xml` / `manifest.dev.xml` from the source template.

---

## Prerequisites

| Tool | Version | Purpose |
|---|---|---|
| Node.js | v20+ | Add-in build (Vite / TypeScript) |
| uv | latest | Python env + dependency management |
| Python | 3.13 | Azure Functions runtime |
| Azure Functions Core Tools | v4 | Local function host |
| Azure CLI | latest | Auth + deployment |
| Terraform | 1.6+ | Infra provisioning |
| office-addin-dev-certs | (via npx) | HTTPS for local add-in |

---

## Quick start — local dev

Full step-by-step instructions, including App Registration setup, are in [`docs/SETUP.md`](docs/SETUP.md).

```bash
# 1. Provision infra (first time only)
cd infra
cp envs/local.tfvars.example envs/local.tfvars   # fill in values
terraform init && terraform plan -var-file=envs/local.tfvars -out tfplan && terraform apply tfplan

# 2. Configure functions
cd src/functions
cp local.settings.json.example local.settings.json  # fill in values from terraform output
uv sync

# 3. Start Azurite (local storage emulator for AzureWebJobsStorage)
npx azurite --silent --location /tmp/azurite

# 4. Start the function host
uv run func start

# 5. Configure and start the add-in
cd src/addin
npm install
# Create src/addin/.env.local with VITE_CLIENT_ID, VITE_TENANT_ID, VITE_API_BASE_URL, VITE_API_SCOPE
npx office-addin-dev-certs install   # once only
node ../../scripts/create-icons.js   # once only
npm run build:manifest:local          # → manifest.local.xml
npm run dev                           # https://localhost:3000

# 6. Sideload manifest.local.xml in Outlook
```

---

## Quick start — deploy to Azure

Full step-by-step instructions are in [`docs/deploy.md`](docs/deploy.md).

```bash
# Functions
cd src/functions
uv export --no-dev --no-hashes -o requirements.txt
func azure functionapp publish <function-app-name-from-tf-output> --python

# Add-in (SWA)
cd src/addin
npm run build   # tsc + vite build --mode dev → dist/
DEPLOY_TOKEN=$(cd ../../infra && terraform output -raw addin_deploy_token)
npx @azure/static-web-apps-cli deploy dist --deployment-token "$DEPLOY_TOKEN"

# Manifest
npm run build:manifest:dev   # → manifest.dev.xml (uses SWA host from .env.dev)
# Sideload manifest.dev.xml in Outlook
```

For Terraform infrastructure reference see [`infra/README.md`](infra/README.md).

---

## SSO caveat

`Office.auth.getAccessToken()` can normally prompt the user to consent. In a function command (no task pane) there is no surface for that dialog. This harness works because admin consent is pre-granted during App Registration setup — making token acquisition fully silent. If consent is ever required, `getAccessToken` returns error code 13001 or 13002 and the add-in surfaces a user-visible error notification.

---

## Key decisions

- **Cloud-only backend.** Service Bus and Storage Account are always real Azure resources in both environments. There is no local emulator for either service; only `AzureWebJobsStorage` (Functions host internal storage) uses Azurite locally.
- **Flex Consumption (FC1).** Classic Consumption (Y1) was replaced after it got stuck in a persistent post-create 503.
- **No task pane.** The ribbon buttons run function commands directly. Feedback is Outlook notification messages only.
- **Silent SSO.** Admin consent is pre-granted so token acquisition never requires user interaction on the happy path.
