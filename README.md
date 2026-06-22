# Outlook Add-in Test Harness

> **Throwaway test harness** — validates the event-driven enquiry pattern (add-in → Azure Function → Service Bus → Function) end to end. Not production code.

## What it does

1. An underwriter selects an email in Outlook and clicks **"Sync to Test API"** in the ribbon.
2. The add-in silently acquires an Entra ID SSO token and POSTs the email metadata to the `enquiry-receiver` Azure Function.
3. The function validates the JWT, stamps a reference ID, and writes a message to the `enquiry-queue` Service Bus queue.
4. The `enquiry-processor` function picks the message off the queue and logs it — proving the event-driven loop end to end.
5. The add-in shows the user a success notification with the `ENQ-xxxx` reference ID, or a specific error, directly on the email.

There is **no task pane** — the button runs a function directly and feedback is shown via Outlook notification messages.

---

## Repo structure

```
/
├── README.md
├── .env.example          ← all config keys (no secrets)
├── scripts/
│   └── create-icons.js   ← generates placeholder PNG icons for the manifest
├── src/
│   ├── addin/            ← TypeScript + Vite add-in (Office.js function command)
│   │   ├── manifest.xml
│   │   ├── package.json
│   │   ├── vite.config.ts
│   │   ├── tsconfig.json
│   │   ├── commands.html ← function runtime host (never shown to user)
│   │   ├── commands.ts   ← SSO + extract + POST + notify
│   │   └── types.ts      ← EnquiryPayload / EnquiryResponse / EnquiryQueueMessage
│   └── functions/        ← Azure Functions Python v2 model, managed by uv
│       ├── function_app.py          ← FRAMEWORK EXCEPTION: v2 model requires root
│       ├── pyproject.toml
│       ├── requirements.txt         ← generated via uv export; needed by Oryx
│       ├── host.json
│       ├── local.settings.json.example
│       └── shared/
│           ├── auth.py              ← JWT validation (PyJWT)
│           ├── service_bus.py       ← ServiceBusClient + DefaultAzureCredential
│           └── models.py            ← payload TypedDicts
└── infra/                ← Terraform (stays flat, no /src)
    ├── main.tf
    ├── variables.tf
    ├── outputs.tf
    ├── providers.tf
    ├── terraform.tfvars.example
    └── README.md
```

**Framework exception:** `function_app.py` must live at `src/functions/` root — the Azure Functions v2 model host discovers functions via `function_app.py` at the project root; it cannot be moved into a subfolder.

---

## Prerequisites

Install once on your machine (see `SETUP.md` Step 0 for detail):

| Tool | Version | Purpose |
|---|---|---|
| Node.js LTS | v20+ | Add-in build (Vite / TypeScript) |
| uv | latest | Python env + dep management |
| Python | 3.11 | Azure Functions runtime |
| Azure Functions Core Tools | v4 | Local function host |
| Azure CLI | latest | Auth + deployment |
| Terraform | 1.6+ | Infra provisioning |
| office-addin-dev-certs | (via npx) | HTTPS for local add-in |

---

## Local setup (end to end)

Work through these in order. The full manual steps live in `SETUP.md` and `infra/README.md`.

### 1. Create the App Registration (Entra ID) — do this first
Follow `SETUP.md` Step 1. Record `TENANT_ID`, `CLIENT_ID`, and `API_AUDIENCE` (the App ID URI).

### 2. Provision infra

```bash
cd infra
cp terraform.tfvars.example terraform.tfvars   # fill in your values
terraform init
terraform plan -out tfplan
terraform apply tfplan
terraform output   # note function_app_hostname and service_bus_fqdn
```

See `infra/README.md` for the full annotated order.

### 3. Configure the functions locally

```bash
cd src/functions
cp local.settings.json.example local.settings.json  # fill in your values
uv sync                                              # creates .venv + installs deps
```

Grant your own user the Service Bus roles for local dev (see `SETUP.md` Step 5):
```bash
az role assignment create --assignee <your-upn> \
  --role "Azure Service Bus Data Sender" \
  --scope <service-bus-namespace-id-from-tf-output>

az role assignment create --assignee <your-upn> \
  --role "Azure Service Bus Data Receiver" \
  --scope <service-bus-namespace-id-from-tf-output>
```

> **Local Service Bus:** the functions use a real Azure Service Bus namespace even locally. `DefaultAzureCredential` falls back to your `az login` identity — no connection strings needed.

### 4. Start the function host

```bash
cd src/functions
uv run func start
# Functions available at http://localhost:7071
```

You also need **Azurite** (Azure Storage emulator) for `AzureWebJobsStorage` locally:
```bash
npx azurite --silent --location /tmp/azurite
```
Or replace `UseDevelopmentStorage=true` in `local.settings.json` with a real storage connection string.

### 5. Configure the add-in

```bash
cd src/addin
npm install
```

Create `src/addin/.env.local`:
```
VITE_API_BASE_URL=http://localhost:7071
```

Install dev certs (once):
```bash
npx office-addin-dev-certs install
```

Generate placeholder icons (once):
```bash
node ../../scripts/create-icons.js
```

Edit `src/addin/manifest.xml` — replace the two `REPLACE_WITH_YOUR_CLIENT_ID` placeholders and the `REPLACE-WITH-NEW-GUID` add-in ID.

### 6. Start the add-in dev server

```bash
cd src/addin
npm run dev
# Serves at https://localhost:3000
```

### 7. Sideload the manifest

Follow `SETUP.md` Step 7. In Outlook on the web: Settings → My add-ins → Add from file → upload `src/addin/manifest.xml`.

---

## Deploying to Azure

```bash
# 1. Deploy function code (Terraform already provisioned the Function App)
cd src/functions
uv export --no-hashes -o requirements.txt
func azure functionapp publish <function-app-name-from-tf-output>

# 2. Point the add-in at the deployed function
#    Create src/addin/.env.production.local:
#    VITE_API_BASE_URL=https://<function-app-name>.azurewebsites.net

# 3. Build the add-in
cd src/addin && npm run build

# 4. Update manifest.xml SourceLocation / AppDomains to your deployed add-in host
# 5. Re-sideload the updated manifest (SETUP.md Step 7)
```

Full annotated steps in `SETUP.md` Steps 4–8.

---

## SSO caveat for UI-less commands

`Office.auth.getAccessToken()` can normally prompt the user to sign in. In a **function command** (no task pane), there is nowhere for that dialog to appear. This test harness works because **admin consent is pre-granted** in `SETUP.md` Step 1c — making token acquisition fully silent.

If consent is ever required, `getAccessToken` will return error code 13001 or 13002 and the add-in will show a notification telling the user to open the add-in settings page. The happy path (pre-consented) is always silent.

---

## Decisions applied (from build brief)

- **Cloud-only backend:** the Service Bus namespace and the Storage Account are always real Azure resources, in both the `local` and `dev` environments — there is no local/emulated equivalent, so a `local` run's `func start` host connects to the same cloud resources. (`AzureWebJobsStorage`, the Functions host's *internal* storage, can use Azurite locally — that's separate from the provisioned Storage Account.)
- **Function App plan:** Flex Consumption (`FC1`) — replaced classic Consumption (`Y1`), which got stuck in a persistent post-create `503`. Change `sku_name = "EP1"` in `infra/main.tf` for Premium if needed.

---

## .gitignore gaps to add manually

The existing `.gitignore` does not cover these:

```
# Azure Functions local settings (contains storage connection strings)
src/functions/local.settings.json

# Python virtual environment created by uv
.venv/
```
